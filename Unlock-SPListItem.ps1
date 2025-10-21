
<#
    .SYNOPSIS
        Unlocks locked record items across multiple SharePoint Online sites and lists.

    .DESCRIPTION
        This script connects to multiple SharePoint Online sites and identifies list items that are locked
        as records (based on _ComplianceFlags values). It then attempts to unlock these items using the
        SharePoint REST API. The script includes comprehensive throttling protection, retry logic, and
        detailed logging capabilities.
        
        The script identifies locked items by examining the _ComplianceFlags field:
        - null/0/771 = unlocked or never set
        - 7 = locked for the first time
        - 519 = relocked after previous unlock
        
        Throttling protection is implemented with configurable delays between operations and exponential
        backoff retry logic for HTTP 429/503 responses.

    .PARAMETER Verbose
        When specified, displays detailed compliance flag information for each item processed.
        This is useful for debugging and understanding why items are or aren't being unlocked.

    .EXAMPLE
        .\Unlock-SPListItem.ps1
        
        Runs the script with default settings, processing all sites listed in SPOSiteList.txt

    .EXAMPLE
        .\Unlock-SPListItem.ps1 -Verbose
        
        Runs the script with verbose output, showing compliance flag details for each item

    .NOTES
        File Name      : Unlock-SPListItem.ps1
        Author         : Mike Lee
        Date           : 10/21/2025
        Prerequisite   : PnP.PowerShell module, Entra App Registration with certificate authentication
        
        Required Permissions:
        - Sites.ReadWrite.All (or Sites.FullControl.All)
        - User.Read.All
        
        Configuration:
        Before running, update the USER CONFIGURATION section with:
        - App ID from your Entra App Registration
        - Tenant ID
        - Certificate thumbprint
        - Path to input file containing site URLs (one per line)
        - Lists to ignore (default: 'Site Assets', 'Site Pages')
        - Throttling delay values if needed
        
        Throttling Protection:
        - DelayBetweenItems: 100ms (default)
        - DelayBetweenLists: 500ms (default)
        - DelayBetweenSites: 1000ms (default)
        - MaxRetryAttempts: 5
        - BaseRetryDelay: 5000ms with exponential backoff
        
        Logging:
        All operations are logged to a timestamped file in %TEMP% directory

    .LINK
        https://docs.microsoft.com/en-us/sharepoint/dev/apis/sharepoint-rest-api
        
    .LINK
        https://pnp.github.io/powershell/

    .OUTPUTS
        Log file in %TEMP% directory named: Unlock_Record_Items_[timestamp]_logfile.log
        Console output showing progress and results
    #>

# =================================================================================================
# USER CONFIGURATION - Update the variables in this section
# =================================================================================================

# --- Script Parameters ---
param(
    [switch]$Verbose = $false  # Add -Verbose to see detailed compliance flag information
)

# --- Tenant and App Registration Details ---
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"                 # This is your Entra App ID
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"                # This is your Tenant ID
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"        # This is certificate thumbprint

# --- Input File Path ---
$sitelist = 'C:\temp\SPOSiteList.txt' # Path to the input file containing site URLs

# --- Lists to Ignore ---
$ignoreListNames = @('Site Assets', 'Site Pages') # Lists to skip during processing

# --- Logging Configuration ---
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss" # Current date and time for unique file naming
$script:LogFilePath = "$env:TEMP\" + 'Unlock_Record_Items_' + $date + '_' + "logfile.log" # Define log file path
$script:EnableLogging = $true

# --- Throttling Protection Settings ---
$DelayBetweenItems = 100        # Milliseconds delay between processing each item (default: 100ms)
$DelayBetweenLists = 500        # Milliseconds delay between processing each list (default: 500ms)
$DelayBetweenSites = 1000       # Milliseconds delay between processing each site (default: 1000ms)
$MaxRetryAttempts = 5           # Maximum number of retry attempts for throttled requests
$BaseRetryDelay = 5000          # Base delay in milliseconds for exponential backoff (default: 5 seconds)

# =================================================================================================
# END OF USER CONFIGURATION
# =================================================================================================

#region LOGGING FUNCTIONS
# =====================================================================================
# Logging functionality to capture console output to file
# =====================================================================================

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'VERBOSE', 'SUCCESS')]
        [string]$Level = 'INFO',
        
        [Parameter()]
        [string]$LogFile = $script:LogFilePath
    )
    
    if (-not $script:EnableLogging -or [string]::IsNullOrEmpty($LogFile)) {
        return
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    try {
        Add-Content -Path $LogFile -Value $logEntry -Encoding UTF8
    }
    catch {
        # If logging fails, don't break the script
        Write-Warning "Failed to write to log file: $($_.Exception.Message)"
    }
}

function Write-VerboseAndLog {
    param([string]$Message)
    Write-Verbose $Message
    Write-Log -Message $Message -Level 'VERBOSE'
}

function Write-WarningAndLog {
    param([string]$Message)
    Write-Warning $Message
    Write-Log -Message $Message -Level 'WARNING'
}

function Write-HostAndLog {
    param(
        [string]$Object,
        [string]$ForegroundColor,
        [switch]$NoNewline
    )
    
    if ($ForegroundColor) {
        if ($NoNewline) {
            Write-Host $Object -ForegroundColor $ForegroundColor -NoNewline
        }
        else {
            Write-Host $Object -ForegroundColor $ForegroundColor
        }
    }
    else {
        if ($NoNewline) {
            Write-Host $Object -NoNewline
        }
        else {
            Write-Host $Object
        }
    }
    
    # Only log if the message is not empty
    if (-not [string]::IsNullOrWhiteSpace($Object)) {
        # Determine log level based on color
        $logLevel = switch ($ForegroundColor) {
            'Red' { 'WARNING' }  # Red items are findings, not errors (locked items found)
            'Yellow' { 'WARNING' }
            'Green' { 'SUCCESS' }
            default { 'INFO' }
        }
        
        Write-Log -Message $Object -Level $logLevel
    }
}

#endregion LOGGING FUNCTIONS

#region THROTTLING FUNCTIONS
# =====================================================================================
# Throttling protection functionality to handle SharePoint Online throttling
# =====================================================================================

function Invoke-WithThrottlingRetry {
    <#
    .SYNOPSIS
    Executes a script block with automatic retry logic for throttling (429/503 responses)
    
    .DESCRIPTION
    This function implements the Microsoft recommended approach for handling SharePoint
    throttling by honoring Retry-After headers and using exponential backoff.
    
    .PARAMETER ScriptBlock
    The script block to execute
    
    .PARAMETER MaxRetries
    Maximum number of retry attempts (default: from configuration)
    
    .PARAMETER Description
    Description of the operation being performed (for logging)
    #>
    param(
        [Parameter(Mandatory)]
        [ScriptBlock]$ScriptBlock,
        
        [Parameter()]
        [int]$MaxRetries = $script:MaxRetryAttempts,
        
        [Parameter()]
        [string]$Description = "Operation"
    )
    
    $attempt = 0
    $success = $false
    $result = $null
    
    while (-not $success -and $attempt -lt $MaxRetries) {
        $attempt++
        
        try {
            $result = & $ScriptBlock
            $success = $true
        }
        catch {
            $statusCode = $null
            $retryAfter = $null
            
            # Try to extract status code from different exception types
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
                
                # Try to get Retry-After header
                if ($_.Exception.Response.Headers) {
                    $retryAfterHeader = $_.Exception.Response.Headers["Retry-After"]
                    if ($retryAfterHeader) {
                        $retryAfter = [int]$retryAfterHeader
                    }
                }
            }
            elseif ($_.Exception.Message -match "429|503") {
                # Try to extract from error message
                if ($_.Exception.Message -match "429") {
                    $statusCode = 429
                }
                elseif ($_.Exception.Message -match "503") {
                    $statusCode = 503
                }
            }
            
            # Check if this is a throttling error (429 or 503)
            if ($statusCode -eq 429 -or $statusCode -eq 503) {
                if ($attempt -lt $MaxRetries) {
                    # Calculate delay: use Retry-After if available, otherwise exponential backoff
                    $delaySeconds = if ($retryAfter) {
                        $retryAfter
                    }
                    else {
                        # Exponential backoff: BaseDelay * 2^(attempt-1)
                        ($script:BaseRetryDelay / 1000) * [Math]::Pow(2, $attempt - 1)
                    }
                    
                    Write-WarningAndLog "$Description - Throttled (HTTP $statusCode). Waiting $delaySeconds seconds before retry $attempt of $MaxRetries..."
                    Start-Sleep -Seconds $delaySeconds
                }
                else {
                    Write-HostAndLog "✗ $Description - Max retry attempts reached after throttling" -ForegroundColor Red
                    throw
                }
            }
            else {
                # Non-throttling error, don't retry
                throw
            }
        }
    }
    
    return $result
}

function Start-ThrottleDelay {
    <#
    .SYNOPSIS
    Introduces a delay to avoid overwhelming SharePoint with requests
    
    .PARAMETER DelayMilliseconds
    The delay in milliseconds
    
    .PARAMETER Description
    Description of what we're delaying (for verbose logging)
    #>
    param(
        [Parameter(Mandatory)]
        [int]$DelayMilliseconds,
        
        [Parameter()]
        [string]$Description = "throttle protection"
    )
    
    if ($DelayMilliseconds -gt 0) {
        Write-VerboseAndLog "Applying $DelayMilliseconds ms delay for $Description"
        Start-Sleep -Milliseconds $DelayMilliseconds
    }
}

#endregion THROTTLING FUNCTIONS

# Function to check if an item is locked as a record
function Test-ItemIsLocked {
    param(
        [string]$SiteUrl,
        [string]$ListTitle,
        [int]$ItemId,
        [string]$ItemTitle
    )
    
    try {
        $item = Invoke-WithThrottlingRetry -Description "Check lock status for item $ItemId" -ScriptBlock {
            Get-PnPListItem -List $ListTitle -Id $ItemId -ErrorAction Stop
        }
        
        # Check the _ComplianceFlags value to determine lock status
        $complianceFlags = $item.FieldValues["_ComplianceFlags"]
        
        # Create a display name for the item
        $itemDisplayName = if ($ItemTitle) { 
            "ID $ItemId ('$ItemTitle')" 
        }
        else { 
            "ID $ItemId" 
        }
        
        # Log the compliance flags for debugging
        if ($Verbose) {
            Write-HostAndLog "      → Item $itemDisplayName - _ComplianceFlags: $complianceFlags" -ForegroundColor Cyan
        }
        
        # Based on your pattern:
        # null = unlocked and never set
        # 7 = locked for the first time
        # 771 = has been unlocked
        # 519 = has been relocked
        
        if ($complianceFlags -eq 7 -or $complianceFlags -eq 519) {
            Write-HostAndLog "      → Item $itemDisplayName is LOCKED (_ComplianceFlags: $complianceFlags)" -ForegroundColor Yellow
            return $true
        }
        elseif ($null -eq $complianceFlags -or $complianceFlags -eq 771 -or $complianceFlags -eq "" -or $complianceFlags -eq 0) {
            # Only log to file, don't display in console for items that don't need unlocking
            # Uncomment the line below for debugging:
            # Write-HostAndLog "      → Item $itemDisplayName has no compliance label or is already unlocked (_ComplianceFlags: $complianceFlags)" -ForegroundColor Gray
            Write-VerboseAndLog "Item $itemDisplayName has no compliance label or is already unlocked (_ComplianceFlags: $complianceFlags)"
            return $false
        }
        else {
            Write-HostAndLog "      → Item $itemDisplayName has unknown compliance flag value: $complianceFlags" -ForegroundColor Yellow
            return $false
        }
    }
    catch {
        $itemDisplayName = if ($ItemTitle) { 
            "ID $ItemId ('$ItemTitle')" 
        }
        else { 
            "ID $ItemId" 
        }
        Write-WarningAndLog "Could not check lock status for item $itemDisplayName in list '$ListTitle': $($_.Exception.Message)"
        return $false
    }
}

# Function to unlock a record item
function Unlock-RecordItem {
    param(
        [string]$SiteUrl,
        [string]$ListTitle,
        [int]$ItemId,
        [string]$ItemTitle
    )
    
    # Create a display name for the item
    $itemDisplayName = if ($ItemTitle) { 
        "ID $ItemId ('$ItemTitle')" 
    }
    else { 
        "ID $ItemId" 
    }
    
    try {
        # Get the item first to get its file reference
        $item = Invoke-WithThrottlingRetry -Description "Get item $ItemId for unlocking" -ScriptBlock {
            Get-PnPListItem -List $ListTitle -Id $ItemId
        }
        
        # Try different approaches based on available PnP version
        $unlocked = $false
        
        # Method 1: Try using the SharePoint REST API for unlocking records
        try {
            # Construct the endpoint for SPPolicyStoreProxy
            $endpoint = "/_api/SP.CompliancePolicy.SPPolicyStoreProxy.UnlockRecordItem"
            
            # Get the list information to construct the listUrl
            $list = Invoke-WithThrottlingRetry -Description "Get list info for unlocking" -ScriptBlock {
                Get-PnPList -Identity $ListTitle
            }
            $listUrl = $list.RootFolder.ServerRelativeUrl
            
            # Construct the request body with correct parameters for SharePoint REST API
            $requestBody = @"
{
    "listUrl": "$listUrl",
    "itemId": "$ItemId"
}
"@
            
            # Execute the REST call with correct parameter syntax and throttling retry
            $result = Invoke-WithThrottlingRetry -Description "Unlock item $ItemId via REST API" -ScriptBlock {
                Invoke-PnPSPRestMethod -Url $endpoint -Method POST -Content $requestBody -ContentType "application/json;odata=verbose"
            }
            Write-HostAndLog "✓ Successfully unlocked record item $itemDisplayName in list '$ListTitle'" -ForegroundColor Green
            $unlocked = $true
        }
        catch {
            # Try alternative parameter format if the first one fails
            try {
                $endpoint2 = "/_api/SP.CompliancePolicy.SPPolicyStoreProxy.UnlockRecordItem(listUrl='$listUrl',itemId='$ItemId')"
                $result2 = Invoke-WithThrottlingRetry -Description "Unlock item $ItemId via REST API (alt format)" -ScriptBlock {
                    Invoke-PnPSPRestMethod -Url $endpoint2 -Method POST
                }
                Write-HostAndLog "✓ Successfully unlocked record item $itemDisplayName in list '$ListTitle' (alternative format)" -ForegroundColor Green
                $unlocked = $true
            }
            catch {
                Write-HostAndLog "      Method 1 (REST API) failed: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        # If the REST API method failed, there's no other reliable way to unlock
        # The compliance fields are read-only and can't be directly modified
        if (-not $unlocked) {
            Write-HostAndLog "✗ REST API unlock failed for item $itemDisplayName in list '$ListTitle'" -ForegroundColor Red
        }
        
        return $unlocked
    }
    catch {
        Write-HostAndLog "✗ Error unlocking record item $itemDisplayName in list '$ListTitle' at $SiteUrl`: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Check if input file exists
if (-not (Test-Path $sitelist)) {
    Write-HostAndLog "Input file not found: $sitelist" -ForegroundColor Red
    exit 1
}

# Initialize logging
Write-Log "=== UNLOCK RECORD ITEMS SCRIPT STARTED ===" -Level 'INFO'
Write-Log "Log file: $script:LogFilePath" -Level 'INFO'
Write-Log "Input file: $sitelist" -Level 'INFO'
Write-Log "Tenant: $tenant" -Level 'INFO'
Write-Log "AppID: $appID" -Level 'INFO'

# Read site URLs from input file
$siteUrls = Get-Content $sitelist | Where-Object { $_.Trim() -ne "" }

Write-HostAndLog "Starting to process $($siteUrls.Count) SharePoint sites..." -ForegroundColor Cyan
Write-HostAndLog "Log file location: $script:LogFilePath" -ForegroundColor Gray
Write-HostAndLog "" -ForegroundColor Gray
Write-HostAndLog "Throttling Protection: Enabled" -ForegroundColor Cyan
Write-HostAndLog "  - Delay between items: $DelayBetweenItems ms" -ForegroundColor Gray
Write-HostAndLog "  - Delay between lists: $DelayBetweenLists ms" -ForegroundColor Gray
Write-HostAndLog "  - Delay between sites: $DelayBetweenSites ms" -ForegroundColor Gray
Write-HostAndLog "  - Max retry attempts: $MaxRetryAttempts" -ForegroundColor Gray
Write-HostAndLog "  - Base retry delay: $BaseRetryDelay ms" -ForegroundColor Gray

$totalUnlockedItems = 0
$totalProcessedSites = 0
$totalProcessedLists = 0
$totalThrottleEvents = 0

foreach ($siteUrl in $siteUrls) {
    $siteUrl = $siteUrl.Trim()
    Write-HostAndLog "`nProcessing site: $siteUrl" -ForegroundColor Yellow
    
    try {
        # Connect to the current SharePoint site with throttling retry
        Invoke-WithThrottlingRetry -Description "Connect to site $siteUrl" -ScriptBlock {
            Connect-PnPOnline -Url $siteUrl -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
        }
        $totalProcessedSites++
        Write-Log "Successfully connected to site: $siteUrl" -Level 'SUCCESS'
        
        # Get all lists in the site with throttling retry
        $lists = Invoke-WithThrottlingRetry -Description "Get lists from site $siteUrl" -ScriptBlock {
            Get-PnPList | Where-Object { 
                $_.Hidden -eq $false -and 
                $_.ItemCount -gt 0 -and 
                $_.Title -notin $ignoreListNames 
            }
        }
        
        Write-HostAndLog "Found $($lists.Count) non-hidden lists with items in this site" -ForegroundColor Cyan
        
        $listCounter = 0
        foreach ($list in $lists) {
            $listCounter++
            Write-HostAndLog "  Processing list $listCounter of $($lists.Count): $($list.Title) ($($list.ItemCount) items)" -ForegroundColor White
            $totalProcessedLists++
            
            try {
                # Get all items in the list (in batches to handle large lists) with throttling retry
                $pageSize = 500
                $items = Invoke-WithThrottlingRetry -Description "Get items from list '$($list.Title)'" -ScriptBlock {
                    Get-PnPListItem -List $list.Title -PageSize $pageSize
                }
                
                $recordItemsFound = 0
                $unlockedInThisList = 0
                $itemCounter = 0
                
                foreach ($item in $items) {
                    $itemCounter++
                    
                    # Throttle protection: delay between items
                    if ($itemCounter -gt 1) {
                        Start-ThrottleDelay -DelayMilliseconds $DelayBetweenItems -Description "between items"
                    }
                    
                    # Progress update every 100 items
                    if ($itemCounter % 100 -eq 0) {
                        Write-VerboseAndLog "    Progress: Processed $itemCounter of $($items.Count) items in list '$($list.Title)'"
                    }
                    
                    # Try to get the item title - use Title field, FileLeafRef, or fallback to ID
                    $itemTitle = $null
                    if ($item.FieldValues.ContainsKey("Title") -and $item.FieldValues["Title"]) {
                        $itemTitle = $item.FieldValues["Title"]
                    }
                    elseif ($item.FieldValues.ContainsKey("FileLeafRef") -and $item.FieldValues["FileLeafRef"]) {
                        $itemTitle = $item.FieldValues["FileLeafRef"]
                    }
                    elseif ($item.FieldValues.ContainsKey("Name") -and $item.FieldValues["Name"]) {
                        $itemTitle = $item.FieldValues["Name"]
                    }
                    
                    # Check if item is locked as a record
                    if (Test-ItemIsLocked -SiteUrl $siteUrl -ListTitle $list.Title -ItemId $item.Id -ItemTitle $itemTitle) {
                        $recordItemsFound++
                        
                        # Attempt to unlock the record item
                        if (Unlock-RecordItem -SiteUrl $siteUrl -ListTitle $list.Title -ItemId $item.Id -ItemTitle $itemTitle) {
                            $unlockedInThisList++
                            $totalUnlockedItems++
                        }
                    }
                }
                
                if ($recordItemsFound -eq 0) {
                    Write-HostAndLog "    No locked record items found in this list" -ForegroundColor Gray
                }
                else {
                    Write-HostAndLog "    Summary: $recordItemsFound locked record items found, $unlockedInThisList unlocked" -ForegroundColor Green
                }
                
                # Throttle protection: delay between lists
                if ($listCounter -lt $lists.Count) {
                    Start-ThrottleDelay -DelayMilliseconds $DelayBetweenLists -Description "between lists"
                }
            }
            catch {
                Write-HostAndLog "    Error processing list '$($list.Title)': $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        # Disconnect from current site
        Disconnect-PnPOnline
        Write-Log "Disconnected from site: $siteUrl" -Level 'INFO'
        
        # Throttle protection: delay between sites
        Start-ThrottleDelay -DelayMilliseconds $DelayBetweenSites -Description "between sites"
    }
    catch {
        Write-HostAndLog "Error connecting to or processing site '$siteUrl': $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Final summary
Write-HostAndLog "=================================" -ForegroundColor Cyan
Write-HostAndLog "PROCESSING COMPLETE" -ForegroundColor Cyan
Write-HostAndLog "=================================" -ForegroundColor Cyan
Write-HostAndLog "Sites processed: $totalProcessedSites of $($siteUrls.Count)" -ForegroundColor Yellow
Write-HostAndLog "Lists processed: $totalProcessedLists" -ForegroundColor Yellow
Write-HostAndLog "Total locked record items unlocked: $totalUnlockedItems" -ForegroundColor Green
Write-HostAndLog "=================================" -ForegroundColor Cyan
Write-HostAndLog "Log file saved to: $script:LogFilePath" -ForegroundColor Gray

# Final log entries
Write-Log "=== PROCESSING SUMMARY ===" -Level 'INFO'
Write-Log "Sites processed: $totalProcessedSites of $($siteUrls.Count)" -Level 'INFO'
Write-Log "Lists processed: $totalProcessedLists" -Level 'INFO'
Write-Log "Total locked record items unlocked: $totalUnlockedItems" -Level 'SUCCESS'
Write-Log "Throttling protection settings: Items=$DelayBetweenItems ms, Lists=$DelayBetweenLists ms, Sites=$DelayBetweenSites ms, MaxRetries=$MaxRetryAttempts" -Level 'INFO'
Write-Log "=== UNLOCK RECORD ITEMS SCRIPT COMPLETED ===" -Level 'INFO'
