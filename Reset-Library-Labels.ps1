<#
.SYNOPSIS
    Reset and reapply retention labels on SharePoint Online lists across multiple sites.

.DESCRIPTION
    This script processes SharePoint Online sites and resets retention labels on lists that match
    a specified target label. It can operate in two modes:
    
    1. REPORT-ONLY MODE ($resetlabel = $false): Scans lists and reports which ones have the target
       label without making any changes.
    
    2. RESET MODE ($resetlabel = $true): Actively resets and reapplies the retention label to 
       refresh metadata and trigger compliance jobs.
    
    The script uses Microsoft Graph API to query retention labels from list items and PnP PowerShell
    cmdlets to reset and apply labels. It includes comprehensive throttling protection with 
    exponential backoff and retry logic to handle SharePoint Online throttling gracefully.

.NOTES
    File Name      : Reset-Library-Labels.ps1
    Author         : Mike Lee
    Prerequisite   : PnP.PowerShell module, Entra App Registration with certificate authentication
    Date:          : 1/14/26
    Version        : 1.0
    
    IMPORTANT: Label application triggers background jobs in Microsoft 365 compliance services.
    These jobs update item metadata and sync changes to the compliance store and Graph endpoints.
    It can take up to 24 hours for labels to appear consistently in Graph queries.

.PARAMETER None
    This script uses configuration variables instead of parameters. All settings are configured
    in the "USER CONFIGURATION" section at the top of the script.

.CONFIGURATION
    $EnableVerbose          - Enable detailed debug output (true/false)
    $LabelName              - OPTIONAL: Specific retention label to target. Leave empty ("") to reset ANY label found
    $appID                  - Entra Application (Client) ID for authentication
    $tenant                 - Tenant ID (GUID)
    $thumbprint             - Certificate thumbprint for certificate-based authentication
    $sitelist               - Path to text file containing SharePoint site URLs (one per line)
    $ignoreListNames        - Array of list names to skip during processing
    $resetlabel             - Enable reset mode (true) or report-only mode (false)
    $DelayBetweenLists      - Throttling delay in milliseconds between list operations
    $DelayBetweenSites      - Throttling delay in milliseconds between site connections
    $MaxRetryAttempts       - Maximum retry attempts for throttled requests
    $BaseRetryDelay         - Base delay for exponential backoff on throttling

.EXAMPLE
    .\Reset-Library-Labels.ps1
    
    Runs the script with settings configured in the USER CONFIGURATION section.
    If $resetlabel = $false, it will report on lists with the target label.
    If $resetlabel = $true, it will reset and reapply labels.

.INPUTS
    Text file containing SharePoint site URLs (one per line)
    Example content:
        https://tenant.sharepoint.com/sites/site1
        https://tenant.sharepoint.com/sites/site2

.OUTPUTS
    Console output with color-coded status messages
    Log file in %TEMP% directory: Reset_Retention_Labels_<timestamp>_logfile.log

.LINK
    PnP PowerShell: https://pnp.github.io/powershell/
    Microsoft Graph API: https://learn.microsoft.com/en-us/graph/api/overview

.AUTHENTICATION
    This script uses certificate-based authentication with an Entra App Registration.
    Required API Permissions:
    - SharePoint: Sites.FullControl.All
    - Graph: Sites.Read.All
    
    The certificate must be installed in the user's certificate store and the thumbprint
    must be configured in the $thumbprint variable.

#>

# =================================================================================================
# START OF USER CONFIGURATION
# =================================================================================================
# --- General Settings ---
$EnableVerbose = $false         # Set to $true to see detailed verbose logging
$LabelName = ""                # OPTIONAL: Specific label to target. Leave empty ("") to reset ANY label found

# --- Tenant and App Registration Details ---
$appID = "0c60e510-69de-44bf-999f-1f64d4d3ceb0"                 # This is your Entra App ID
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"                # This is your Tenant ID
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"        # This is certificate thumbprint

# --- Input File Path ---
$sitelist = 'C:\temp\SPOSiteList.txt' # Path to the input file containing site URLs

# --- Lists to Ignore ---
$ignoreListNames = @('Site Assets', 'Site Pages') # Lists to skip during processing

# --- Reset Label Configuration ---
$resetlabel = $false # Set to $true to reset labels, $false for report-only mode

# --- Logging Configuration ---
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss" # Current date and time for unique file naming
$script:LogFilePath = "$env:TEMP\" + 'Reset_Retention_Labels_' + $date + '_' + "logfile.log" # Define log file path
$script:EnableLogging = $true

# --- Throttling Protection Settings ---
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
    if ($script:EnableVerbose) {
        Write-Host $Message -ForegroundColor DarkGray
    }
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

#region GRAPH API FUNCTIONS
# =====================================================================================
# Functions to interact with Microsoft Graph API for retention label information
# =====================================================================================

function Get-GraphAccessToken {
    try {
        $token = Invoke-WithThrottlingRetry -Description "Get Graph API access token" -ScriptBlock {
            Get-PnPAccessToken
        }
        return $token
    }
    catch {
        Write-WarningAndLog "Failed to get Graph API access token: $($_.Exception.Message)"
        return $null
    }
}

# Function to get retention label for a list using Graph API
function Get-ListRetentionLabel {
    param(
        [string]$ListTitle
    )
    
    try {
        # Get the list object first to get its ID
        $list = Invoke-WithThrottlingRetry -Description "Get list object for '$ListTitle'" -ScriptBlock {
            Get-PnPList -Identity $ListTitle -ErrorAction Stop
        }
        
        if (-not $list) {
            Write-WarningAndLog "Could not find list '$ListTitle'"
            return $null
        }
        
        # Get site ID from current connection
        $site = Invoke-WithThrottlingRetry -Description "Get current site" -ScriptBlock {
            Get-PnPSite -Includes Id
        }
        
        $siteId = $site.Id
        $listId = $list.Id
        
        Write-VerboseAndLog "Site ID: $siteId, List ID: $listId"
        
        # Get Graph API access token
        $accessToken = Get-GraphAccessToken
        if (-not $accessToken) {
            Write-WarningAndLog "Could not get access token for Graph API"
            return $null
        }
        
        # Build Graph API URL to get list items with compliance tag field
        $graphUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items?`$expand=fields(`$select=id,_ComplianceTag,_ComplianceTagWrittenTime)&`$top=1"
        
        if ($script:EnableVerbose) {
            Write-HostAndLog "    [DEBUG] Graph API URL: $graphUrl" -ForegroundColor Magenta
        }
        
        # Make Graph API call
        $result = Invoke-WithThrottlingRetry -Description "Get retention label via Graph API for '$ListTitle'" -ScriptBlock {
            $headers = @{
                "Authorization" = "Bearer $accessToken"
                "Accept"        = "application/json"
            }
            
            $response = Invoke-RestMethod -Uri $graphUrl -Headers $headers -Method Get -ErrorAction Stop
            return $response
        }
        
        # Debug output - show what we got back
        if ($script:EnableVerbose) {
            Write-HostAndLog "    [DEBUG] Graph API result:" -ForegroundColor Magenta
            if ($null -eq $result) {
                Write-HostAndLog "    [DEBUG] - Result is NULL" -ForegroundColor Magenta
            }
            elseif ($result.value -and $result.value.Count -gt 0) {
                Write-HostAndLog "    [DEBUG] - Found $($result.value.Count) items" -ForegroundColor Magenta
                $firstItem = $result.value[0]
                if ($firstItem.fields) {
                    Write-HostAndLog "    [DEBUG] - _ComplianceTag: $($firstItem.fields._ComplianceTag)" -ForegroundColor Magenta
                    Write-HostAndLog "    [DEBUG] - _ComplianceTagWrittenTime: $($firstItem.fields._ComplianceTagWrittenTime)" -ForegroundColor Magenta
                }
            }
            else {
                Write-HostAndLog "    [DEBUG] - No items found in list or empty response" -ForegroundColor Magenta
            }
        }
        
        # Try to extract label name from the first item
        $labelName = $null
        if ($result.value -and $result.value.Count -gt 0) {
            $firstItem = $result.value[0]
            if ($firstItem.fields -and $firstItem.fields._ComplianceTag) {
                $labelName = $firstItem.fields._ComplianceTag
            }
        }
        
        if ($labelName) {
            Write-VerboseAndLog "List '$ListTitle' has retention label: $labelName"
            return @{TagName = $labelName }
        }
        else {
            Write-VerboseAndLog "List '$ListTitle' has no retention label set."
            return $null
        }
    }
    catch {
        Write-WarningAndLog "Could not get retention label for list '$ListTitle': $($_.Exception.Message)"
        if ($script:EnableVerbose -and $_.Exception.Response) {
            Write-HostAndLog "    [DEBUG] HTTP Status: $($_.Exception.Response.StatusCode)" -ForegroundColor Magenta
        }
        return $null
    }
}

#endregion GRAPH API FUNCTIONS

#region LABEL MANAGEMENT FUNCTIONS
# =====================================================================================
# Functions to reset and apply retention labels to lists
# =====================================================================================

function Reset-AndSetRetentionLabel {
    param(
        [string]$ListTitle,
        [string]$TargetLabelName  # Optional: if empty, reset any label found
    )
    
    try {
        # Get current label
        $currentLabel = Get-ListRetentionLabel -ListTitle $ListTitle
        
        if (-not $currentLabel) {
            Write-HostAndLog "    No label set - skipping" -ForegroundColor Gray
            return $false
        }
        
        $foundLabelName = $currentLabel.TagName
        Write-HostAndLog "    Current label: $foundLabelName" -ForegroundColor Cyan
        
        # If a specific target label is specified, check if current label matches
        if (-not [string]::IsNullOrWhiteSpace($TargetLabelName)) {
            # Only process if the current label matches the target label
            # The TagName might include additional text like "(Retain for 1 years)", so check if it starts with the target
            if (-not ($foundLabelName -like "$TargetLabelName*")) {
                Write-HostAndLog "    Label '$foundLabelName' does not match target '$TargetLabelName' - skipping" -ForegroundColor Gray
                return $false
            }
        }
        # If no target specified, process any label found
        
        if ($script:resetlabel) {
            # RESET MODE: Actually perform the reset and set operations
            Write-HostAndLog "    Label found - resetting..." -ForegroundColor Yellow
            
            # Reset the retention label
            Write-VerboseAndLog "Resetting retention label for list '$ListTitle'"
            Invoke-WithThrottlingRetry -Description "Reset retention label for list '$ListTitle'" -ScriptBlock {
                Reset-PnPRetentionLabel -List $ListTitle
            }
            Write-HostAndLog "    ✓ Reset retention label" -ForegroundColor Green
            
            # Set the retention label back using the label name we found
            Write-VerboseAndLog "Re-applying retention label '$foundLabelName' for list '$ListTitle'"
            Invoke-WithThrottlingRetry -Description "Set retention label for list '$ListTitle'" -ScriptBlock {
                Set-PnPRetentionLabel -List $ListTitle -Label $foundLabelName
            }
            Write-HostAndLog "    ✓ Re-applied label: $foundLabelName" -ForegroundColor Green
        }
        else {
            # REPORT-ONLY MODE: Just report what would be done
            Write-HostAndLog "    Label found - WOULD RESET (report-only mode)" -ForegroundColor Yellow
            Write-HostAndLog "    [REPORT] Would reset and re-apply label: $foundLabelName" -ForegroundColor Cyan
        }
        
        return $true
    }
    catch {
        Write-HostAndLog "    ✗ Error resetting/setting retention label for list '$ListTitle': $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

#endregion LABEL MANAGEMENT FUNCTIONS

#region MAIN SCRIPT EXECUTION
# =====================================================================================
# Main script execution: validation, initialization, and processing
# =====================================================================================

# Check if input file exists
if (-not (Test-Path $sitelist)) {
    Write-HostAndLog "Input file not found: $sitelist" -ForegroundColor Red
    exit 1
}

#region INITIALIZATION
# Initialize logging
Write-Log "=== RESET RETENTION LABELS SCRIPT STARTED ===" -Level 'INFO'
Write-Log "Mode: $(if ($resetlabel) {'RESET MODE'} else {'REPORT-ONLY MODE'})" -Level 'INFO'
Write-Log "Log file: $script:LogFilePath" -Level 'INFO'
Write-Log "Input file: $sitelist" -Level 'INFO'
Write-Log "Tenant: $tenant" -Level 'INFO'
Write-Log "AppID: $appID" -Level 'INFO'
Write-Log "Target Label: $(if ([string]::IsNullOrWhiteSpace($LabelName)) {'ANY label found'} else {$LabelName})" -Level 'INFO'

# Read site URLs from input file
$siteUrls = Get-Content $sitelist | Where-Object { $_.Trim() -ne "" }

Write-HostAndLog "Starting to $(if ($resetlabel) {'reset'} else {'report on'}) retention labels on $($siteUrls.Count) SharePoint sites..." -ForegroundColor Cyan
Write-HostAndLog "Mode: $(if ($resetlabel) {'RESET MODE - Labels will be reset'} else {'REPORT-ONLY MODE - No changes will be made'})" -ForegroundColor $(if ($resetlabel) { 'Yellow' } else { 'Green' })
Write-HostAndLog "Target Label: $(if ([string]::IsNullOrWhiteSpace($LabelName)) {'ANY label found'} else {$LabelName})" -ForegroundColor Yellow
Write-HostAndLog "" -ForegroundColor Gray
Write-HostAndLog "⚠ IMPORTANT: Label application triggers background jobs in Microsoft 365 compliance services." -ForegroundColor Yellow
Write-HostAndLog "  These jobs update item metadata and sync changes to the compliance store and Graph endpoints." -ForegroundColor Yellow
Write-HostAndLog "  It can take up to 24 hours for the label to appear consistently in Graph queries." -ForegroundColor Yellow
Write-HostAndLog "" -ForegroundColor Gray
Write-HostAndLog "Log file location: $script:LogFilePath" -ForegroundColor Gray
Write-HostAndLog "" -ForegroundColor Gray
Write-HostAndLog "Throttling Protection: Enabled" -ForegroundColor Cyan
Write-HostAndLog "  - Delay between lists: $DelayBetweenLists ms" -ForegroundColor Gray
Write-HostAndLog "  - Delay between sites: $DelayBetweenSites ms" -ForegroundColor Gray
Write-HostAndLog "  - Max retry attempts: $MaxRetryAttempts" -ForegroundColor Gray
Write-HostAndLog "  - Base retry delay: $BaseRetryDelay ms" -ForegroundColor Gray

$totalUpdatedLists = 0
$totalProcessedSites = 0
$totalProcessedLists = 0

#endregion INITIALIZATION

#region SITE PROCESSING LOOP
# Process each SharePoint site
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
            Write-HostAndLog "  Processing list $listCounter of $($lists.Count): $($list.Title)" -ForegroundColor White
            $totalProcessedLists++
            
            try {
                # Reset and set retention label for the list
                if (Reset-AndSetRetentionLabel -ListTitle $list.Title -TargetLabelName $LabelName) {
                    $totalUpdatedLists++
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

#endregion SITE PROCESSING LOOP

#region FINAL SUMMARY
# Display and log final processing results
Write-HostAndLog "=================================" -ForegroundColor Cyan
Write-HostAndLog "PROCESSING COMPLETE" -ForegroundColor Cyan
Write-HostAndLog "=================================" -ForegroundColor Cyan
Write-HostAndLog "Sites processed: $totalProcessedSites of $($siteUrls.Count)" -ForegroundColor Yellow
Write-HostAndLog "Lists processed: $totalProcessedLists" -ForegroundColor Yellow
Write-HostAndLog "Lists $(if ($resetlabel) {'updated'} else {'found'}) with retention labels: $totalUpdatedLists" -ForegroundColor Green
Write-HostAndLog "=================================" -ForegroundColor Cyan
Write-HostAndLog "Log file saved to: $script:LogFilePath" -ForegroundColor Gray

# Final log entries
Write-Log "=== PROCESSING SUMMARY ===" -Level 'INFO'
Write-Log "Mode: $(if ($resetlabel) {'RESET MODE'} else {'REPORT-ONLY MODE'})" -Level 'INFO'
Write-Log "Sites processed: $totalProcessedSites of $($siteUrls.Count)" -Level 'INFO'
Write-Log "Lists processed: $totalProcessedLists" -Level 'INFO'
Write-Log "Lists $(if ($resetlabel) {'updated'} else {'found'}) with retention labels: $totalUpdatedLists" -Level 'SUCCESS'
Write-Log "Throttling protection settings: Lists=$DelayBetweenLists ms, Sites=$DelayBetweenSites ms, MaxRetries=$MaxRetryAttempts" -Level 'INFO'
Write-Log "=== RESET RETENTION LABELS SCRIPT COMPLETED ===" -Level 'INFO'

#endregion FINAL SUMMARY

#endregion MAIN SCRIPT EXECUTION
