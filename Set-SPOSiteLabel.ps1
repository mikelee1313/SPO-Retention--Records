#Requires -Version 5.1

<#
.SYNOPSIS
Applies one sensitivity label to sites listed in a text file.
.DESCRIPTION
Uses an existing SharePoint Online Management Shell installation. Processes sites
sequentially and writes a unique CSV audit log. Supports -WhatIf.
On PowerShell 7 or later, imports the SPO module using Windows PowerShell
compatibility mode. Requires Windows and Windows PowerShell 5.1.
.EXAMPLE
.\Set-SPOSiteLabel.ps1 -SiteListPath .\sites.txt -AdminSiteURL https://contoso-admin.sharepoint.com -LabelGuid '00000000-0000-0000-0000-000000000001'
.NOTES
To obtain the label GUID, an administrator must connect to Microsoft Purview
Security & Compliance PowerShell and list the sensitivity labels:

    Import-Module ExchangeOnlineManagement
    Connect-IPPSSession -UserPrincipalName admin@contoso.com
    Get-Label | Select-Object DisplayName, Name, Guid, ContentType | FL

The ExchangeOnlineManagement module must be installed, and the account must
have Purview permissions to view sensitivity labels. SharePoint admin permissions
alone might not grant this access.

Pass the desired label's Guid value to -LabelGuid, not its display name.
Display names can be duplicated. Confirm that the label is configured for
Groups & sites.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$SiteListPath,

    [uri]$AdminSiteURL = 'https://m365cpi13246019-admin.sharepoint.com',
    [guid]$LabelGuid = 'cf456e86-4a06-4a56-9884-6f52cbe98f0c',
    [string]$LogDirectory = (Join-Path $PSScriptRoot 'Logs'),

    [ValidateRange(0, 20)]
    [int]$MaxRetries = 6,

    [ValidateRange(1, 3600)]
    [int]$InitialRetrySeconds = 5,

    [ValidateRange(1, 3600)]
    [int]$MaxBackoffSeconds = 300,

    [ValidateRange(0, 3600)]
    [int]$RequestDelaySeconds = 1
)

#region Input validation and site list loading
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($AdminSiteURL.Scheme -ne 'https' -or -not $AdminSiteURL.IsAbsoluteUri) {
    throw 'AdminSiteURL must be an absolute HTTPS URL.'
}
if ($LabelGuid -eq [guid]::Empty) {
    throw 'LabelGuid cannot be the empty GUID.'
}

$sites = [System.Collections.Generic.List[string]]::new()
$seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($line in (Get-Content -LiteralPath $SiteListPath -ErrorAction Stop)) {
    $site = $line.Trim()
    if (-not $site -or $site.StartsWith('#')) { continue }
    $uri = $null
    if (-not [uri]::TryCreate($site, [UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -ne 'https' -or $uri.UserInfo -or $uri.Query -or $uri.Fragment) {
        throw "Invalid site URL in input: '$site'. Supply absolute HTTPS site URLs without queries or fragments."
    }
    $site = $uri.AbsoluteUri.TrimEnd('/')
    if ($seen.Add($site)) { $sites.Add($site) }
}
if ($sites.Count -eq 0) { throw 'The input file contains no site URLs.' }
#endregion

#region Audit log initialization
$runId = [guid]::NewGuid().ToString()
$targetLabel = $LabelGuid.ToString()
$auditFailed = $false
$auditStarted = $false
$null = New-Item -ItemType Directory -Path $LogDirectory -Force -WhatIf:$false
$auditPath = Join-Path $LogDirectory ("SiteLabels-{0}-{1}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $runId)
#endregion

#region Audit logging helper
function Write-Audit {
    param(
        [string]$SiteUrl, [string]$Operation, [string]$Outcome,
        [string]$BeforeLabel = '', [string]$AfterLabel = '',
        [int]$Attempt = 0, [string]$HttpStatus = '',
        [double]$WaitSeconds = 0, [string]$Message = ''
    )
    $row = [pscustomobject][ordered]@{
        TimestampUtc = [DateTimeOffset]::UtcNow.ToString('o')
        RunId = $runId
        Operator = [Environment]::UserName
        AdminUrl = $AdminSiteURL.AbsoluteUri
        SiteUrl = $SiteUrl
        Operation = $Operation
        Outcome = $Outcome
        BeforeLabel = $BeforeLabel
        RequestedLabel = $targetLabel
        AfterLabel = $AfterLabel
        Attempt = $Attempt
        HttpStatus = $HttpStatus
        WaitSeconds = $WaitSeconds
        Message = $Message
    }
    try {
        if ($script:auditStarted) {
            $row | Export-Csv -LiteralPath $auditPath -NoTypeInformation -Encoding UTF8 -Append -ErrorAction Stop -WhatIf:$false
        }
        else {
            $row | Export-Csv -LiteralPath $auditPath -NoTypeInformation -Encoding UTF8 -NoClobber -ErrorAction Stop -WhatIf:$false
            $script:auditStarted = $true
        }
    }
    catch {
        $script:auditFailed = $true
        throw "Audit write failed; stopping to prevent unaudited changes. Path: $auditPath. $($_.Exception.Message)"
    }
}

#endregion

#region Throttling detection
function Get-ThrottleInfo {
    param([System.Management.Automation.ErrorRecord]$Record)

    $status = ''
    $retryAfter = ''
    $isThrottle = $false
    $exception = $Record.Exception
    while ($null -ne $exception) {
        $response = $null
        if ($exception.PSObject.Properties['Response']) { $response = $exception.Response }
        foreach ($source in @($response, $exception)) {
            if ($null -eq $source) { continue }
            if ($source.PSObject.Properties['StatusCode']) {
                $code = [int]$source.StatusCode
                if ($code -in 429, 503) {
                    $status = [string]$code
                    $isThrottle = $true
                }
            }
            if ($source.PSObject.Properties['Headers'] -and $null -ne $source.Headers) {
                $headers = $source.Headers
                if ($headers.PSObject.Properties['RetryAfter']) {
                    if ($null -ne $headers.RetryAfter) { $retryAfter = $headers.RetryAfter.ToString() }
                }
                else {
                    $value = $headers['Retry-After']
                    if ($null -ne $value) { $retryAfter = [string](@($value)[0]) }
                }
            }
        }
        # CSOM can wrap throttling without exposing the HTTP response.
        if ($exception.PSObject.Properties['ServerErrorCode'] -and
            $exception.ServerErrorCode -eq -2147024860) { $isThrottle = $true }
        if ($exception.Message -match '(?i)\b(429|503)\b|throttl|too many requests') {
            $isThrottle = $true
        }
        $exception = $exception.InnerException
    }
    if ($retryAfter) { $isThrottle = $true }
    [pscustomobject]@{ IsThrottle = $isThrottle; Status = $status; RetryAfter = $retryAfter }
}

#endregion

#region Request pacing and retry handling
function Invoke-SPOWithRetry {
    param([scriptblock]$Action, [string]$SiteUrl, [string]$Operation)

    for ($attempt = 1; ; $attempt++) {
        try {
            if ($RequestDelaySeconds -gt 0) { Start-Sleep -Seconds $RequestDelaySeconds }
            return (& $Action)
        }
        catch {
            $info = Get-ThrottleInfo -Record $_
            if (-not $info.IsThrottle) {
                Write-Audit -SiteUrl $SiteUrl -Operation $Operation -Outcome 'RequestFailed' -Attempt $attempt -Message $_.Exception.Message
                throw
            }
            $seconds = [Math]::Min($MaxBackoffSeconds, $InitialRetrySeconds * [Math]::Pow(2, $attempt - 1))
            $seconds += (Get-Random -Minimum 0 -Maximum 1000) / 1000.0
            $delaySource = 'exponential backoff (Retry-After unavailable or invalid)'
            if ($info.RetryAfter) {
                $delta = 0L
                $date = [DateTimeOffset]::MinValue
                if ([long]::TryParse($info.RetryAfter, [ref]$delta) -and $delta -ge 0) {
                    $seconds = [double]$delta
                    $delaySource = 'Retry-After seconds'
                }
                elseif ([DateTimeOffset]::TryParse($info.RetryAfter, [Globalization.CultureInfo]::InvariantCulture,
                    [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$date)) {
                    $seconds = [Math]::Max(0, [Math]::Ceiling(($date - [DateTimeOffset]::UtcNow).TotalSeconds))
                    $delaySource = 'Retry-After HTTP date'
                }
            }
            $outcome = 'RetryScheduled'
            if ($attempt -gt $MaxRetries) { $outcome = 'RetriesExhausted' }
            Write-Audit -SiteUrl $SiteUrl -Operation $Operation -Outcome $outcome -Attempt $attempt `
                -HttpStatus $info.Status -WaitSeconds $seconds -Message "$delaySource. $($_.Exception.Message)"
            Write-Warning "$Operation throttled for '$SiteUrl' (attempt $attempt). Waiting $seconds seconds using $delaySource."
            # Even after exhaustion, honor the cooldown before moving to another site.
            $remaining = [Math]::Ceiling($seconds)
            while ($remaining -gt 0) {
                $chunk = [Math]::Min($remaining, 3600)
                Start-Sleep -Seconds $chunk
                $remaining -= $chunk
            }
            if ($attempt -gt $MaxRetries) { throw }
        }
    }
}

#endregion

#region Run execution and error handling
Write-Audit -Operation 'Run' -Outcome 'Started' -Message "Input: $SiteListPath; unique sites: $($sites.Count); WhatIf: $WhatIfPreference"
Write-Host "Audit log: $auditPath"
$failed = 0
try {
    #region Module compatibility and SharePoint connection
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
    }

    Invoke-SPOWithRetry -Operation 'Connect' -Action {
        Connect-SPOService -Url $AdminSiteURL.AbsoluteUri -ErrorAction Stop
    } | Out-Null
    #endregion

    #region Site label updates and verification
    foreach ($site in $sites) {
        $before = ''
        $after = ''
        $setAccepted = $false
        try {
            $current = Invoke-SPOWithRetry -SiteUrl $site -Operation 'ReadBefore' -Action {
                Get-SPOSite -Identity $site -ErrorAction Stop
            }
            $before = [string]$current.SensitivityLabel
            if ($before -eq $targetLabel) {
                Write-Audit -SiteUrl $site -Operation 'SetLabel' -Outcome 'AlreadyLabeled' -BeforeLabel $before -AfterLabel $before
                continue
            }
            if (-not $PSCmdlet.ShouldProcess($site, "Set sensitivity label to $targetLabel")) {
                Write-Audit -SiteUrl $site -Operation 'SetLabel' -Outcome 'Skipped' -BeforeLabel $before -AfterLabel $before
                continue
            }
            Write-Audit -SiteUrl $site -Operation 'SetLabel' -Outcome 'ChangeStarted' -BeforeLabel $before
            Invoke-SPOWithRetry -SiteUrl $site -Operation 'SetLabel' -Action {
                Set-SPOSite -Identity $site -SensitivityLabel $targetLabel -ErrorAction Stop
            } | Out-Null
            $setAccepted = $true
            Write-Audit -SiteUrl $site -Operation 'SetLabel' -Outcome 'ChangeAccepted' -BeforeLabel $before
            $updated = Invoke-SPOWithRetry -SiteUrl $site -Operation 'ReadAfter' -Action {
                Get-SPOSite -Identity $site -ErrorAction Stop
            }
            $after = [string]$updated.SensitivityLabel
            if ($after -ne $targetLabel) {
                throw "Verification mismatch: requested '$targetLabel', observed '$after'. Reconcile the site before rerunning."
            }
            Write-Audit -SiteUrl $site -Operation 'SetLabel' -Outcome 'Verified' -BeforeLabel $before -AfterLabel $after
        }
        catch {
            if ($auditFailed) { throw }
            $failed++
            $outcome = 'Failed'
            if ($setAccepted) { $outcome = 'Unverified' }
            Write-Audit -SiteUrl $site -Operation 'SetLabel' -Outcome $outcome -BeforeLabel $before -AfterLabel $after -Message $_.Exception.Message
            Write-Warning "Site '$site': $outcome. $($_.Exception.Message)"
        }
    }
    #endregion
}
catch {
    if (-not $auditFailed) { Write-Audit -Operation 'Run' -Outcome 'Aborted' -Message $_.Exception.Message }
    throw
}
#endregion

#region Final audit summary and run status
Write-Audit -Operation 'Run' -Outcome 'Completed' -Message "Sites: $($sites.Count); failed or unverified: $failed"
if ($failed -gt 0) {
    throw "$failed site(s) failed or could not be verified. Review audit log: $auditPath"
}
Write-Host "Completed. Audit log: $auditPath"
#endregion
