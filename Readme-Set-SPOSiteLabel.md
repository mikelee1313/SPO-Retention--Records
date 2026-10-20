# Batch SharePoint site labeling

Requires PowerShell 5.1 or later and the SharePoint Online Management Shell
(`Microsoft.Online.SharePoint.PowerShell`). Authenticate with an account permitted
to read sites and change their sensitivity labels.

On PowerShell 7 or later, the script automatically loads the SPO module with
Windows PowerShell compatibility mode:

```powershell
Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
```

This requires Windows, Windows PowerShell 5.1, and the SPO module installed where
Windows PowerShell can discover it. Module import failures abort the run and are
recorded in the audit log. Windows PowerShell 5.1 retains its existing module
loading behavior.

## Get the label GUID

Before running the script, an administrator must connect to Microsoft Purview
Security & Compliance PowerShell using `Connect-IPPSSession`, then list the
available sensitivity labels:

```powershell
Import-Module ExchangeOnlineManagement
Connect-IPPSSession -UserPrincipalName admin@contoso.com
Get-Label | Select-Object DisplayName, Name, Guid, ContentType | FL
```

The `ExchangeOnlineManagement` module must be installed, and the account must
have Purview permissions to view sensitivity labels. SharePoint admin permissions
alone might not grant this access.

Copy the desired label's `Guid` value and pass it to `-LabelGuid`. The script
does not accept a display name, and display names can be duplicated. Confirm
that the selected label is configured for **Groups & sites**.

## Prepare the site list and run

Create `sites.txt` with one absolute HTTPS site URL per line. Blank lines and
lines starting with `#` are ignored; duplicate URLs are processed once.
All input is validated before connecting or changing any site.

```text
https://contoso.sharepoint.com/sites/Finance
https://contoso.sharepoint.com/teams/Operations
```

```powershell
# Preview: connects and reads current labels but does not change sites.
.\Set-SPOSiteLabel.ps1 -SiteListPath .\sites.txt `
    -AdminSiteURL https://contoso-admin.sharepoint.com `
    -LabelGuid '00000000-0000-0000-0000-000000000001' -WhatIf

# Apply after reviewing the preview and selecting the real label GUID.
.\Set-SPOSiteLabel.ps1 -SiteListPath .\sites.txt `
    -AdminSiteURL https://contoso-admin.sharepoint.com `
    -LabelGuid '00000000-0000-0000-0000-000000000001' `
    -LogDirectory .\Logs
```

The existing development admin URL and label GUID remain the defaults. Pass
both explicitly for other tenants. This changes **site/container** sensitivity
labels, not labels on individual files.

## Throttling and failures

- Sites and requests run sequentially, with a default one-second delay before
  each request (`-RequestDelaySeconds`).
- Connection, reads, writes, and verification retry HTTP 429/503 and recognizable
  wrapped CSOM throttling errors. Other errors are not retried.
- When exposed by the cmdlet exception/inner exception, `Retry-After` delta
  seconds and HTTP dates take precedence and are never capped by the fallback
  limit. Even the last failed attempt waits before the next site is attempted.
- **SPO cmdlets do not always expose response headers.** When a header is
  unavailable or invalid, the script warns and audits that it is using
  exponential backoff with jitter instead. This cannot recover headers hidden
  inside the SPO module or control any internal retries performed by that module.
- Defaults: six retries after the initial attempt, five-second initial fallback,
  300-second fallback cap (plus less than one second of jitter). Configure using
  `-MaxRetries`, `-InitialRetrySeconds`, and `-MaxBackoffSeconds`.
- Site failures are logged and processing continues. Any failed/unverified sites
  produce a terminating error at the end. Connection failure aborts the run.

## Audit records

Each invocation creates a unique UTF-8 CSV in `Logs` (or `-LogDirectory`), also
for previews. Rows include UTC time, run ID, local operator identity, tenant
admin URL, site, requested/before/after labels, operation, outcome, retry attempt,
HTTP status when available, wait duration, and error text.

`ChangeStarted` is written **before** the write request; `ChangeAccepted` means
the cmdlet returned successfully; `Verified` requires a read-back matching the
requested label. Already-matching sites are logged without a write. A successful
write whose verification fails is `Unverified`, not a reported success. Label
propagation may be delayed; reconcile these sites before rerunning.

An audit write failure stops the run immediately. Abrupt termination or lost
responses can leave a `ChangeStarted` without a final outcome: inspect that site,
as a request might have applied despite an error. These logs are a local
operational trail, not tamper-proof compliance storage or proof of the signed-in
cloud identity. Protect/archive them appropriately, and use Microsoft Purview
audit for independent service-side evidence. When importing logs into spreadsheet
software, import text fields as text rather than evaluating formulas.

## Local validation

Run `.\Test-Set-SPOSiteLabel.ps1` for offline tests with mocked SPO commands.
Tests do not connect to SharePoint or change live sites.
