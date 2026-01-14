# Reset Library Labels for SharePoint Online

A PowerShell automation script to reset and reapply retention labels on SharePoint Online document libraries across multiple sites. This tool is designed to refresh compliance metadata and trigger Microsoft 365 compliance background jobs without manually processing each library.

## 🎯 Overview

When retention labels are applied to SharePoint libraries, the metadata and compliance settings can sometimes become stale or need to be refreshed. This script automates the process of:

1. **Discovering** retention labels applied to libraries using Microsoft Graph API
2. **Resetting** the existing label to clear cached metadata
3. **Reapplying** the same label to trigger fresh compliance processing

The script supports both **report-only mode** (discover and report without changes) and **reset mode** (actively reset labels).

## ✨ Key Features

- **Flexible Label Targeting**: Reset a specific label by name, or reset ANY label found on libraries
- **Batch Processing**: Process multiple SharePoint sites from a simple text file
- **Two Operating Modes**:
  - **Report-Only Mode**: Scan and report which libraries have labels without making changes
  - **Reset Mode**: Actually reset and reapply labels
- **Microsoft Graph API Integration**: Query retention labels directly from list items for real-time accuracy
- **Comprehensive Throttling Protection**:
  - Automatic retry with exponential backoff
  - Honors SharePoint Online throttling (HTTP 429/503)
  - Configurable delays between operations
- **Detailed Logging**: Console output with color-coding + timestamped log file
- **Certificate-Based Authentication**: Secure authentication using Entra App Registration
- **Verbose Debug Mode**: Optional detailed output for troubleshooting

## 📋 Prerequisites

### Required Software
- **PowerShell 5.1** or higher (PowerShell 7+ recommended)
- **PnP.PowerShell Module**: Latest version
  ```powershell
  Install-Module PnP.PowerShell -Scope CurrentUser
  ```

### Required Permissions
- **Entra App Registration** with certificate-based authentication
- **API Permissions**:
  - SharePoint: `Sites.FullControl.All`
  - Microsoft Graph: `Sites.Read.All`
- **Certificate**: Installed in user's certificate store with private key

## 🚀 Quick Start


### 1. Configure Authentication

Create an Entra App Registration with certificate authentication:

1. Navigate to **Azure Portal** > **Entra ID** > **App registrations**
2. Click **New registration**
3. Name it (e.g., "SharePoint Label Reset Tool")
4. Set **Supported account types** to "Single tenant"
5. Click **Register**
6. Note the **Application (client) ID** and **Directory (tenant) ID**

#### Upload Certificate

1. Generate a self-signed certificate (or use existing):
   ```powershell
   $cert = New-SelfSignedCertificate -Subject "CN=SPOLabelReset" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256
   ```
2. Export the public key (.cer file)
3. In the App Registration, go to **Certificates & secrets**
4. Upload the .cer file
5. Note the certificate **thumbprint**

#### Grant API Permissions

1. In the App Registration, go to **API permissions**
2. Click **Add a permission**
3. Select **SharePoint** > **Application permissions** > `Sites.FullControl.All`
4. Select **Microsoft Graph** > **Application permissions** > `Sites.Read.All`
5. Click **Grant admin consent**

### 2. Create Site List File

Create a text file with SharePoint site URLs (one per line):

**C:\temp\SPOSiteList.txt**
```
https://contoso.sharepoint.com/sites/site1
https://contoso.sharepoint.com/sites/site2
https://contoso.sharepoint.com/sites/site3
```

### 3. Configure the Script

Edit the configuration section at the top of `Reset-Library-Labels.ps1`:

```powershell
# --- General Settings ---
$EnableVerbose = $false         # Set to $true for detailed debug output
$LabelName = ""                 # Leave empty to reset ANY label, or specify a label name

# --- Tenant and App Registration Details ---
$appID = "YOUR-APP-ID-HERE"                 
$tenant = "YOUR-TENANT-ID-HERE"             
$thumbprint = "YOUR-CERT-THUMBPRINT-HERE"   

# --- Input File Path ---
$sitelist = 'C:\temp\SPOSiteList.txt'

# --- Reset Label Configuration ---
$resetlabel = $false # Start with report-only mode
```

### 4. Run the Script

First, run in **report-only mode** to see what will be affected:

```powershell
.\Reset-Library-Labels.ps1
```

Review the output and log file. When ready, enable reset mode:

```powershell
# Edit script: Set $resetlabel = $true
.\Reset-Library-Labels.ps1
```

## ⚙️ Configuration Options

| Variable | Type | Default | Description |
|----------|------|---------|-------------|
| `$EnableVerbose` | Boolean | `False` | Enable detailed debug output including Graph API calls |
| `$LabelName` | String | `""` | **Empty string**: Reset ANY label found<br>**Specific name**: Only reset that label |
| `$appID` | String | Required | Entra Application (Client) ID |
| `$tenant` | String | Required | Tenant ID (GUID) |
| `$thumbprint` | String | Required | Certificate thumbprint |
| `$sitelist ` | String | Required | Path to text file with site URLs |
| `$ignoreListNames` | Array | `@('Site Assets', 'Site Pages')` | List names to skip |
| `$resetlabel` | Boolean | `False` | `False` = Report only<br>`True` = Reset labels |
| `$DelayBetweenLists` | Integer | `500` | Milliseconds delay between list operations |
| `$DelayBetweenSites` | Integer | `1000` | Milliseconds delay between site connections |
| `$MaxRetryAttempts ` | Integer | `5` | Maximum retry attempts for throttled requests |
| `$BaseRetryDelay` | Integer | `5000` | Base delay (ms) for exponential backoff |

## 📖 Usage Examples

### Example 1: Report on Specific Label

Find all libraries with a specific retention label:

```powershell
# Configuration
$LabelName = "Mark-Record-ReadOnly"
$resetlabel = $false

# Run script
.\Reset-Library-Labels.ps1
```

**Output:**
```
Starting to report on retention labels on 3 SharePoint sites...
Mode: REPORT-ONLY MODE - No changes will be made
Target Label: Mark-Record-ReadOnly

Processing site: https://contoso.sharepoint.com/sites/site1
Found 5 non-hidden lists with items in this site
  Processing list 1 of 5: Documents
    Current label: Mark-Record-ReadOnly
    Label found - WOULD RESET (report-only mode)
    [REPORT] Would reset and re-apply label: Mark-Record-ReadOnly
```

### Example 2: Reset ANY Label Found

Process all libraries regardless of which label they have:

```powershell
# Configuration
$LabelName = ""  # Empty = ANY label
$resetlabel = $true

# Run script
.\Reset-Library-Labels.ps1
```

This will reset and reapply whatever label is currently on each library.

### Example 3: Verbose Debug Mode

Enable detailed output to troubleshoot issues:

```powershell
# Configuration
$EnableVerbose = $true
$resetlabel = $false

# Run script
.\Reset-Library-Labels.ps1
```

**Additional output includes:**
- Graph API URLs
- Raw API responses
- Site and List GUIDs
- Compliance tag fields

## 🔍 How It Works

### Label Detection via Graph API

The script queries the `_ComplianceTag` field from list items using Microsoft Graph API:

```
GET https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items
  ?$expand=fields($select=id,_ComplianceTag,_ComplianceTagWrittenTime)
  &$top=1
```

This approach:
- ✅ Provides real-time label information
- ✅ Detects labels even if metadata is stale
- ✅ Works across all tenants and label configurations

### Reset Process

For each library with a label:

1. **Query**: Get current `_ComplianceTag` value via Graph API
2. **Reset**: Execute `Reset-PnPRetentionLabel -List $ListTitle`
3. **Reapply**: Execute `Set-PnPRetentionLabel -List $ListTitle -Label $FoundLabelName`
4. **Delay**: Wait configured milliseconds before next operation

### Throttling Protection

The script implements Microsoft's recommended throttling handling:

1. **Exponential Backoff**: Delay increases exponentially with each retry
2. **Retry-After Headers**: Honors server-provided retry delays
3. **Configurable Delays**: Proactive delays between operations
4. **Max Retry Limit**: Prevents infinite loops

## 📁 Output Files

### Console Output

Color-coded messages:
- 🟢 **Green**: Success operations
- 🟡 **Yellow**: Warnings and operations in progress
- 🔴 **Red**: Errors
- ⚪ **White**: Informational
- 🟣 **Magenta**: Debug output (when verbose enabled)
- ⚫ **Gray**: Verbose/detailed logs

### Log File

Location: `%TEMP%\Reset_Retention_Labels_<timestamp>_logfile.log`

Example:
```
[2026-01-14 10:30:15] [INFO] === RESET RETENTION LABELS SCRIPT STARTED ===
[2026-01-14 10:30:15] [INFO] Mode: REPORT-ONLY MODE
[2026-01-14 10:30:15] [INFO] Target Label: ANY label found
[2026-01-14 10:30:20] [SUCCESS] Successfully connected to site: https://...
[2026-01-14 10:30:25] [SUCCESS] Lists found with retention labels: 12
```

## ⚠️ Important Notes

### Label Propagation Delay

> **IMPORTANT**: Label application triggers background jobs in Microsoft 365 compliance services. These jobs update item metadata and sync changes to the compliance store and Graph endpoints. **It can take up to 24 hours** for labels to appear consistently in Graph queries.

**Best Practices:**
- Run report-only mode first to verify scope
- Wait 24-48 hours after label changes before re-running
- Use verbose mode to confirm label detection

### Throttling Considerations

SharePoint Online has throttling limits. This script includes protection, but for large tenants:

- Increase `` (e.g., `1000` ms)
- Increase `` (e.g., `2000` ms)
- Process sites in smaller batches

### Library vs List

This script processes **document libraries** and **lists**. It filters out:
- Hidden lists
- Empty lists (no items)
- Lists in the ignore list (configurable)

## 🐛 Troubleshooting

### Authentication Fails

**Error:** "Could not connect to site"

**Solutions:**
1. Verify certificate is installed with private key
2. Confirm thumbprint matches (no spaces)
3. Check App ID and Tenant ID
4. Verify API permissions granted and admin consent applied

### No Labels Found

**Error:** "No label set - skipping"

**Possible Causes:**
1. Label hasn't propagated yet (wait 24 hours)
2. Library is empty (no items)
3. Label is set at item level, not library level
4. Graph API permissions insufficient

**Debug:**
- Enable ` = $true`
- Check debug output for `_ComplianceTag` field value

### Throttling Errors

**Error:** "Throttled (HTTP 429)"

**Solutions:**
1. Script will automatically retry with exponential backoff
2. If persistent, increase delay values:
   ```powershell
   $DelayBetweenLists = 1000
   $DelayBetweenSites = 2000
   ```

## 📝 Version History

### Version 1.0 (January 2026)
- Initial release
- Support for specific label targeting or "any label"
- Graph API integration for label detection
- Comprehensive throttling protection
- Report-only and reset modes
- Detailed logging and verbose debug mode

## 🤝 Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 👤 Author

**Mike Lee**  
Microsoft

## 🙏 Acknowledgments

- PnP PowerShell Community
- Microsoft Graph API Team
- SharePoint Community

## 📞 Support

For issues, questions, or suggestions:
- Open an issue on GitHub
- Review existing issues for solutions
- Check the troubleshooting section above

---

**Note**: This tool is provided as-is without warranty. Always test in a non-production environment first.
