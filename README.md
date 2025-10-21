# SharePoint Online Record Item Unlocker

A PowerShell script to automatically unlock record-labeled items across multiple SharePoint Online sites. This script is designed to handle large-scale enterprise environments with built-in throttling protection and comprehensive logging.

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Understanding Compliance Flags](#understanding-compliance-flags)
- [Throttling Protection](#throttling-protection)
- [Logging](#logging)
- [Troubleshooting](#troubleshooting)
- [Best Practices](#best-practices)
- [License](#license)

## 🎯 Overview

This PowerShell script automates the process of unlocking SharePoint Online items that have been locked as records through compliance labels. It can process multiple sites and lists, handling thousands of items while respecting SharePoint Online throttling limits.

### What It Does

- Connects to multiple SharePoint Online sites
- Scans all non-hidden lists for locked record items
- Identifies items locked with compliance flags (7, 519)
- Unlocks items using the SharePoint REST API
- Provides detailed logging and progress tracking
- Handles SharePoint throttling automatically

## ✨ Features

- **Multi-Site Processing**: Process multiple SharePoint sites from a text file
- **Intelligent Detection**: Automatically identifies locked items using compliance flags
- **Item Title Display**: Shows both Item ID and Title for easy identification
- **Throttling Protection**: Implements Microsoft's recommended throttling protection strategies
  - Honors Retry-After headers
  - Exponential backoff for throttled requests
  - Configurable delays between operations
- **Comprehensive Logging**: Detailed logs with timestamps and severity levels
- **Progress Tracking**: Real-time progress updates during processing
- **Error Handling**: Robust error handling with automatic retry logic
- **Configurable Filtering**: Skip specific lists (e.g., Site Assets, Site Pages)
- **Verbose Mode**: Optional detailed output for debugging

## 📦 Prerequisites

### Required Software

- **PowerShell 5.1 or later** (Windows PowerShell or PowerShell 7+)
- **PnP.PowerShell Module** (SharePoint Patterns and Practices PowerShell)

### Required Permissions

- **SharePoint Online Administrator** or **Site Collection Administrator** rights
- **Azure AD App Registration** with:
  - Application permissions: `Sites.FullControl.All`
  - Certificate-based authentication configured

### Installation

1. **Install PnP.PowerShell Module**

   ```powershell
   Install-Module -Name PnP.PowerShell -Force -AllowClobber
   ```

2. **Set Up Azure AD App Registration**

   - Register an application in Azure AD
   - Grant `Sites.FullControl.All` API permissions
   - Upload a certificate for authentication
   - Note the **Application ID**, **Tenant ID**, and **Certificate Thumbprint**

3. **Download the Script**

   ```powershell
   git clone https://github.com/yourusername/spo-record-unlocker.git
   cd spo-record-unlocker
   ```

## ⚙️ Configuration

### 1. Update Script Variables

Edit the script and update these variables in the **USER CONFIGURATION** section:

```powershell
# --- Tenant and App Registration Details ---
$appID = "your-app-id-here"                    # Your Entra App ID
$tenant = "your-tenant-id-here"                # Your Tenant ID
$thumbprint = "your-certificate-thumbprint"    # Certificate thumbprint

# --- Input File Path ---
$sitelist = 'C:\temp\SPOSiteList.txt'         # Path to site URLs file

# --- Lists to Ignore ---
$ignoreListNames = @('Site Assets', 'Site Pages')  # Lists to skip
```

### 2. Create Site List File

Create a text file (e.g., `SPOSiteList.txt`) with one SharePoint site URL per line:

```
https://contoso.sharepoint.com/sites/Site1
https://contoso.sharepoint.com/sites/Site2
https://contoso.sharepoint.com/sites/Site3
```

### 3. Throttling Configuration (Optional)

Adjust throttling settings based on your environment size:

```powershell
# --- Throttling Protection Settings ---
$DelayBetweenItems = 100        # Delay between items (ms)
$DelayBetweenLists = 500        # Delay between lists (ms)
$DelayBetweenSites = 1000       # Delay between sites (ms)
$MaxRetryAttempts = 5           # Max retry attempts
$BaseRetryDelay = 5000          # Base retry delay (ms)
```

**For larger environments**: Increase delays (e.g., 200ms, 1000ms, 2000ms)  
**For smaller environments**: Decrease delays for faster processing

## 🚀 Usage

### Basic Usage

```powershell
.\unlock-listitem.ps1
```

### Verbose Mode (Debugging)

```powershell
.\unlock-listitem.ps1 -Verbose
```

Verbose mode displays:
- All compliance flag values
- Items that don't need unlocking
- Additional diagnostic information

### Example Output

```
Starting to process 3 SharePoint sites...
Log file location: C:\Users\user\AppData\Local\Temp\Unlock_Record_Items_2025-10-21_14-30-45_logfile.log

Throttling Protection: Enabled
  - Delay between items: 100 ms
  - Delay between lists: 500 ms
  - Delay between sites: 1000 ms
  - Max retry attempts: 5
  - Base retry delay: 5000 ms

Processing site: https://contoso.sharepoint.com/sites/Site1
Found 5 non-hidden lists with items in this site
  Processing list 1 of 5: Documents (250 items)
      → Item ID 42 ('Q3 Report.docx') is LOCKED (_ComplianceFlags: 519)
      ✓ Successfully unlocked record item ID 42 ('Q3 Report.docx') in list 'Documents'
    Summary: 1 locked record items found, 1 unlocked

=================================
PROCESSING COMPLETE
=================================
Sites processed: 3 of 3
Lists processed: 15
Total locked record items unlocked: 47
=================================
Log file saved to: C:\Users\user\AppData\Local\Temp\Unlock_Record_Items_2025-10-21_14-30-45_logfile.log
```

## 🔍 Understanding Compliance Flags

The script identifies locked items using the `_ComplianceFlags` field:

| Flag Value | Status | Action |
|------------|--------|--------|
| `null` or `0` | No compliance label | Skip (no action needed) |
| `7` | Locked for the first time | **Unlock** |
| `519` | Locked/relocked | **Unlock** |
| `771` | Already unlocked | Skip (no action needed) |
| Other | Unknown status | Log as warning |

### How It Works

1. The script retrieves the `_ComplianceFlags` value for each item
2. Items with flags `7` or `519` are identified as locked
3. The script calls the SharePoint REST API to unlock these items
4. Success/failure is logged for audit purposes

## 🛡️ Throttling Protection

This script implements [Microsoft's official throttling guidance](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online) to ensure reliable operation in large environments.

### Protection Features

1. **Retry-After Header Handling**
   - Automatically detects HTTP 429 and 503 responses
   - Honors the `Retry-After` header from SharePoint
   - Waits the exact time specified before retrying

2. **Exponential Backoff**
   - If no Retry-After header is provided
   - Uses exponential backoff: 5s, 10s, 20s, 40s, 80s
   - Prevents overwhelming SharePoint with retries

3. **Rate Limiting**
   - Configurable delays between items, lists, and sites
   - Prevents hitting rate limits in the first place
   - Recommended for environments with thousands of items

4. **Automatic Retry**
   - Up to 5 retry attempts by default
   - Only retries on throttling errors (429, 503)
   - Other errors fail immediately

### Adjusting for Your Environment

**Small environment (< 1,000 items)**:
```powershell
$DelayBetweenItems = 50
$DelayBetweenLists = 250
$DelayBetweenSites = 500
```

**Medium environment (1,000 - 10,000 items)**:
```powershell
$DelayBetweenItems = 100   # Default
$DelayBetweenLists = 500   # Default
$DelayBetweenSites = 1000  # Default
```

**Large environment (> 10,000 items)**:
```powershell
$DelayBetweenItems = 200
$DelayBetweenLists = 1000
$DelayBetweenSites = 2000
```

## 📊 Logging

### Log File Location

Logs are saved to: `%TEMP%\Unlock_Record_Items_YYYY-MM-DD_HH-mm-ss_logfile.log`

Example: `C:\Users\john\AppData\Local\Temp\Unlock_Record_Items_2025-10-21_14-30-45_logfile.log`

### Log Levels

- **INFO**: General information (connections, disconnections)
- **SUCCESS**: Successful operations (unlocked items)
- **WARNING**: Expected findings (locked items found, throttling events)
- **ERROR**: Actual errors (connection failures, API errors)
- **VERBOSE**: Detailed debugging information (only when -Verbose is used)

### Sample Log Entry

```
[2025-10-21 14:30:45] [INFO] === UNLOCK RECORD ITEMS SCRIPT STARTED ===
[2025-10-21 14:30:45] [INFO] Log file: C:\Users\john\AppData\Local\Temp\...
[2025-10-21 14:30:46] [SUCCESS] Successfully connected to site: https://contoso.sharepoint.com/sites/Site1
[2025-10-21 14:30:47] [WARNING] Item ID 42 ('Q3 Report.docx') is LOCKED (_ComplianceFlags: 519)
[2025-10-21 14:30:48] [SUCCESS] Successfully unlocked record item ID 42 ('Q3 Report.docx') in list 'Documents'
```

## 🔧 Troubleshooting

### Common Issues

#### 1. Authentication Failures

**Error**: "Connect-PnPOnline: Access denied"

**Solution**:
- Verify your App ID, Tenant ID, and Certificate Thumbprint are correct
- Ensure the certificate is installed in the Current User certificate store
- Verify the app has `Sites.FullControl.All` permissions
- Check that admin consent has been granted for the permissions

#### 2. Throttling (429/503 Errors)

**Error**: "Throttled (HTTP 429). Waiting X seconds..."

**Solution**:
- This is expected and handled automatically
- If it happens frequently, increase delay values
- Consider running the script during off-peak hours
- The script will automatically retry

#### 3. Module Not Found

**Error**: "The term 'Connect-PnPOnline' is not recognized..."

**Solution**:
```powershell
Install-Module -Name PnP.PowerShell -Force -AllowClobber
Import-Module PnP.PowerShell
```

#### 4. Permission Denied to Unlock Items

**Error**: "REST API unlock failed for item..."

**Solution**:
- Ensure the app has `Sites.FullControl.All` permissions
- Verify you're a Site Collection Administrator
- Some items may have additional restrictions that prevent unlocking

#### 5. Site List File Not Found

**Error**: "Input file not found: C:\temp\SPOSiteList.txt"

**Solution**:
- Create the file at the specified path
- Or update `$sitelist` variable to point to your file location

### Debugging Tips

1. **Use Verbose Mode**
   ```powershell
   .\unlock-listitem.ps1 -Verbose
   ```

2. **Check the Log File**
   - Log file path is shown at script start
   - Contains detailed information about all operations

3. **Test with a Single Site First**
   - Create a test site list with just one site
   - Verify the script works before processing all sites

4. **Uncomment Debug Output**
   - Find this line in the script:
     ```powershell
     # Write-HostAndLog "      → Item $itemDisplayName has no compliance label or is already unlocked..."
     ```
   - Remove the `#` to see all items being processed

## 📝 Best Practices

### Before Running

1. **Test in a non-production environment first**
2. **Backup your compliance labels and policies**
3. **Create a small test site list** to verify functionality
4. **Run with -Verbose** on a test site to understand the output
5. **Review the log file** after test runs

### During Execution

1. **Monitor the console output** for errors
2. **Don't interrupt the script** during processing
3. **Allow throttling protection to work** - be patient
4. **Watch for consistent errors** that might indicate a configuration issue

### After Running

1. **Review the log file** for any errors or warnings
2. **Verify items were unlocked** in SharePoint
3. **Document the results** (sites processed, items unlocked)
4. **Archive the log file** for compliance/audit purposes

### Performance Tips

1. **Run during off-peak hours** for better performance
2. **Adjust throttling delays** based on your environment
3. **Process sites in batches** if you have hundreds of sites
4. **Use filtering** to skip lists that don't contain records

## 🔐 Security Considerations

1. **Certificate Storage**: Store certificates securely in the certificate store
2. **App Registration**: Use least-privilege principle (only grant necessary permissions)
3. **Credential Management**: Never hardcode passwords; use certificate authentication
4. **Audit Logging**: Retain log files for compliance and audit purposes
5. **Access Control**: Limit who can run this script in production

## 📄 Script Information

- **Author**: Your Name
- **Version**: 1.0
- **Last Updated**: October 2025
- **PowerShell Version**: 5.1+
- **Dependencies**: PnP.PowerShell

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built using [PnP PowerShell](https://pnp.github.io/powershell/)
- Throttling protection based on [Microsoft's SharePoint Online throttling guidance](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online)
- Compliance policy APIs from [SharePoint REST API documentation](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)

## 📞 Support

If you encounter issues:

1. Check the [Troubleshooting](#troubleshooting) section
2. Review the log file for detailed error messages
3. Open an issue on GitHub with:
   - Error message
   - Log file excerpt (remove sensitive information)
   - PowerShell version (`$PSVersionTable.PSVersion`)
   - PnP.PowerShell version (`Get-Module PnP.PowerShell -ListAvailable`)

---

**⚠️ Disclaimer**: This script modifies SharePoint compliance settings. Test thoroughly in a non-production environment before using in production. The authors are not responsible for any data loss or compliance violations resulting from the use of this script.
