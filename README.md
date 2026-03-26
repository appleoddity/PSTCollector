# PSTCollector

PowerShell scripts to find, collect, and assist in the upload of .PST files to Microsoft 365.

**Developed by:** Appleoddity

## Overview

PSTCollector automates the discovery, collection, and migration of Outlook PST files from domain-joined workstations and network file shares to Microsoft 365 Exchange Online. The tool generates a CSV mapping file compatible with the Microsoft 365 PST Import Service.

## Requirements

- Windows 7 or newer on all target systems
- PowerShell 5.0+ on the CollectorMaster system
- Domain-joined computers
- Domain Admin (or equivalent) privileges for CollectorMaster
- [PsExec](https://docs.microsoft.com/en-us/sysinternals/downloads/psexec) from Sysinternals
- Active Directory users must have accurate email addresses matching their Microsoft 365 accounts

## Components

| Component | Description |
|-----------|-------------|
| `CollectorMaster.ps1` | Central orchestration script run from an admin workstation |
| `CollectorAgent.ps1` | Agent pushed to workstations to scan, collect, and remove PST files |
| `unloadPST.vbs` | VBScript to remove PST files from Outlook before deletion |

## Quick Start

### 0. Pre-Migration Preparation

Before starting collection, deploy Group Policy to:
- **Disable PST file creation** - Prevents new PSTs during migration
- **Disable PST file growth** - Prevents additions to existing PSTs

Users can still remove items from PSTs, so notify them in advance to clean up before collection.

### 1. Setup File Shares

**Scripts Share** (read-only for domain):
```
\\server\PSTCollector\
├── CollectorAgent.ps1
├── CollectorMaster.ps1
└── psexec.exe
```

**Collection Share** with special permissions:
- **Share Permissions:** Everyone → Full Control
- **NTFS Permissions:**
  - Administrators, SYSTEM → Full Control (This folder, subfolders, files)
  - CREATOR OWNER → Full Control (Subfolders and files only)
  - Authenticated Users → Special (This folder only):
    - Traverse folder / execute file
    - List folder / read data
    - Create folders / append data

### 2. Run Collection

```powershell
# Step 1: Find all PST files
.\CollectorMaster.ps1 -Mode FIND -JobName "Migration2024" `
    -Locations "OU=Workstations,DC=corp,DC=local","\\fileserver\home$" `
    -CollectPath "\\server\PSTCollection"

# Step 2: Collect PST files to central location
.\CollectorMaster.ps1 -Mode COLLECT -JobName "Migration2024" `
    -Locations "OU=Workstations,DC=corp,DC=local","\\fileserver\home$" `
    -CollectPath "\\server\PSTCollection"

# Step 3: Remove source PST files (after import verification)
.\CollectorMaster.ps1 -Mode REMOVE -JobName "Migration2024" `
    -Locations "OU=Workstations,DC=corp,DC=local","\\fileserver\home$" `
    -CollectPath "\\server\PSTCollection"
```

**Note:** Because computers may be offline, run each mode multiple times until all locations complete. The tool intelligently resumes progress and skips already-processed items.

## CollectorMaster Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Mode` | Yes | `FIND`, `COLLECT`, or `REMOVE` |
| `-JobName` | Yes | Unique job identifier (must be consistent across runs) |
| `-Locations` | Yes | Comma-separated OUs or UNC paths |
| `-CollectPath` | Yes | UNC path to collection share |
| `-ConfigPath` | No | Master config location (default: `C:\PSTCollector`) |
| `-ForceRestart` | No | Wipe existing config and start fresh |
| `-Noping` | No | Skip ping checks for offline detection |
| `-ThrottleLimit` | No | Max concurrent jobs (default: 25) |
| `-NoSkipCommon` | No | Scan all folders including Program Files, Windows, etc. |
| `-IsArchive` | No | Import to archive mailbox instead of primary (default: True) |
| `-IPG` | No | RoboCopy interpacket gap in ms for bandwidth throttling (default: 1) |
| `-EnabledOnly` | No | Only target enabled computers in AD when processing OUs (default: all computers) |

## Output Files

After a successful COLLECT, you'll have:

```
\\server\PSTCollection\
├── MASTER-Migration2024.xml    # Full job status and file inventory
├── MASTER-Migration2024.csv    # Microsoft 365 import mapping file
├── MASTER-Migration2024.log    # Execution log
└── [computer/location folders with PST files]
    ├── Migration2024.xml
    ├── Migration2024.log
```

## Status Values

Use these in the XML to control behavior:

| Status | Description |
|--------|-------------|
| `Incomplete` | Job started but not finished |
| `Found` | PST files discovered |
| `Collected` | Files copied to collection share |
| `Removed` | Source files deleted |
| `Void` | Skip this location/file |
| `Offline` | Computer was unreachable |
| `*Error` | Error occurred (FindError, CollectError, RemoveError) |

---

## Importing PST Files to Exchange Online

After collection, use AzCopy to upload PST files to Microsoft 365, then create an import job using the generated CSV file.

### Recommended: Enable Archive Mailboxes

PSTCollector's CSV export is configured by default to import PST files into users' **archive mailboxes** (`IsArchive = True`). Archive mailboxes provide additional storage without affecting primary mailbox quotas and are the recommended destination for historical PST data.

Before importing, ensure archive mailboxes are enabled for target users:

```powershell
# Connect to Exchange Online
Connect-ExchangeOnline

# Enable archive mailbox for all users who don't have one
Get-Mailbox -Filter {ArchiveStatus -Eq "None" -AND RecipientTypeDetails -Eq "UserMailbox"} -ResultSize Unlimited | Enable-Mailbox -Archive

# Verify archive status for a specific user
Get-Mailbox -Identity user@domain.com | Select-Object DisplayName, ArchiveStatus, ArchiveName
```

**Note:** To import to primary mailboxes instead of archives, run CollectorMaster with `-IsArchive $false`.

### Step 1: Get Azure Storage Credentials

1. Go to [Microsoft Purview Portal](https://purview.microsoft.com)
2. Navigate to **Data lifecycle management** → **Import**
3. Click **+ New import job** → **Upload your data**
4. Copy the **SAS URL** (valid for 10 days)

### Step 2: Install AzCopy

```powershell
# Download AzCopy
Invoke-WebRequest -Uri "https://aka.ms/downloadazcopy-v10-windows" -OutFile "azcopy.zip"
Expand-Archive -Path "azcopy.zip" -DestinationPath "C:\AzCopy"

# Find the extracted folder and add to PATH
$azcopyFolder = Get-ChildItem "C:\AzCopy" -Directory | Select-Object -First 1
$env:PATH += ";$($azcopyFolder.FullName)"
```

### Step 3: Upload PST Files

The destination path must be inserted into the SAS URL **before** the query string (the `?` and everything after it).

```powershell
# Your SAS URL from Step 1
$SasUrl = "https://[account].blob.core.windows.net/ingestiondata?skoid=...&sig=..."

# Build the destination URL by inserting the path BEFORE the query string
# Original:  https://account.blob.core.windows.net/ingestiondata?skoid=...
# With path: https://account.blob.core.windows.net/ingestiondata/pstcollector?skoid=...

# Split the SAS URL at the query string
$baseUrl = $SasUrl.Split('?')[0]
$sasToken = $SasUrl.Split('?')[1]

# Upload entire collection (throttled to 10 Mbps)
$destUrl = "$baseUrl/pstcollector?$sasToken"
azcopy copy "\\server\PSTCollection\*" $destUrl --recursive --cap-mbps 10

# Or upload specific job data
$destUrl = "$baseUrl/pstcollector/fileserver/home`$?$sasToken"
azcopy copy "\\server\PSTCollection\fileserver\home$\*" $destUrl --recursive --cap-mbps 10

$destUrl = "$baseUrl/pstcollector/computername?$sasToken"
azcopy copy "\\server\PSTCollection\computername\*" $destUrl --recursive --cap-mbps 10
```

**Bandwidth Throttling:** The `--cap-mbps 10` switch limits upload speed to 10 Mbps to avoid saturating network links. Adjust this value based on available bandwidth (e.g., `--cap-mbps 50` for faster uploads during off-hours). Remove the switch entirely to upload at maximum speed.

**Resuming Interrupted Uploads:** If the upload is interrupted or fails, you can resume it:

```powershell
# List recent jobs
azcopy jobs list

# Resume a failed/canceled job (use job ID from list)
azcopy jobs resume "<job-id>" --destination-sas="$sasToken"
```

**Important:** The folder structure in Azure must match the `FilePath` column in the CSV. PSTCollector generates paths starting with `pstcollector/...`

### Step 4: Create Import Job

1. Return to the import job in Compliance Portal
2. Click **I'm done uploading my files**
3. Select **I have access to the mapping file**
4. Upload `MASTER-[JobName].csv` from `\\server\PSTCollection\`
5. Click **Validate** and fix any errors
6. Submit the job and complete analysis
7. Begin the import to Office 365

### CSV File Format

The generated CSV follows Microsoft's PST Import format:

| Column | Description | Example |
|--------|-------------|---------|
| `Workload` | Always "Exchange" | Exchange |
| `FilePath` | Azure blob path (no filename) | pstcollector/server/users/jsmith |
| `Name` | PST filename | archive.pst |
| `Mailbox` | Target email address | jsmith@company.com |
| `IsArchive` | Import to archive mailbox | TRUE |
| `TargetRootFolder` | Destination folder | /ImportedPST/archive |
| `ContentCodePage` | (Optional) Language code | |
| `SPFileContainer` | (SharePoint only) | |
| `SPManifestContainer` | (SharePoint only) | |
| `SPSiteUrl` | (SharePoint only) | |

### Mailbox Mapping

PSTCollector determines the target mailbox by:
1. Reading the file owner from the PST's ACL
2. Looking up that user's `EmailAddress` attribute in Active Directory
3. If lookup fails, defaulting to the Administrator account's email address

**Important:** Review `MASTER-[JobName].xml` before the final COLLECT to correct any mapping errors by updating the `Owner` attribute on problem files.

### Microsoft 365 Import Limits

- **Maximum 250 PST files per import job** - Split your CSV into multiple files if needed
- Large PSTs (>20GB) significantly increase import time
- Import jobs run in the background and may take days for large datasets

### Troubleshooting Import Issues

**CSV validation errors:**
- Ensure email addresses in `Mailbox` column are valid Microsoft 365 accounts
- Verify `FilePath` matches the actual Azure blob structure
- Check for Unicode characters in filenames (PSTCollector handles these properly)

**Upload failures:**
- Verify SAS URL hasn't expired (10-day validity)
- Check network connectivity to Azure
- Ensure sufficient Azure storage quota

**Import job stuck:**
- Large PST files can take hours/days to process
- Check import job status in Compliance Portal
- Contact Microsoft Support for jobs stuck >72 hours

---

## Removing PST Files from Outlook

**CRITICAL:** Before running REMOVE mode, ensure:
1. PST files are successfully imported to Microsoft 365
2. Users have verified their imported mail
3. PST files are detached from Outlook

### Using unloadPST.vbs

Deploy via login script or Group Policy:

```batch
cscript \\server\PSTCollector\unloadPST.vbs
```

This script:
- Connects to running Outlook instance
- Removes all PST data stores (except SharePoint Lists and Internet Calendar Subscriptions)
- Safe to run multiple times

---

## Unicode Support

PSTCollector properly handles international characters in:
- File names (e.g., `données.pst`, `日本語.pst`)  
- Folder paths
- User names

Key implementation details:
- Uses `XmlDocument.Load()` instead of `Get-Content` for proper UTF-8 XML handling
- RoboCopy `/UNILOG:` parameter for Unicode path support
- `-LiteralPath` for PowerShell file operations

---

## Troubleshooting

### AzCopy Upload Issues

AzCopy creates log and plan files for every job. By default, these are stored in `%USERPROFILE%\.azcopy`.

**View job history and status:**
```powershell
# List all jobs
azcopy jobs list

# Show details for a specific job
azcopy jobs show "<job-id>"

# Show only failed transfers
azcopy jobs show "<job-id>" --with-status=Failed
```

**Review logs for errors:**
```powershell
# Find upload failures in a log file
Select-String UPLOADFAILED "$env:USERPROFILE\.azcopy\<job-id>.log"
```

**Resume a failed or canceled job:**
```powershell
# Resume with SAS token (required since tokens aren't persisted)
azcopy jobs resume "<job-id>" --destination-sas="<sas-token>"
```

**Clean up old job files:**
```powershell
# Remove all plan and log files
azcopy jobs clean

# Remove files for a specific job
azcopy jobs rm "<job-id>"
```

For more details, see [Microsoft's AzCopy troubleshooting guide](https://learn.microsoft.com/en-us/azure/storage/common/storage-use-azcopy-configure).

### Common Issues

**"Cannot find path" errors with special characters:**
- Ensure you're running the latest version with Unicode fixes
- Delete corrupted XML files and restart with `-ForceRestart`

**PSExec access denied:**
- Verify Domain Admin privileges
- Check Windows Firewall settings on target computers
- Ensure admin shares (C$) are accessible

**Collector stuck or slow:**
- Reduce `-ThrottleLimit` for network-constrained environments
- Use `-Noping` only if ping is blocked by policy

**Jobs never complete:**
- Check XML file for locations stuck in `Incomplete` or `Error` status
- Set problem locations to `Void` to skip them

### Log Files

- Master log: `C:\PSTCollector\MASTER-[JobName].log`
- Agent logs: `C:\PSTCollector\[JobName].log` on each workstation
- Collection logs: `\\server\PSTCollection\[location]\[JobName].log`

### Working with XML Configuration Files

The XML files track all job progress and can be edited to control behavior:
- Use **Notepad++** or similar for easier viewing/editing of large XML files
- Set any location or file status to `Void` to skip it permanently
- Review `Owner` attributes to fix mailbox mapping before final collection
- Check for stuck `Incomplete` or `Error` statuses that need attention

---

## Known Limitations

### No VSS Snapshots for UNC Paths

CollectorAgent uses Volume Shadow Copy Service (VSS) to access locked PST files on local drives, but **VSS is not available for UNC paths**. When collecting from network shares (`\\server\share`), open PST files may fail to transfer with "file in use" errors.

**Workarounds:**

1. **Target the file server directly (recommended for Windows servers):**
   Instead of using UNC paths, add the file server to an OU target and let the agent run locally on the server. This enables VSS snapshots for locked files.
   ```powershell
   # Instead of:
   -Locations "\\fileserver\home$"
   
   # Use:
   -Locations "OU=FileServers,DC=corp,DC=local"
   ```

2. **Schedule collection during off-hours:**
   Run COLLECT mode when users are logged off and PST files are closed.

3. **Use server-side snapshot tools:**
   For non-Windows file servers (NetApp, EMC, NAS devices), use the vendor's snapshot capabilities to create a point-in-time copy, then collect from the snapshot path.

4. **Manually close PST files:**
   Deploy `unloadPST.vbs` to users via Group Policy logoff script, or use Outlook GPO settings to prevent PST file usage.

---

## Contributing

Pull requests welcome.