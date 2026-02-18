# Stale AD Computer Management GUI - User Guide

## Overview

The **Stale AD Computer Management GUI** (`Manage-StaleADComputers-GUI.ps1`) provides a Windows Forms-based graphical interface for managing stale computer objects in Active Directory. It implements a safe, 3-stage lifecycle approach:

1. **Identify** - Scan for workstations inactive for a specified number of days
2. **Disable/Tag** - Disable active accounts and tag them for future deletion
3. **Delete** - Permanently remove computers that have been disabled for a safety period

> **Scope:** This tool automatically targets only **Windows 10** and **Windows 11** workstations. Windows Server systems are excluded.

---

## Prerequisites

- **PowerShell 5.1** or later
- **ActiveDirectory** PowerShell module
- Appropriate AD permissions to:
  - Query computer objects
  - Disable computer accounts
  - Move objects between OUs (if using Staging OU)
  - Delete computer objects

---

## Launching the GUI

```powershell
.\Manage-StaleADComputers-GUI.ps1
```

The window will open centered on your screen with the SMPA logo in the title bar.

---

## GUI Elements

### Settings Panel

The top section contains configuration options organized in three rows.

#### Row 1: Core Settings

| Element | Description | Default | Min | Max |
|---------|-------------|---------|-----|-----|
| **Inactive Days** | Number of days since last logon to consider a computer "stale" | `365` | `30` | `730` |
| **Delete After Days** | Minimum days a computer must be disabled before it can be deleted (safety buffer) | `30` | `7` | `365` |
| **Staging OU (DN)** | Distinguished Name of an OU to move disabled computers to. Leave empty to keep computers in their current location. | *(empty)* | N/A | N/A |

**Example Staging OU:**
```
OU=Disabled Computers,OU=Workstations,DC=contoso,DC=com
```

#### Row 2: File Paths

| Element | Description | Default |
|---------|-------------|---------|
| **Log Path** | Directory where operation log files are saved | `C:\Temp\StaleADComputerLogs` |
| **Export Folder** | Directory where CSV reports are saved | `C:\Temp\StaleADComputerReports` |
| **[...] Button** | Opens a folder browser to select the export folder | N/A |

#### Row 3: Action Buttons

| Button | Color | Description |
|--------|-------|-------------|
| **Scan for Stale Computers** | Blue | Queries AD for workstations matching the inactive days threshold |
| **Select All** | Default | Checks the "Select" checkbox for all rows in the grid |
| **Select None** | Default | Unchecks the "Select" checkbox for all rows |
| **Disable/Tag** | Yellow | Disables selected enabled computers and/or tags already-disabled computers for deletion |
| **Delete Selected** | Red | Permanently deletes selected disabled computers (requires confirmation) |
| **Export to CSV** | Green | Exports the current scan results to a timestamped CSV file |

> **Note:** The Disable/Tag, Delete, and Export buttons are disabled until a scan is performed.

---

### Summary Bar

Located below the settings panel, displays real-time statistics:

```
Total: X | Enabled: Y | Disabled: Z | Marked for Deletion: W
```

- **Total** - Number of stale computers found
- **Enabled** - Computers with active AD accounts
- **Disabled** - Computers with disabled AD accounts
- **Marked for Deletion** - Computers tagged by this script with a `DISABLED: YYYY-MM-DD` timestamp

---

### Data Grid

The main data grid displays scan results with the following columns:

| Column | Description |
|--------|-------------|
| **Select** | Checkbox to select computers for bulk operations |
| **Name** | Computer name (NetBIOS name) |
| **Enabled** | `True` if the AD account is enabled, `False` if disabled |
| **LastLogonDate** | Date of last logon (YYYY-MM-DD format) or "Never" |
| **DaysSinceLogon** | Number of days since last logon |
| **OperatingSystem** | Operating system (Windows 10, Windows 11, etc.) |
| **MarkedForDeletion** | `True` if tagged by this script for deletion |
| **DisableDate** | Date when the computer was disabled (if tagged) |
| **DaysSinceDisabled** | Days since the computer was disabled |
| **DistinguishedName** | Full AD path of the computer object |

**Grid Features:**
- Alternating row colors for readability
- Click column headers to sort
- Multi-row selection supported
- Checkbox in "Select" column for bulk operations

---

### Activity Log

The bottom panel shows a dark-themed, read-only log with timestamped entries:

```
[2026-02-17 10:30:45] [Info] Starting scan for stale computers
[2026-02-17 10:30:46] [Warning] Skipping PC001 - only disabled for 15 days
[2026-02-17 10:30:47] [Error] Failed to delete PC002: Access denied
```

**Log Levels:**
- **Info** - Normal operations
- **Warning** - Non-critical issues or skipped items
- **Error** - Operation failures

Logs are also saved to files in the configured **Log Path** directory.

---

## Workflow Guide

### Step 1: Configure Settings

1. Set **Inactive Days** to your organization's policy (default: 365 days = 1 year)
2. Set **Delete After Days** safety buffer (default: 30 days)
3. Optionally specify a **Staging OU** to quarantine disabled computers
4. Verify **Log Path** and **Export Folder** are accessible

### Step 2: Scan for Stale Computers

1. Click **Scan for Stale Computers**
2. Wait for the scan to complete (progress shown in Activity Log)
3. Review the summary bar and data grid results

### Step 3: Review and Export

1. Click **Export to CSV** to save results for review
2. Share the report with stakeholders if needed
3. Files are named: `StaleComputers_Report_YYYYMMDD_HHMMSS.csv`

### Step 4: Disable/Tag Computers

1. Select computers using checkboxes or **Select All**
2. Click **Disable/Tag**
3. Review the confirmation dialog showing:
   - Number of computers to disable (currently enabled)
   - Number of computers to tag (already disabled externally)
4. Click **Yes** to proceed

**What happens:**
- **Enabled computers** → Disabled, description updated with `DISABLED: YYYY-MM-DD`, optionally moved to Staging OU
- **Already disabled (untagged) computers** → Description updated with deletion tag only

### Step 5: Delete Computers (After Safety Period)

1. Wait for the **Delete After Days** period to pass
2. Run another scan
3. Select computers that are ready for deletion
4. Click **Delete Selected**
5. First confirmation dialog: Review the list, click **OK**
6. Second confirmation: Type `DELETE` (case-sensitive) and click **Confirm**

**Safety checks:**
- Only disabled computers can be deleted
- Tagged computers must meet the minimum days threshold
- Externally-disabled computers (no tag) can be deleted immediately

---

## Safety Features

### Multi-Layer Confirmation

| Action | Confirmation Required |
|--------|----------------------|
| Disable/Tag | Yes/No dialog |
| Delete | OK/Cancel dialog + type "DELETE" |

### Audit Trail

All operations are logged to:
- **On-screen** Activity Log
- **Log files** in the configured Log Path:
  - `StaleADComputers_YYYYMMDD_HHMMSS.log` - Session log
  - `DisableTagActions_YYYYMMDD_HHMMSS.csv` - Disable/tag results
  - `DeleteActions_YYYYMMDD_HHMMSS.csv` - Delete results

### Description Tracking

Disabled computers are tagged in their AD Description field:
```
DISABLED: 2026-02-17 | Original: Finance PC | Stale 365+ days
```

This allows:
- Tracking when the computer was disabled
- Preserving the original description
- Calculating days since disabled for deletion safety

### Staging OU (Optional)

If configured, disabled computers are moved to a quarantine OU, making it easy to:
- Isolate disabled accounts
- Review before permanent deletion
- Restore if disabled in error

---

## Parameter Reference

### Numeric Controls

| Parameter | Default | Minimum | Maximum | Description |
|-----------|---------|---------|---------|-------------|
| Inactive Days | 365 | 30 | 730 | Days of inactivity to flag as stale |
| Delete After Days | 30 | 7 | 365 | Safety buffer before deletion allowed |

### Text Fields

| Field | Default | Description |
|-------|---------|-------------|
| Staging OU (DN) | *(empty)* | Target OU for disabled computers |
| Log Path | `C:\Temp\StaleADComputerLogs` | Directory for log files |
| Export Folder | `C:\Temp\StaleADComputerReports` | Directory for CSV exports |

---

## Troubleshooting

### "Access Denied" Errors

Ensure your account has:
- Read access to query computers
- Write access to modify computer properties (disable, update description)
- Delete access to remove computer objects
- Move access if using Staging OU

### No Computers Found

- Verify the **Inactive Days** threshold isn't too low
- Check that Windows 10/11 workstations exist in your AD
- Ensure you have read permissions across relevant OUs

### Staging OU Errors

- Verify the Distinguished Name format is correct
- Ensure the OU exists in AD
- Confirm you have move permissions to the target OU

---

## File Outputs

| File Pattern | Content |
|--------------|---------|
| `StaleComputers_Report_*.csv` | Scan results export |
| `StaleADComputers_*.log` | Session activity log |
| `DisableTagActions_*.csv` | Disable/tag operation results |
| `DeleteActions_*.csv` | Delete operation results |

All timestamps use format: `YYYYMMDD_HHMMSS`
