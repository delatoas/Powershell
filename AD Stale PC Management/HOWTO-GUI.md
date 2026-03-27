# Stale AD Computer Management GUI - User Guide

## Overview

The **Stale AD Computer Management GUI** (`Manage-StaleADComputers-GUI.ps1`) provides a Windows Forms-based graphical interface for managing stale computer objects in Active Directory. It implements a safe, 3-stage lifecycle approach:

1. **Identify** - Scan for workstations/servers inactive for a specified time
2. **Disable/Tag** - Disable active accounts and tag them for future deletion
3. **Delete** - Permanently remove computers that have been disabled for a safety period

A **Re-enable** action is also available to recover computers that were disabled in error, restoring their original AD description.

> **Scope:** This tool allows you to target **Windows Workstations**, **Windows Servers**, or both through configurable checkboxes, and optionally restrict the scan to a specific **OU (Search Base)**. By default, it targets Workstations across the entire domain.

---

## Prerequisites

- **PowerShell 5.1** or later
- **Active Directory PowerShell module** — automatically installed at launch if missing
  - Windows 10/11: Installed via `Add-WindowsCapability` (RSAT)
  - Windows Server: Installed via `Install-WindowsFeature`
  - Requires **Administrator** privileges for first-time auto-install
- Appropriate AD permissions to:
  - Query computer objects
  - Disable and re-enable computer accounts
  - Move objects between OUs (if using Staging OU)
  - Delete computer objects

---

## Launching the GUI

```powershell
.\Manage-StaleADComputers-GUI.ps1
```

The window will open centered on your screen. The **title bar** displays the current operator, for example:

```
Stale AD Computer Management  -  CONTOSO\admin.user
```

![GUI Main Window — Settings panel, filter bar, color-coded data grid, and activity log](screenshots/gui-main-window.png)

> **Tip:** Press **F5** at any time to trigger a new scan without reaching for the mouse.

---

## GUI Layout

The window is divided into five sections from top to bottom:

1. **Settings Panel** — configuration options
2. **Summary Bar** — live statistics from the last scan
3. **Filter Bar** — live search and progress indicator
4. **Data Grid** — scan results with color-coded rows
5. **Activity Log** — dark-themed, timestamped operation log

---

## Settings Panel

The settings panel is organized in five rows.

### Row 1: Core Settings

| Element | Description | Default | Min | Max |
|---------|-------------|---------|-----|-----|
| **Inactive Days** | Days since last logon to consider a computer stale | `365` | `30` | `730` |
| **Delete After Days** | Minimum days a computer must be disabled before deletion (safety buffer) | `30` | `7` | `365` |
| **Staging OU (DN)** | Distinguished Name of an OU to move disabled computers to. Leave empty to keep in place. | *(empty)* | — | — |

**Example Staging OU:**
```
OU=Disabled Computers,OU=Workstations,DC=contoso,DC=com
```

### Row 2: Logging and OS Filter

| Element | Description | Default |
|---------|-------------|---------|
| **Log Path** | Directory where activity log files are saved | `C:\Temp\StaleADComputerLogs` |
| **[...] Button** | Opens a folder browser to select the log path | — |
| **Include OS Types** | Checkboxes to include **Workstations (Win 10/11)** and/or **Servers** | `Workstations` checked |

### Row 3: Search Base (OU Scope)

| Element | Description | Default |
|---------|-------------|---------|
| **Search Base (OU)** | Restrict the scan to a specific OU. Leave blank to search the entire domain. | *(empty — entire domain)* |

**Example Search Base:**
```
OU=Workstations,OU=Computers,DC=contoso,DC=com
```

> Using a Search Base significantly speeds up scans in large environments by eliminating irrelevant OUs before the query reaches AD.

### Row 4: Export Folder

| Element | Description | Default |
|---------|-------------|---------|
| **Export Folder** | Directory where CSV reports and action logs are saved | `C:\Temp\StaleADComputerReports` |
| **[...] Button** | Opens a folder browser to select the export folder | — |

> **Note:** All CSV exports (scan reports, disable/tag actions, delete actions, re-enable actions) go to the **Export Folder**. Activity logs go to the **Log Path**.

### Row 5: Action Buttons

| Button | Color | Description |
|--------|-------|-------------|
| **Scan for Stale Computers** | Blue | Queries AD and populates the grid (also triggered by **F5**) |
| **Select All** | Default | Checks every row's Select checkbox |
| **Select None** | Default | Unchecks every row's Select checkbox |
| **Select Eligible** | Dark Red | Auto-selects computers ready for deletion (disabled past threshold, or externally disabled) |
| **Disable/Tag** | Yellow | Disables selected enabled computers; tags already-disabled untagged computers |
| **Delete Selected** | Red | Permanently deletes selected disabled computers (two-step confirmation) |
| **Re-enable** | Grey | Re-enables selected disabled computers and restores their original description |
| **Export to CSV** | Green | Exports current scan results to a timestamped CSV file |

> The Select Eligible, Disable/Tag, Delete, Re-enable, and Export buttons are disabled until a scan is performed.

---

## Summary Bar

Located directly below the settings panel, the summary bar displays real-time statistics from the last scan:

```
Total: 24  |  Enabled: 15  |  Disabled: 9  |  Marked: 6  |  Ready for Deletion: 3
```

| Stat | Description |
|------|-------------|
| **Total** | All stale computers returned by the scan |
| **Enabled** | Computers with active AD accounts (candidates for Disable/Tag) |
| **Disabled** | Computers with disabled AD accounts |
| **Marked** | Computers tagged by this script with a `DISABLED: YYYY-MM-DD` timestamp |
| **Ready for Deletion** | Tagged computers disabled for ≥ **Delete After Days** — actionable right now |

When a grid filter is active, a note is appended showing how many rows are currently visible:

```
Total: 24  |  ...  |  Ready for Deletion: 3  [filter: 'PC-FINANCE' — 4 shown]
```

---

## Filter Bar

Located between the summary bar and the data grid.

| Element | Description |
|---------|-------------|
| **Filter:** text box | Type any part of a computer name or OU path — the grid updates live |
| **x button** | Clears the filter and restores all rows |
| **Progress bar** | Marquee animation shown during scan, disable, delete, and re-enable operations |

> The filter only affects what is displayed — it does not change the scan results or the summary counts.

---

## Data Grid

The main data grid displays scan results. Rows are **color-coded by lifecycle stage**:

| Row Color | Meaning |
|-----------|---------|
| White | Enabled and stale — candidate for Disable/Tag |
| Orange | Disabled externally (no script tag) — eligible for deletion immediately |
| Yellow | Disabled and tagged by script — waiting out the safety period |
| Salmon / Pink | Disabled, tagged, and past the Delete After Days threshold — **ready to delete** |

### Columns

| Column | Description |
|--------|-------------|
| **Select** | Checkbox for bulk operations |
| **Name** | Computer name (NetBIOS name) |
| **Enabled** | `True` if the AD account is enabled, `False` if disabled |
| **LastLogonDate** | Date of last logon (`YYYY-MM-DD`) or `Never` |
| **DaysSinceLogon** | Number of days since last logon |
| **OperatingSystem** | Operating system string from AD |
| **MarkedForDeletion** | `True` if tagged by this script |
| **DisableDate** | Date the computer was tagged/disabled (`YYYY-MM-DD`) |
| **DaysSinceDisabled** | Days since the computer was disabled |
| **DistinguishedName** | Full AD path |

**Grid Features:**
- Click any column header to sort
- Multi-row selection supported via checkboxes
- Row colors update dynamically when **Delete After Days** is changed

### Right-Click Context Menu

Right-clicking any row shows a context menu for quick actions on that computer:

| Menu Item | Action |
|-----------|--------|
| **Disable / Tag** | Same as the Disable/Tag button for the clicked row |
| **Delete** | Same as Delete Selected for the clicked row |
| **Re-enable** | Same as Re-enable for the clicked row |
| **Copy Name** | Copies the computer name to the clipboard |
| **Copy Distinguished Name** | Copies the full DN to the clipboard |

> If the right-clicked row is not already checked, it will be selected and all other selections cleared before the menu appears. If it is already checked (part of a multi-selection), the existing selection is preserved.

---

## Activity Log

The bottom panel shows a dark-themed, read-only log. Every entry includes the **operator's identity**:

```
[2026-03-04 10:30:45] [Info] [CONTOSO\admin.user] Starting scan for stale computers
[2026-03-04 10:30:45] [Info] [CONTOSO\admin.user] Search Base: OU=Workstations,DC=contoso,DC=com
[2026-03-04 10:30:46] [Info] [CONTOSO\admin.user] Found 142 computers matching OS filter in AD
[2026-03-04 10:30:47] [Info] [CONTOSO\admin.user] Found 24 stale computers
[2026-03-04 10:30:47] [Info] [CONTOSO\admin.user] Scan completed.
```

**Log Levels:**
- `Info` — Normal operations
- `Warning` — Non-critical issues or skipped items
- `Error` — Operation failures

Logs are also written to a file in the configured **Log Path** (`StaleADComputers_YYYYMMDD.log`).

---

## Workflow Guide

### Step 1: Configure Settings

1. Set **Inactive Days** to your organization's policy (default: 365)
2. Set **Delete After Days** safety buffer (default: 30)
3. Optionally enter a **Staging OU** DN to quarantine disabled computers
4. Optionally enter a **Search Base** DN to limit the scan to a specific OU
5. Verify **Log Path** and **Export Folder** are accessible (use `...` browse buttons)

### Step 2: Scan for Stale Computers

1. Click **Scan for Stale Computers** — or press **F5**
2. The progress bar animates while the AD query runs
3. Results populate the grid with color-coded rows
4. Review the summary bar for a quick count breakdown

### Step 3: Review and Export

1. Use the **Filter** box to find specific computers or OUs
2. Click **Export to CSV** to save results for stakeholder review
3. When prompted "Open export folder?" click **Yes** to open it in Explorer
4. Files are named: `StaleComputers_Report_YYYYMMDD_HHMMSS.csv`

### Step 4: Disable / Tag Computers

1. Select computers using checkboxes, **Select All**, or the Filter + Select All combination
2. Click **Disable/Tag**
3. Review the confirmation dialog showing exactly which computers will be disabled vs. tagged
4. Click **Yes** to proceed

![Disable/Tag confirmation dialog](screenshots/confirm-disable-dialog.png)

**What happens:**
- **Enabled stale computers** → Disabled in AD, description updated with `DISABLED: YYYY-MM-DD`, optionally moved to Staging OU
- **Already-disabled untagged computers** → Description updated with deletion tag (not re-disabled)

The action results are exported to `DisableTagActions_YYYYMMDD_HHMMSS.csv` in the Export Folder.

### Step 5: Delete Computers (After Safety Period)

1. Wait for the **Delete After Days** period to pass
2. Press **F5** to re-scan
3. Click **Select Eligible** — this auto-checks:
   - Computers tagged by this script that are past the Delete After Days threshold
   - Computers disabled externally (no tag) — eligible immediately
4. Review the grid (salmon rows = eligible)
5. Click **Delete Selected**
6. First confirmation: review the list, click **OK**
7. Second confirmation: type `DELETE` (case-sensitive) and click **Confirm**

![Delete confirmation — type DELETE to proceed](screenshots/confirm-delete-dialog.png)

The action results are exported to `DeleteActions_YYYYMMDD_HHMMSS.csv` in the Export Folder.

### Step 6: Re-enable a Computer (Recovery)

If a computer was incorrectly disabled by this script:

1. Run a scan (the computer will appear as a disabled row)
2. Check its **Select** checkbox (or right-click → **Re-enable**)
3. Click **Re-enable**
4. Review the confirmation dialog listing the computers to restore
5. Click **Yes**

**What happens:**
- The computer account is re-enabled in AD
- The original description is restored from the embedded tag format:
  `DISABLED: 2026-02-17 | Original: Finance Laptop | Stale 365+ days`
  → description restored to: `Finance Laptop`
- If the description format cannot be parsed, the current description is left unchanged
- Results are exported to `ReEnableActions_YYYYMMDD_HHMMSS.csv` in the Export Folder

---

## Deletion Logic: Managed vs. Unmanaged

The **Delete Selected** button handles objects differently based on their state:

1. **Managed Objects (Tagged by Script)**
   - Identified by a `DISABLED: YYYY-MM-DD` tag in the description
   - Behavior: Enforces the **Delete After Days** threshold
   - Purpose: Ensures a cooling-off period for computers disabled by this tool

2. **Unmanaged Objects (Externally Disabled)**
   - Already disabled in AD but missing the script's timestamp tag
   - Behavior: Can be deleted immediately, bypassing the threshold
   - Purpose: Assumes manual review has already occurred for these objects

3. **Safety Lock**
   - The script will **refuse to delete any enabled account**, even if stale. It must be disabled first.

---

## Safety Features

### Multi-Layer Confirmations

| Action | Confirmation |
|--------|-------------|
| Disable/Tag | Yes/No dialog listing affected computers |
| Delete | OK/Cancel dialog + separate "type DELETE" input |
| Re-enable | Yes/No dialog listing affected computers |

### Audit Trail

All operations capture **operator identity** (`DOMAIN\Username`) in:

- **On-screen** Activity Log (every line)
- **Title bar** shows current operator
- **Log files** in the Log Path: `StaleADComputers_YYYYMMDD.log`
- **Action CSVs** in the Export Folder — every record has a `PerformedBy` field

#### Action CSV Fields

| Field | Description |
|-------|-------------|
| ComputerName | Computer acted upon |
| Action | `Disabled`, `Tagged`, `Deleted`, or `Re-enabled` |
| Status | `Success` or `Failed` |
| **PerformedBy** | `DOMAIN\Username` of the operator |
| Timestamp | Date/time of the action |

### Description Tracking

When a computer is disabled by this script, its AD Description field is updated:
```
DISABLED: 2026-02-17 | Original: Finance Laptop | Stale 365+ days
```

This encodes:
- When the computer was disabled
- The original description (for Re-enable restoration)
- Why it was disabled (stale threshold or external tag)

When Re-enable is used, the description is restored to `Finance Laptop`.

### Staging OU (Optional)

If a **Staging OU** DN is configured, disabled computers are physically moved to that OU, making it easy to:
- Isolate disabled accounts from active OU structure
- Review before permanent deletion
- Restore using Re-enable if disabled in error

---

## Parameter Reference

### Numeric Controls

| Setting | Default | Min | Max | Description |
|---------|---------|-----|-----|-------------|
| Inactive Days | 365 | 30 | 730 | Days of inactivity before flagged as stale |
| Delete After Days | 30 | 7 | 365 | Safety buffer days before deletion is allowed |

### Text Fields

| Field | Default | Description |
|-------|---------|-------------|
| Staging OU (DN) | *(empty)* | Target OU DN for disabled computers |
| Log Path | `C:\Temp\StaleADComputerLogs` | Directory for activity log files |
| Search Base (OU) | *(empty — entire domain)* | Restrict scan to a specific OU |
| Export Folder | `C:\Temp\StaleADComputerReports` | Directory for all CSV exports |

---

## Troubleshooting

### Module Installation Failures

If the automatic RSAT installation fails:
- Ensure PowerShell is running **as Administrator**
- Windows 10/11: Verify internet access (required for RSAT via `Add-WindowsCapability`)
- Windows Server: Verify `ServerManager` module is available
- Manual install: `Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"`

### "Access Denied" Errors

Ensure your account has:
- Read access to query computers
- Write access to modify computer properties (disable, update description)
- Delete access to remove computer objects
- Move access if using Staging OU

### No Computers Found

- Verify the **Inactive Days** threshold isn't too low
- If using a **Search Base**, confirm the OU DN is correct and contains computer objects
- Ensure at least one OS type checkbox is checked
- Ensure you have read permissions across the relevant OUs

### Search Base Errors

- Verify the Distinguished Name format: `OU=Name,DC=domain,DC=com`
- Confirm the OU exists in AD (`Get-ADOrganizationalUnit -Identity "OU=..."`)
- The field is case-insensitive but the DN structure must be valid

### Staging OU Errors

- Verify the Distinguished Name format is correct
- Ensure the OU exists in AD
- Confirm you have move permissions to the target OU

### Re-enable Does Not Restore Original Description

This occurs when the computer's current description doesn't match the expected format:
```
DISABLED: YYYY-MM-DD | Original: <desc> | <reason>
```
In this case the description is left as-is. You can manually edit it in AD Users and Computers after re-enabling.

---

## File Outputs

| File Pattern | Saved To | Content |
|---|---|---|
| `StaleADComputers_YYYYMMDD.log` | Log Path | Session activity log with operator identity |
| `StaleComputers_Report_YYYYMMDD_HHMMSS.csv` | Export Folder | Full scan results |
| `DisableTagActions_YYYYMMDD_HHMMSS.csv` | Export Folder | Disable/tag results with `PerformedBy` |
| `DeleteActions_YYYYMMDD_HHMMSS.csv` | Export Folder | Delete results with `PerformedBy` |
| `ReEnableActions_YYYYMMDD_HHMMSS.csv` | Export Folder | Re-enable results with `PerformedBy` |

All timestamps use `YYYYMMDD` for daily logs and `YYYYMMDD_HHMMSS` for per-operation CSVs.
