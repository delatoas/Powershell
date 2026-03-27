# GUI Improvement Plan — Manage-StaleADComputers-GUI.ps1

Branch: `feature/gui-improvements`

Work through each phase sequentially. Test after completing each phase before moving to the next.

---

## Phase 1 — Bug Fixes (High Priority)

### 1.1 — Action CSVs exported to wrong folder
**Problem:** `DisableTagActions_*.csv` and `DeleteActions_*.csv` are saved to `$script:LogPath` (lines 880, 1003), but the report CSV goes to `$txtExportPath.Text`. Inconsistent — admins may not find their action logs.

**Fix:** In both the Disable/Tag and Delete click handlers, change the `$actionFile` path from `Join-Path $script:LogPath ...` to `Join-Path $txtExportPath.Text.Trim() ...`. Add the same directory-creation guard that the Export button already uses.

**Affected lines:** 879–882 (Disable handler), 1002–1005 (Delete handler)

---

### 1.2 — `$script:LogPath` can diverge from the Log Path textbox
**Problem:** `$script:LogPath` is set to the hardcoded default at line 84 and only updated inside the Scan click handler (line 718). If any log write happens before a scan (e.g., after a future code change), it uses the stale hardcoded value instead of whatever the user typed.

**Fix:** Remove the `$script:LogPath = $txtLogPath.Text` assignment from inside the Scan handler. Instead, everywhere `$script:LogPath` is read (log file init, action CSV paths), replace with `$txtLogPath.Text.Trim()` directly. The script-level variable can remain as the initial default for the textbox but should not be the runtime source of truth.

**Affected lines:** 84, 718, 880, 1003

---

### 1.3 — Rename `Tag-SelectedComputers` to an approved PowerShell verb
**Problem:** `Tag-` is not an approved PowerShell verb and generates a warning when the script is analyzed by PSScriptAnalyzer or dot-sourced in a module context.

**Fix:** Rename `Tag-SelectedComputers` to `Set-ComputerDeletionTag`. Update the function definition (line 329) and all call sites (line 865).

**Affected lines:** 329, 865

---

## Phase 2 — UX Improvements

### 2.1 — Row color coding by status
**Problem:** All rows are white or alternating grey. There is no visual indicator of a computer's lifecycle stage.

**Fix:** Add a `CellFormatting` event handler to `$dataGridView`. For each row, read the `Enabled`, `MarkedForDeletion`, and `DaysSinceDisabled` values and apply a background color:

| Condition | Color |
|---|---|
| Enabled (stale, not yet acted on) | White (default) |
| Disabled + tagged, not yet past threshold | Light yellow (`#FFF3CD`) |
| Disabled + past threshold (ready to delete) | Light salmon (`#FFCCBC`) |
| Disabled + externally disabled (no tag) | Light orange (`#FFE0B2`) |

The `DeleteAfterDays` value (`$numDeleteAfterDays.Value`) must be read inside the event handler so it reflects the current setting at render time.

**New code:** Add a `$dataGridView.Add_CellFormatting({...})` block after the DataGridView is defined. Also remove `$dataGridView.AlternatingRowsDefaultCellStyle.BackColor` (line 665) since manual formatting will override it anyway.

**Affected lines:** 665, after line 666

---

### 2.2 — Add "Ready for Deletion" count to summary bar
**Problem:** The summary bar (line 430) shows `Total | Enabled | Disabled | Marked for Deletion` but does not show how many computers are actually actionable right now.

**Fix:** Add a 5th count: `Ready for Deletion` — computers where `$_.MarkedForDeletion -eq $true` AND `$_.DaysSinceDisabled -ge $numDeleteAfterDays.Value`. Update the `Update-DataGridView` function to compute this count and append it to `$lblSummary.Text`.

**Affected lines:** 424–430 (inside `Update-DataGridView`)

---

### 2.3 — Add Browse button for Log Path field
**Problem:** The Export Folder has a `...` browse button but the Log Path field does not, making it inconsistent.

**Fix:** Create a `$btnBrowseLog` button (same style as `$btnBrowseExport`) and add it to the settings layout next to `$txtLogPath`. Wire a `FolderBrowserDialog` click handler identical to `$btnBrowseExport.Add_Click`.

**Affected lines:** After line 537 (new button), line 615 (add to settings layout), after line 695 (new event handler)

---

### 2.4 — F5 keyboard shortcut to trigger Scan
**Problem:** No keyboard shortcut exists for the most common action. Users expect F5 to refresh in Windows tools.

**Fix:** Add a `KeyDown` event handler on `$form` that calls `$btnScan.PerformClick()` when `$e.KeyCode -eq [System.Windows.Forms.Keys]::F5` and the scan button is enabled. Set `$form.KeyPreview = $true` so the form receives key events before controls.

**New code:** After form creation (after line 456), set `$form.KeyPreview = $true`. Add a `$form.Add_KeyDown({...})` handler in the Event Handlers region.

---

### 2.5 — "Open Folder" prompt after CSV export
**Problem:** After exporting, the result dialog only shows a message. Users then have to manually navigate to the folder.

**Fix:** Change the export success `MessageBox` from `OK` to `YesNo` with text "Report exported successfully. Open folder?" If the user clicks Yes, call `Start-Process explorer.exe $exportFolder`.

**Affected lines:** 1067

---

## Phase 3 — Feature Additions

### 3.1 — SearchBase / OU scope filter
**Problem:** The scan always targets the entire domain (`Get-ADComputer` with no `-SearchBase`), pulling all matching objects and filtering in PowerShell. This is slow and cannot be scoped to a specific OU.

**Fix:**
1. Add a `$txtSearchBase` textbox and label to the settings panel (Row 1, after Staging OU, or a new row).
2. Add a `[string]$SearchBase` parameter to `Get-StaleComputers`.
3. In `Get-StaleComputers`, if `$SearchBase` is non-empty, add `-SearchBase $SearchBase` to the `Get-ADComputer` call.
4. Pass `$txtSearchBase.Text.Trim()` from the Scan click handler.

**Affected lines:** Settings panel layout (~line 505–612), `Get-StaleComputers` function (line 141), Scan click handler (line 734)

---

### 3.2 — Grid search/filter textbox
**Problem:** With large result sets there is no way to find a specific computer without scrolling.

**Fix:**
1. Add a `$txtGridFilter` textbox and a `$lblGridFilter` label in a new panel between the summary bar and the DataGridView.
2. On the `TextChanged` event of `$txtGridFilter`, re-call a new function `Update-DataGridView` that accepts an optional filter string. When a filter is active, only rows where `Name` or `DistinguishedName` contains the filter text (case-insensitive) are added to the DataTable.
3. Add a small "Clear" (×) button next to the filter box.

**New controls:** `$txtGridFilter`, `$lblGridFilter`, `$btnClearFilter`

**Layout change:** Add a 5th row to `$mainLayout` for the filter bar, or use a `Panel` inserted between the summary label and the DataGridView.

**Affected lines:** `$mainLayout` row styles (lines 464–468), `Update-DataGridView` function (line 376)

---

### 3.3 — Fix O(n²) lookup in `Get-SelectedComputers`
**Problem:** For every checked row, `Get-SelectedComputers` does `Where-Object { $_.Name -eq $computerName }` over all of `$script:ComputerReport`. In a 1,000-computer environment this is ~500,000 comparisons per call.

**Fix:** Build a `[hashtable]` index from `$script:ComputerReport` keyed by `Name` at the top of `Get-SelectedComputers` (or maintain it as `$script:ComputerIndex` and rebuild it at the end of each scan). Replace the `Where-Object` lookup with a direct `$index[$computerName]` hash lookup.

**Affected lines:** `Get-SelectedComputers` function (lines 434–448), end of Scan click handler (line 744) to build the index.

---

### 3.4 — BackgroundWorker for scan and operations
**Problem:** `Get-ADComputer` and bulk `Remove-ADObject` run on the UI thread. For large environments, the window can freeze for 30+ seconds. `DoEvents()` partially mitigates this but can cause reentrancy bugs.

**Fix:**
1. Add a `System.Windows.Forms.ProgressBar` control (`$progressBar`) to the layout — placed between the summary label and the grid filter (or in a status strip).
2. Create a `System.ComponentModel.BackgroundWorker` (`$bgWorker`) with `WorkerReportsProgress = $true` and `WorkerSupportsCancellation = $true`.
3. Move the `Get-StaleComputers` call and report generation into `$bgWorker.Add_DoWork({...})`.
4. Use `$bgWorker.ReportProgress()` to update `$progressBar` and log messages via `$form.Invoke()` (required for cross-thread UI updates).
5. On `RunWorkerCompleted`, update the grid and re-enable buttons.
6. Add a "Cancel" button that calls `$bgWorker.CancelAsync()`.

**Note:** This is the most complex item. PowerShell runspaces or `[System.Threading.Tasks.Task]` are alternatives if BackgroundWorker proves awkward with PS variable scoping.

**Affected lines:** Scan click handler (lines 709–766), layout (add progress bar row)

---

### 3.5 — Right-click context menu on grid rows
**Problem:** Acting on a single computer requires: check the box → scroll up to button row → click button. A context menu is faster for one-off operations.

**Fix:**
1. Create a `$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip`.
2. Add items: `Disable/Tag`, `Delete`, `Re-enable`, `Copy Name`, `Copy DN`.
3. On `$dataGridView.Add_CellMouseDown({...})`, if right-click, select the clicked row and show the context menu.
4. Each menu item's `Click` handler calls the same underlying function as the corresponding button, but pre-populates the selection to just that row.

**New code:** After DataGridView definition. Depends on Phase 3.4 (Re-enable function) for the Re-enable menu item, but the rest can be implemented independently.

---

### 3.6 — Re-enable selected computers
**Problem:** If a computer was incorrectly disabled by the script, there is no way to undo it from the GUI.

**Fix:**
1. Add a function `Restore-SelectedComputers` that:
   - Calls `Set-ADComputer -Identity $dn -Enabled $true`
   - If the description starts with `DISABLED: `, strips that prefix and restores the original description (which is stored after `| Original: ` in the description field).
   - Logs the action with operator identity.
2. Add a `$btnReEnable` button (grey/neutral color) to the button panel, enabled after a scan.
3. Wire confirmation dialog: "You are about to re-enable N computer(s). Continue?"
4. After completion, trigger a re-scan.

**Regex to strip description:** `^DISABLED:\s*\d{4}-\d{2}-\d{2}\s*\|\s*Original:\s*(.+?)\s*\|.*$` → capture group 1 is the original description.

**Affected lines:** New function after `Tag-SelectedComputers` (line 374), button panel (line 632), new event handler in Event Handlers region.

---

## Implementation Order & Dependencies

```
Phase 1.1 → Phase 1.2 → Phase 1.3        (independent, do in order)
Phase 2.1 → Phase 2.2                     (color coding before summary count)
Phase 2.3, 2.4, 2.5                       (independent of each other)
Phase 3.1                                 (independent)
Phase 3.2                                 (depends on Update-DataGridView being stable — do after Phase 2)
Phase 3.3                                 (independent, quick win)
Phase 3.6                                 (implement before 3.5 so Re-enable menu item is ready)
Phase 3.5                                 (depends on 3.6)
Phase 3.4                                 (do last — most invasive, touches scan handler and layout)
```

---

## Risk Notes

- **BackgroundWorker (3.4):** PowerShell's variable scoping makes cross-thread updates awkward. Test thoroughly in a non-production AD environment. If runspace-based threading causes instability, a simpler alternative is to just add a marquee-style `ProgressBar` with `Style = Marquee` during blocking operations — no threading required, just visual feedback.
- **Row coloring (2.1):** The `CellFormatting` event fires for every cell render. Keep the logic lightweight (no AD calls, no regex) to avoid sluggish scrolling.
- **SearchBase (3.1):** If an invalid DN is entered, `Get-ADComputer` will throw. Add a `try/catch` around the DN validation or pre-validate with `Get-ADOrganizationalUnit` before scanning.
- **Description stripping (3.6):** The regex relies on the description format written by this script. If a computer's original description itself contained `| Original:` or `| Stale`, the regex could mis-parse it. Consider using a more robust delimiter when writing descriptions in `Disable-SelectedComputers` and `Set-ComputerDeletionTag`.
