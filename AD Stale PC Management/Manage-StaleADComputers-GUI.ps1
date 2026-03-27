<#
.SYNOPSIS
    GUI version - Identifies, disables, and deletes stale computer objects in Active Directory.

.DESCRIPTION
    Provides a graphical interface for managing stale AD computer objects.
    Implements a 3-stage lifecycle for stale AD computer management:
    Stage 1: Identify workstations inactive for X days (default 365)
    Stage 2: Disable workstations and move to staging OU, tag with disable date
    Stage 3: Delete workstations that have been disabled for Y days (default 30)

    Features:
    - Automatic RSAT/ActiveDirectory module installation if not detected
    - Full audit trail: operator identity (DOMAIN\User) in every log line and CSV record
    - PerformedBy field in all action CSV exports
    - Operator identity displayed in window title bar
    - Row color coding by lifecycle stage
    - SearchBase / OU scope filter for targeted scans
    - Grid search/filter for large result sets
    - Re-enable incorrectly disabled computers with description restoration
    - Right-click context menu for quick single-computer actions
    - Marquee progress bar during long operations
    - F5 keyboard shortcut to trigger scan

.NOTES
    Author: Alberto de la Torre
    Version: 3.0
    Date: March 2026

    Changelog:
    3.0 - Phase 1-3 improvements:
          Bugs: action CSV paths fixed (now go to export folder), LogPath reads from
                textbox at runtime instead of stale variable, Tag- verb renamed to
                Set-ComputerDeletionTag
          UX: row color coding by lifecycle stage, Ready for Deletion count in summary
              bar, Log Path browse button, F5 scan shortcut, Open Folder prompt after
              export, tooltips on all action buttons
          Features: SearchBase/OU scope filter, grid filter textbox, O(1) hashtable
                    lookup replaces O(n^2) search, marquee progress bar, right-click
                    context menu, Re-enable computers with description restoration
    2.2 - Added Select Eligible button for one-click deletion-ready selection
    2.1 - Added user account audit logging, auto-install AD module, fixed single-object delete bug
    1.0 - Initial release
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Active Directory Module Check and Auto-Install
if (-not (Get-Module -ListAvailable ActiveDirectory)) {
    try {
        $isServer  = (Get-CimInstance Win32_OperatingSystem).Caption -match "Server"
        $statusMsg = "Active Directory module not detected. Attempting to install required RSAT tools... This may take a few minutes."
        [System.Windows.Forms.MessageBox]::Show($statusMsg, "Module Installation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        if ($isServer) {
            if (Get-Command Install-WindowsFeature -ErrorAction SilentlyContinue) {
                Install-WindowsFeature RSAT-AD-PowerShell -ErrorAction Stop
            } else {
                throw "Install-WindowsFeature command not found. Cannot install RSAT on this Server version automatically."
            }
        } else {
            if (Get-Command Add-WindowsCapability -ErrorAction SilentlyContinue) {
                Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop
            } else {
                throw "Add-WindowsCapability command not found. Cannot install RSAT on this Windows version automatically."
            }
        }

        if (-not (Get-Module -ListAvailable ActiveDirectory)) {
            throw "Installation completed but module still not detected. Please install manually."
        }
        [System.Windows.Forms.MessageBox]::Show("Active Directory module installed successfully. Please restart the script if it does not load correctly.", "Installation Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to automatically install Active Directory module: $($_.Exception.Message)`n`nPlease run PowerShell as Administrator or install RSAT (Active Directory Domain Services & Lightweight Directory Services Tools) manually.", "Installation Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
}

Import-Module ActiveDirectory -ErrorAction SilentlyContinue
if (-not (Get-Module ActiveDirectory)) {
    [System.Windows.Forms.MessageBox]::Show("Active Directory module could not be loaded. Please ensure RSAT is installed and that you have appropriate permissions.", "Module Load Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}
#endregion

#region Script Variables
$script:LogFile        = $null
$script:StaleComputers = @()
$script:ComputerReport = @()
$script:ComputerIndex  = @{}   # hashtable for O(1) name->report lookup
$script:RunningUser    = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
#endregion

#region Functions

function Write-GUILog {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )

    $timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] [$($script:RunningUser)] $Message"

    $txtLog.AppendText("$logMessage`r`n")
    $txtLog.ScrollToCaret()

    # Write to log file - read path from textbox at runtime, not a stale variable
    if ($script:LogFile -and (Test-Path (Split-Path $script:LogFile -Parent))) {
        Add-Content -Path $script:LogFile -Value $logMessage
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function Get-StaleComputers {
    param(
        [int]$DaysInactive,
        [bool]$IncludeWorkstations,
        [bool]$IncludeServers,
        [string]$SearchBase = ""
    )

    $cutoffDate = (Get-Date).AddDays(-$DaysInactive)

    $osFilters = @()
    if ($IncludeWorkstations) {
        $osFilters += "(OperatingSystem -like 'Windows 10*')"
        $osFilters += "(OperatingSystem -like 'Windows 11*')"
    }
    if ($IncludeServers) {
        $osFilters += "(OperatingSystem -like '*Server*')"
    }

    if ($osFilters.Count -eq 0) {
        Write-GUILog "No OS types selected for search." -Level Warning
        return @()
    }

    $filterString = $osFilters -join " -or "
    $scopeMsg     = if ($SearchBase) { "within OU: $SearchBase" } else { "entire domain" }
    Write-GUILog "Searching for computers inactive since $($cutoffDate.ToString('yyyy-MM-dd')) ($scopeMsg)"

    $adParams = @{
        Filter     = $filterString
        Properties = @(
            'Name', 'DistinguishedName', 'LastLogonTimestamp', 'LastLogonDate',
            'pwdLastSet', 'whenChanged', 'whenCreated', 'Enabled',
            'Description', 'OperatingSystem', 'OperatingSystemVersion'
        )
    }
    if ($SearchBase) { $adParams.SearchBase = $SearchBase }

    $allComputers = Get-ADComputer @adParams
    Write-GUILog "Found $($allComputers.Count) computers matching OS filter in AD"

    $staleComputers = $allComputers | Where-Object {
        $lastLogon = if ($_.LastLogonTimestamp -and $_.LastLogonTimestamp -ne 0) {
            [DateTime]::FromFileTime($_.LastLogonTimestamp)
        } else {
            $_.whenCreated
        }
        return $lastLogon -lt $cutoffDate
    }

    Write-GUILog "Found $($staleComputers.Count) stale computers"
    return $staleComputers
}

function Get-ComputerReportData {
    param(
        [Parameter(ValueFromPipeline)]
        $Computer
    )

    process {
        $lastLogon = if ($Computer.LastLogonTimestamp -and $Computer.LastLogonTimestamp -ne 0) {
            [DateTime]::FromFileTime($Computer.LastLogonTimestamp)
        } else {
            $null
        }

        $pwdLastSet = if ($Computer.pwdLastSet -and $Computer.pwdLastSet -ne 0) {
            [DateTime]::FromFileTime($Computer.pwdLastSet)
        } else {
            $null
        }

        $daysSinceLogon = if ($lastLogon) {
            [math]::Round(((Get-Date) - $lastLogon).TotalDays, 0)
        } else {
            -1
        }

        $markedForDeletion = $false
        $disableDate       = $null
        if ($Computer.Description -match "DISABLED:\s*(\d{4}-\d{2}-\d{2})") {
            $markedForDeletion = $true
            $disableDate       = [DateTime]::Parse($Matches[1])
        }

        [PSCustomObject]@{
            Select             = $false
            Name               = $Computer.Name
            Enabled            = $Computer.Enabled
            DistinguishedName  = $Computer.DistinguishedName
            LastLogonDate      = if ($lastLogon) { $lastLogon.ToString("yyyy-MM-dd") } else { "Never" }
            DaysSinceLogon     = if ($daysSinceLogon -eq -1) { "Never" } else { $daysSinceLogon }
            DaysSinceLogonSort = $daysSinceLogon
            PasswordLastSet    = if ($pwdLastSet) { $pwdLastSet.ToString("yyyy-MM-dd") } else { "" }
            WhenCreated        = $Computer.whenCreated.ToString("yyyy-MM-dd")
            OperatingSystem    = $Computer.OperatingSystem
            Description        = $Computer.Description
            MarkedForDeletion  = $markedForDeletion
            DisableDate        = if ($disableDate) { $disableDate.ToString("yyyy-MM-dd") } else { "" }
            DaysSinceDisabled  = if ($disableDate) { [math]::Round(((Get-Date) - $disableDate).TotalDays, 0) } else { $null }
        }
    }
}

function Disable-SelectedComputers {
    param(
        [array]$Computers,
        [string]$TargetOU,
        [int]$InactiveDays
    )

    $results = @()
    $today   = Get-Date -Format "yyyy-MM-dd"

    foreach ($computer in $Computers) {
        $computerName = $computer.Name
        try {
            $adComputer          = Get-ADComputer -Identity $computerName -Properties Description
            $originalDescription = $adComputer.Description
            $newDescription      = "DISABLED: $today | Original: $originalDescription | Stale $InactiveDays+ days"

            Set-ADComputer -Identity $adComputer.DistinguishedName -Enabled $false -Description $newDescription
            Write-GUILog "Disabled computer: $computerName"

            if ($TargetOU) {
                Move-ADObject -Identity $adComputer.DistinguishedName -TargetPath $TargetOU
                Write-GUILog "Moved $computerName to staging OU"
            }

            $results += [PSCustomObject]@{
                ComputerName = $computerName
                Action       = "Disabled"
                Status       = "Success"
                PerformedBy  = $script:RunningUser
                Timestamp    = Get-Date
                NewLocation  = if ($TargetOU) { $TargetOU } else { "Not moved" }
            }
        }
        catch {
            Write-GUILog "Failed to disable $computerName : $_" -Level Error
            $results += [PSCustomObject]@{
                ComputerName = $computerName
                Action       = "Disable"
                Status       = "Failed"
                PerformedBy  = $script:RunningUser
                Timestamp    = Get-Date
                Error        = $_.Exception.Message
            }
        }
    }
    return $results
}

function Remove-SelectedComputers {
    param(
        [array]$Computers,
        [int]$MinDaysDisabled
    )

    $results = @()

    foreach ($computer in $Computers) {
        $computerName = $computer.Name

        if ($computer.MarkedForDeletion -and $computer.DaysSinceDisabled -and $computer.DaysSinceDisabled -lt $MinDaysDisabled) {
            Write-GUILog "Skipping $computerName - only disabled for $($computer.DaysSinceDisabled) days (min: $MinDaysDisabled)" -Level Warning
            continue
        }

        try {
            $adComputer = Get-ADComputer -Identity $computerName
            Remove-ADObject -Identity $adComputer -Recursive -Confirm:$false
            Write-GUILog "Deleted computer: $computerName"

            $results += [PSCustomObject]@{
                ComputerName = $computerName
                Action       = "Deleted"
                Status       = "Success"
                PerformedBy  = $script:RunningUser
                Timestamp    = Get-Date
                DaysDisabled = $computer.DaysSinceDisabled
            }
        }
        catch {
            Write-GUILog "Failed to delete $computerName : $_" -Level Error
            $results += [PSCustomObject]@{
                ComputerName = $computerName
                Action       = "Delete"
                Status       = "Failed"
                PerformedBy  = $script:RunningUser
                Timestamp    = Get-Date
                Error        = $_.Exception.Message
            }
        }
    }
    return $results
}

function Set-ComputerDeletionTag {
    # Renamed from Tag-SelectedComputers — 'Tag' is not an approved PowerShell verb
    param(
        [array]$Computers,
        [int]$InactiveDays
    )

    $results = @()
    $today   = Get-Date -Format "yyyy-MM-dd"

    foreach ($computer in $Computers) {
        $computerName = $computer.Name
        try {
            $adComputer          = Get-ADComputer -Identity $computerName -Properties Description
            $originalDescription = $adComputer.Description
            $newDescription      = "DISABLED: $today | Original: $originalDescription | Tagged for deletion (stale $InactiveDays+ days)"

            Set-ADComputer -Identity $adComputer.DistinguishedName -Description $newDescription
            Write-GUILog "Tagged computer for deletion: $computerName"

            $results += [PSCustomObject]@{
                ComputerName   = $computerName
                Action         = "Tagged"
                Status         = "Success"
                PerformedBy    = $script:RunningUser
                Timestamp      = Get-Date
                NewDescription = $newDescription
            }
        }
        catch {
            Write-GUILog "Failed to tag $computerName : $_" -Level Error
            $results += [PSCustomObject]@{
                ComputerName = $computerName
                Action       = "Tag"
                Status       = "Failed"
                PerformedBy  = $script:RunningUser
                Timestamp    = Get-Date
                Error        = $_.Exception.Message
            }
        }
    }
    return $results
}

function Restore-SelectedComputers {
    param([array]$Computers)

    $results = @()

    foreach ($computer in $Computers) {
        $computerName = $computer.Name
        try {
            $adComputer  = Get-ADComputer -Identity $computerName -Properties Description
            $currentDesc = $adComputer.Description

            # Restore original description from the DISABLED tag format:
            # "DISABLED: YYYY-MM-DD | Original: <desc> | Stale X+ days"
            $restoredDesc = if ($currentDesc -match "^DISABLED:\s*\d{4}-\d{2}-\d{2}\s*\|\s*Original:\s*(.+?)\s*\|") {
                $Matches[1]
            } else {
                $currentDesc
            }

            Set-ADComputer -Identity $adComputer.DistinguishedName -Enabled $true -Description $restoredDesc
            Write-GUILog "Re-enabled computer: $computerName"

            $results += [PSCustomObject]@{
                ComputerName        = $computerName
                Action              = "Re-enabled"
                Status              = "Success"
                PerformedBy         = $script:RunningUser
                Timestamp           = Get-Date
                RestoredDescription = $restoredDesc
            }
        }
        catch {
            Write-GUILog "Failed to re-enable $computerName : $_" -Level Error
            $results += [PSCustomObject]@{
                ComputerName = $computerName
                Action       = "Re-enable"
                Status       = "Failed"
                PerformedBy  = $script:RunningUser
                Timestamp    = Get-Date
                Error        = $_.Exception.Message
            }
        }
    }
    return $results
}

function Update-DataGridView {
    param([string]$Filter = "")

    $dataGridView.DataSource = $null

    # Apply name/OU filter to display subset; summary counts always reflect full report
    $reportData = $script:ComputerReport
    if ($Filter) {
        $reportData = @($reportData | Where-Object {
            $_.Name -like "*$Filter*" -or $_.DistinguishedName -like "*$Filter*"
        })
    }

    if ($reportData.Count -gt 0) {
        $dataTable = New-Object System.Data.DataTable

        $dataTable.Columns.Add("Select",            [bool])   | Out-Null
        $dataTable.Columns.Add("Name",              [string]) | Out-Null
        $dataTable.Columns.Add("Enabled",           [bool])   | Out-Null
        $dataTable.Columns.Add("LastLogonDate",     [string]) | Out-Null
        $dataTable.Columns.Add("DaysSinceLogon",    [string]) | Out-Null
        $dataTable.Columns.Add("OperatingSystem",   [string]) | Out-Null
        $dataTable.Columns.Add("MarkedForDeletion", [bool])   | Out-Null
        $dataTable.Columns.Add("DisableDate",       [string]) | Out-Null
        $dataTable.Columns.Add("DaysSinceDisabled", [string]) | Out-Null
        $dataTable.Columns.Add("DistinguishedName", [string]) | Out-Null

        foreach ($computer in $reportData) {
            $row = $dataTable.NewRow()
            $row["Select"]            = $false
            $row["Name"]              = $computer.Name
            $row["Enabled"]           = $computer.Enabled
            $row["LastLogonDate"]     = $computer.LastLogonDate
            $row["DaysSinceLogon"]    = $computer.DaysSinceLogon
            $row["OperatingSystem"]   = $computer.OperatingSystem
            $row["MarkedForDeletion"] = $computer.MarkedForDeletion
            $row["DisableDate"]       = $computer.DisableDate
            $row["DaysSinceDisabled"] = if ($computer.DaysSinceDisabled) { $computer.DaysSinceDisabled.ToString() } else { "" }
            $row["DistinguishedName"] = $computer.DistinguishedName
            $dataTable.Rows.Add($row)
        }

        $dataGridView.DataSource = $dataTable

        $dataGridView.Columns["Select"].Width            = 50
        $dataGridView.Columns["Name"].Width              = 120
        $dataGridView.Columns["Enabled"].Width           = 60
        $dataGridView.Columns["LastLogonDate"].Width     = 100
        $dataGridView.Columns["DaysSinceLogon"].Width    = 80
        $dataGridView.Columns["OperatingSystem"].Width   = 150
        $dataGridView.Columns["MarkedForDeletion"].Width = 100
        $dataGridView.Columns["DisableDate"].Width       = 90
        $dataGridView.Columns["DaysSinceDisabled"].Width = 100
        $dataGridView.Columns["DistinguishedName"].Width = 300
    }

    # Summary always reflects the full unfiltered report
    $totalCount    = $script:ComputerReport.Count
    $enabledCount  = ($script:ComputerReport | Where-Object { $_.Enabled }).Count
    $disabledCount = ($script:ComputerReport | Where-Object { -not $_.Enabled }).Count
    $markedCount   = ($script:ComputerReport | Where-Object { $_.MarkedForDeletion }).Count
    $readyCount    = ($script:ComputerReport | Where-Object {
        $_.MarkedForDeletion -and ($null -ne $_.DaysSinceDisabled) -and $_.DaysSinceDisabled -ge $numDeleteAfterDays.Value
    }).Count

    $filterNote      = if ($Filter) { "  [filter: '$Filter' — $($reportData.Count) shown]" } else { "" }
    $lblSummary.Text = "Total: $totalCount  |  Enabled: $enabledCount  |  Disabled: $disabledCount  |  Marked: $markedCount  |  Ready for Deletion: $readyCount$filterNote"
}

function Get-SelectedComputers {
    # O(1) hashtable lookup — replaces the previous O(n^2) Where-Object per row
    $selected = @()
    foreach ($row in $dataGridView.Rows) {
        if ($row.Cells["Select"].Value -eq $true) {
            $name = $row.Cells["Name"].Value
            if ($script:ComputerIndex.ContainsKey($name)) {
                $selected += $script:ComputerIndex[$name]
            }
        }
    }
    return @($selected)
}

#endregion Functions

#region Build Form

$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Stale AD Computer Management  -  $($script:RunningUser)"
$form.Size            = New-Object System.Drawing.Size(1200, 860)
$form.StartPosition   = "CenterScreen"
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)
$form.KeyPreview      = $true   # required for F5 handler

# Main layout: Settings | Summary | Filter+Progress | DataGrid | Log
$mainLayout             = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock        = [System.Windows.Forms.DockStyle]::Fill
$mainLayout.ColumnCount = 1
$mainLayout.RowCount    = 5
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 220))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 25)))  | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))  | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))  | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 150))) | Out-Null
$mainLayout.Padding = New-Object System.Windows.Forms.Padding(10)

#region Settings Panel
$settingsGroup      = New-Object System.Windows.Forms.GroupBox
$settingsGroup.Text = "Settings"
$settingsGroup.Dock = [System.Windows.Forms.DockStyle]::Fill

$settingsLayout             = New-Object System.Windows.Forms.TableLayoutPanel
$settingsLayout.Dock        = [System.Windows.Forms.DockStyle]::Fill
$settingsLayout.ColumnCount = 6
$settingsLayout.RowCount    = 5
$settingsLayout.Padding     = New-Object System.Windows.Forms.Padding(5)

# Row 0: Inactive Days | Delete After Days | Staging OU
$lblInactiveDays          = New-Object System.Windows.Forms.Label
$lblInactiveDays.Text     = "Inactive Days:"
$lblInactiveDays.AutoSize = $true
$lblInactiveDays.Anchor   = [System.Windows.Forms.AnchorStyles]::Left

$numInactiveDays          = New-Object System.Windows.Forms.NumericUpDown
$numInactiveDays.Minimum  = 30
$numInactiveDays.Maximum  = 730
$numInactiveDays.Value    = 365
$numInactiveDays.Width    = 80

$lblDeleteAfterDays          = New-Object System.Windows.Forms.Label
$lblDeleteAfterDays.Text     = "Delete After Days:"
$lblDeleteAfterDays.AutoSize = $true
$lblDeleteAfterDays.Anchor   = [System.Windows.Forms.AnchorStyles]::Left

$numDeleteAfterDays          = New-Object System.Windows.Forms.NumericUpDown
$numDeleteAfterDays.Minimum  = 7
$numDeleteAfterDays.Maximum  = 365
$numDeleteAfterDays.Value    = 30
$numDeleteAfterDays.Width    = 80

$lblStagingOU          = New-Object System.Windows.Forms.Label
$lblStagingOU.Text     = "Staging OU (DN):"
$lblStagingOU.AutoSize = $true
$lblStagingOU.Anchor   = [System.Windows.Forms.AnchorStyles]::Right

$txtStagingOU        = New-Object System.Windows.Forms.TextBox
$txtStagingOU.Width  = 350
$txtStagingOU.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

# Row 1: Log Path + Browse | OS Types
$lblLogPath          = New-Object System.Windows.Forms.Label
$lblLogPath.Text     = "Log Path:"
$lblLogPath.AutoSize = $true
$lblLogPath.Anchor   = [System.Windows.Forms.AnchorStyles]::Left

$txtLogPath       = New-Object System.Windows.Forms.TextBox
$txtLogPath.Text  = "C:\Temp\StaleADComputerLogs"
$txtLogPath.Width = 230

$btnBrowseLog       = New-Object System.Windows.Forms.Button
$btnBrowseLog.Text  = "..."
$btnBrowseLog.Width = 30

$lblOSTypes          = New-Object System.Windows.Forms.Label
$lblOSTypes.Text     = "Include OS Types:"
$lblOSTypes.AutoSize = $true
$lblOSTypes.Anchor   = [System.Windows.Forms.AnchorStyles]::Left

$chkWorkstations          = New-Object System.Windows.Forms.CheckBox
$chkWorkstations.Text     = "Workstations (Win 10/11)"
$chkWorkstations.Checked  = $true
$chkWorkstations.AutoSize = $true

$chkServers          = New-Object System.Windows.Forms.CheckBox
$chkServers.Text     = "Servers"
$chkServers.Checked  = $false
$chkServers.AutoSize = $true

$osPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$osPanel.AutoSize = $true
$osPanel.Controls.Add($chkWorkstations)
$osPanel.Controls.Add($chkServers)

# Row 2: Search Base (OU scope filter)
$lblSearchBase          = New-Object System.Windows.Forms.Label
$lblSearchBase.Text     = "Search Base (OU):"
$lblSearchBase.AutoSize = $true
$lblSearchBase.Anchor   = [System.Windows.Forms.AnchorStyles]::Left

$txtSearchBase                 = New-Object System.Windows.Forms.TextBox
$txtSearchBase.Anchor          = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$txtSearchBase.PlaceholderText = "Optional: OU=Workstations,DC=contoso,DC=com  (leave blank to search entire domain)"

# Row 3: Export Folder + Browse
$lblExportPath          = New-Object System.Windows.Forms.Label
$lblExportPath.Text     = "Export Folder:"
$lblExportPath.AutoSize = $true
$lblExportPath.Anchor   = [System.Windows.Forms.AnchorStyles]::Left

$txtExportPath       = New-Object System.Windows.Forms.TextBox
$txtExportPath.Width = 300
$txtExportPath.Text  = "C:\Temp\StaleADComputerReports"

$btnBrowseExport       = New-Object System.Windows.Forms.Button
$btnBrowseExport.Text  = "..."
$btnBrowseExport.Width = 30

# Row 4: Action Buttons
$btnScan               = New-Object System.Windows.Forms.Button
$btnScan.Text          = "Scan for Stale Computers"
$btnScan.Width         = 180
$btnScan.Height        = 30
$btnScan.BackColor     = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnScan.ForeColor     = [System.Drawing.Color]::White
$btnScan.FlatStyle     = [System.Windows.Forms.FlatStyle]::Flat

$btnSelectAll          = New-Object System.Windows.Forms.Button
$btnSelectAll.Text     = "Select All"
$btnSelectAll.Width    = 80
$btnSelectAll.Height   = 30

$btnSelectNone         = New-Object System.Windows.Forms.Button
$btnSelectNone.Text    = "Select None"
$btnSelectNone.Width   = 80
$btnSelectNone.Height  = 30

$btnSelectEligible           = New-Object System.Windows.Forms.Button
$btnSelectEligible.Text      = "Select Eligible"
$btnSelectEligible.Width     = 110
$btnSelectEligible.Height    = 30
$btnSelectEligible.BackColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
$btnSelectEligible.ForeColor = [System.Drawing.Color]::White
$btnSelectEligible.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSelectEligible.Enabled   = $false

$btnDisable            = New-Object System.Windows.Forms.Button
$btnDisable.Text       = "Disable/Tag"
$btnDisable.Width      = 100
$btnDisable.Height     = 30
$btnDisable.BackColor  = [System.Drawing.Color]::FromArgb(255, 185, 0)
$btnDisable.FlatStyle  = [System.Windows.Forms.FlatStyle]::Flat
$btnDisable.Enabled    = $false

$btnDelete             = New-Object System.Windows.Forms.Button
$btnDelete.Text        = "Delete Selected"
$btnDelete.Width       = 120
$btnDelete.Height      = 30
$btnDelete.BackColor   = [System.Drawing.Color]::FromArgb(232, 17, 35)
$btnDelete.ForeColor   = [System.Drawing.Color]::White
$btnDelete.FlatStyle   = [System.Windows.Forms.FlatStyle]::Flat
$btnDelete.Enabled     = $false

$btnReEnable           = New-Object System.Windows.Forms.Button
$btnReEnable.Text      = "Re-enable"
$btnReEnable.Width     = 90
$btnReEnable.Height    = 30
$btnReEnable.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$btnReEnable.ForeColor = [System.Drawing.Color]::White
$btnReEnable.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnReEnable.Enabled   = $false

$btnExport             = New-Object System.Windows.Forms.Button
$btnExport.Text        = "Export to CSV"
$btnExport.Width       = 120
$btnExport.Height      = 30
$btnExport.BackColor   = [System.Drawing.Color]::FromArgb(16, 124, 16)
$btnExport.ForeColor   = [System.Drawing.Color]::White
$btnExport.FlatStyle   = [System.Windows.Forms.FlatStyle]::Flat
$btnExport.Enabled     = $false

# Tooltips
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($btnScan,           "Scan Active Directory for stale computers (F5)")
$toolTip.SetToolTip($btnSelectEligible, "Select all computers currently eligible for deletion")
$toolTip.SetToolTip($btnDisable,        "Disable enabled stale computers and tag externally-disabled ones")
$toolTip.SetToolTip($btnDelete,         "Permanently delete selected disabled computers from AD")
$toolTip.SetToolTip($btnReEnable,       "Re-enable selected computers and restore their original description")
$toolTip.SetToolTip($btnExport,         "Export current scan results to a CSV file")
$toolTip.SetToolTip($txtSearchBase,     "Limit scan to a specific OU. Leave blank to search the entire domain.")

# Wire controls into settings layout
$settingsLayout.Controls.Add($lblInactiveDays,    0, 0)
$settingsLayout.Controls.Add($numInactiveDays,    1, 0)
$settingsLayout.Controls.Add($lblDeleteAfterDays, 2, 0)
$settingsLayout.Controls.Add($numDeleteAfterDays, 3, 0)
$settingsLayout.Controls.Add($lblStagingOU,       4, 0)
$settingsLayout.Controls.Add($txtStagingOU,       5, 0)

$settingsLayout.Controls.Add($lblLogPath,   0, 1)
$settingsLayout.Controls.Add($txtLogPath,   1, 1)
$settingsLayout.Controls.Add($btnBrowseLog, 2, 1)
$settingsLayout.Controls.Add($lblOSTypes,   3, 1)
$settingsLayout.Controls.Add($osPanel,      4, 1)
$settingsLayout.SetColumnSpan($osPanel, 2)

$settingsLayout.Controls.Add($lblSearchBase, 0, 2)
$settingsLayout.Controls.Add($txtSearchBase, 1, 2)
$settingsLayout.SetColumnSpan($txtSearchBase, 5)

$settingsLayout.Controls.Add($lblExportPath,   0, 3)
$settingsLayout.Controls.Add($txtExportPath,   1, 3)
$settingsLayout.SetColumnSpan($txtExportPath, 4)
$settingsLayout.Controls.Add($btnBrowseExport, 5, 3)

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$buttonPanel.Controls.Add($btnScan)
$buttonPanel.Controls.Add($btnSelectAll)
$buttonPanel.Controls.Add($btnSelectNone)
$buttonPanel.Controls.Add($btnSelectEligible)
$buttonPanel.Controls.Add($btnDisable)
$buttonPanel.Controls.Add($btnDelete)
$buttonPanel.Controls.Add($btnReEnable)
$buttonPanel.Controls.Add($btnExport)
$settingsLayout.Controls.Add($buttonPanel, 0, 4)
$settingsLayout.SetColumnSpan($buttonPanel, 6)

$settingsGroup.Controls.Add($settingsLayout)
#endregion Settings Panel

#region Summary Label
$lblSummary           = New-Object System.Windows.Forms.Label
$lblSummary.Text      = "Total: 0  |  Enabled: 0  |  Disabled: 0  |  Marked: 0  |  Ready for Deletion: 0"
$lblSummary.Dock      = [System.Windows.Forms.DockStyle]::Fill
$lblSummary.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblSummary.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
#endregion

#region Filter / Progress Bar Panel
$filterPanel      = New-Object System.Windows.Forms.Panel
$filterPanel.Dock = [System.Windows.Forms.DockStyle]::Fill

$lblFilter          = New-Object System.Windows.Forms.Label
$lblFilter.Text     = "Filter:"
$lblFilter.AutoSize = $true
$lblFilter.Location = New-Object System.Drawing.Point(0, 8)

$txtGridFilter                 = New-Object System.Windows.Forms.TextBox
$txtGridFilter.Location        = New-Object System.Drawing.Point(45, 5)
$txtGridFilter.Width           = 250
$txtGridFilter.PlaceholderText = "Filter by computer name or OU..."

$btnClearFilter           = New-Object System.Windows.Forms.Button
$btnClearFilter.Text      = "x"
$btnClearFilter.Width     = 25
$btnClearFilter.Height    = 23
$btnClearFilter.Location  = New-Object System.Drawing.Point(300, 5)
$btnClearFilter.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

$progressBar          = New-Object System.Windows.Forms.ProgressBar
$progressBar.Style    = [System.Windows.Forms.ProgressBarStyle]::Marquee
$progressBar.Location = New-Object System.Drawing.Point(340, 6)
$progressBar.Size     = New-Object System.Drawing.Size(220, 21)
$progressBar.Visible  = $false

$filterPanel.Controls.AddRange(@($lblFilter, $txtGridFilter, $btnClearFilter, $progressBar))
#endregion

#region DataGridView
$dataGridView                       = New-Object System.Windows.Forms.DataGridView
$dataGridView.Dock                  = [System.Windows.Forms.DockStyle]::Fill
$dataGridView.AllowUserToAddRows    = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.SelectionMode         = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.MultiSelect           = $true
$dataGridView.ReadOnly              = $false
$dataGridView.AutoSizeColumnsMode   = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$dataGridView.RowHeadersVisible     = $false
# AlternatingRowsDefaultCellStyle removed — CellFormatting handles all row coloring
#endregion DataGridView

#region Right-Click Context Menu
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$ctxDisable  = New-Object System.Windows.Forms.ToolStripMenuItem("Disable / Tag")
$ctxDelete   = New-Object System.Windows.Forms.ToolStripMenuItem("Delete")
$ctxReEnable = New-Object System.Windows.Forms.ToolStripMenuItem("Re-enable")
$ctxSep      = New-Object System.Windows.Forms.ToolStripSeparator
$ctxCopyName = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Name")
$ctxCopyDN   = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Distinguished Name")
$contextMenu.Items.AddRange(@($ctxDisable, $ctxDelete, $ctxReEnable, $ctxSep, $ctxCopyName, $ctxCopyDN))
$dataGridView.ContextMenuStrip = $contextMenu
#endregion

#region Log TextBox
$logGroup      = New-Object System.Windows.Forms.GroupBox
$logGroup.Text = "Activity Log"
$logGroup.Dock = [System.Windows.Forms.DockStyle]::Fill

$txtLog            = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline  = $true
$txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtLog.Dock       = [System.Windows.Forms.DockStyle]::Fill
$txtLog.ReadOnly   = $true
$txtLog.Font       = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BackColor  = [System.Drawing.Color]::FromArgb(30, 30, 30)
$txtLog.ForeColor  = [System.Drawing.Color]::FromArgb(200, 200, 200)

$logGroup.Controls.Add($txtLog)
#endregion

# Assemble main layout
$mainLayout.Controls.Add($settingsGroup, 0, 0)
$mainLayout.Controls.Add($lblSummary,    0, 1)
$mainLayout.Controls.Add($filterPanel,   0, 2)
$mainLayout.Controls.Add($dataGridView,  0, 3)
$mainLayout.Controls.Add($logGroup,      0, 4)
$form.Controls.Add($mainLayout)

#endregion Build Form

#region Event Handlers

# F5 - trigger scan
$form.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::F5 -and $btnScan.Enabled) {
        $btnScan.PerformClick()
    }
})

# Browse Log Path
$btnBrowseLog.Add_Click({
    $dlg                    = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description        = "Select folder for log files"
    $dlg.ShowNewFolderButton = $true
    if ($txtLogPath.Text.Trim() -and (Test-Path $txtLogPath.Text.Trim())) {
        $dlg.SelectedPath = $txtLogPath.Text.Trim()
    }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtLogPath.Text = $dlg.SelectedPath
    }
})

# Browse Export Folder
$btnBrowseExport.Add_Click({
    $dlg                    = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description        = "Select folder to save CSV reports"
    $dlg.ShowNewFolderButton = $true
    if ($txtExportPath.Text.Trim() -and (Test-Path $txtExportPath.Text.Trim())) {
        $dlg.SelectedPath = $txtExportPath.Text.Trim()
    }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtExportPath.Text = $dlg.SelectedPath
    }
})

# Grid filter - live filtering as user types
$txtGridFilter.Add_TextChanged({
    if ($script:ComputerReport.Count -gt 0) {
        Update-DataGridView -Filter $txtGridFilter.Text.Trim()
    }
})

$btnClearFilter.Add_Click({
    $txtGridFilter.Text = ""
})

# Scan
$btnScan.Add_Click({
    $btnScan.Enabled           = $false
    $btnDisable.Enabled        = $false
    $btnDelete.Enabled         = $false
    $btnExport.Enabled         = $false
    $btnReEnable.Enabled       = $false
    $btnSelectEligible.Enabled = $false
    $progressBar.Visible       = $true
    $form.Cursor               = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        # Initialize log file - always read from textbox, never a stale cached variable
        $logDir = $txtLogPath.Text.Trim()
        if ($logDir -and -not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        $script:LogFile = if ($logDir) {
            Join-Path $logDir "StaleADComputers_$(Get-Date -Format 'yyyyMMdd').log"
        } else { $null }

        Write-GUILog "=========================================="
        Write-GUILog "Starting scan for stale computers"
        Write-GUILog "Executed by: $($script:RunningUser)"
        Write-GUILog "Inactive Days Threshold: $($numInactiveDays.Value)"
        Write-GUILog "Delete After Days: $($numDeleteAfterDays.Value)"
        Write-GUILog "Workstations: $($chkWorkstations.Checked)  |  Servers: $($chkServers.Checked)"
        if ($txtSearchBase.Text.Trim()) { Write-GUILog "Search Base: $($txtSearchBase.Text.Trim())" }
        Write-GUILog "=========================================="

        $script:StaleComputers = Get-StaleComputers `
            -DaysInactive        $numInactiveDays.Value `
            -IncludeWorkstations $chkWorkstations.Checked `
            -IncludeServers      $chkServers.Checked `
            -SearchBase          $txtSearchBase.Text.Trim()

        if ($script:StaleComputers.Count -eq 0) {
            Write-GUILog "No stale computers found matching criteria."
            $script:ComputerReport = @()
            $script:ComputerIndex  = @{}
        } else {
            $script:ComputerReport = @($script:StaleComputers | Get-ComputerReportData)
            # Rebuild O(1) lookup index
            $script:ComputerIndex = @{}
            foreach ($c in $script:ComputerReport) { $script:ComputerIndex[$c.Name] = $c }
            Write-GUILog "Generated report for $($script:ComputerReport.Count) computers"
        }

        Update-DataGridView -Filter $txtGridFilter.Text.Trim()

        $btnDisable.Enabled        = $true
        $btnDelete.Enabled         = $true
        $btnExport.Enabled         = $true
        $btnReEnable.Enabled       = $true
        $btnSelectEligible.Enabled = $true
        Write-GUILog "Scan completed."
    }
    catch {
        Write-GUILog "Error during scan: $_" -Level Error
        [System.Windows.Forms.MessageBox]::Show("Error during scan: $_", "Scan Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $btnScan.Enabled     = $true
        $progressBar.Visible = $false
        $form.Cursor         = [System.Windows.Forms.Cursors]::Default
    }
})

# Select All / None
$btnSelectAll.Add_Click({
    foreach ($row in $dataGridView.Rows) { $row.Cells["Select"].Value = $true }
})

$btnSelectNone.Add_Click({
    foreach ($row in $dataGridView.Rows) { $row.Cells["Select"].Value = $false }
})

# Select Eligible
$btnSelectEligible.Add_Click({
    $minDays       = $numDeleteAfterDays.Value
    $eligibleCount = 0

    foreach ($row in $dataGridView.Rows) {
        $name     = $row.Cells["Name"].Value
        $computer = $script:ComputerIndex[$name]

        if ($computer -and -not $computer.Enabled -and (
                ($computer.MarkedForDeletion -and $computer.DaysSinceDisabled -ge $minDays) -or
                (-not $computer.MarkedForDeletion)
            )) {
            $row.Cells["Select"].Value = $true
            $eligibleCount++
        } else {
            $row.Cells["Select"].Value = $false
        }
    }

    Write-GUILog "Selected $eligibleCount computer(s) eligible for deletion"

    if ($eligibleCount -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No computers are currently eligible for deletion.`n`nEligibility requires:`n- Disabled account`n- Either tagged for $minDays+ days or disabled externally",
            "No Eligible Computers",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Disable / Tag
$btnDisable.Add_Click({
    $selected = Get-SelectedComputers

    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No computers selected. Please check the 'Select' column for computers you want to process.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $toDisable      = @($selected | Where-Object { $_.Enabled -and -not $_.MarkedForDeletion })
    $toTag          = @($selected | Where-Object { -not $_.Enabled -and -not $_.MarkedForDeletion })
    $totalToProcess = $toDisable.Count + $toTag.Count

    if ($totalToProcess -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No actionable computers in selection.`n`nAll selected computers are already tagged for deletion.", "Nothing to Process", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $confirmMessage = ""
    if ($toDisable.Count -gt 0) {
        $confirmMessage += "You are about to DISABLE $($toDisable.Count) computer account(s) in Active Directory.`n`nComputers to disable:`n"
        $confirmMessage += ($toDisable | Select-Object -First 5 | ForEach-Object { "  - $($_.Name)" }) -join "`n"
        if ($toDisable.Count -gt 5) { $confirmMessage += "`n  ... and $($toDisable.Count - 5) more" }
        $confirmMessage += "`n`n"
    }
    if ($toTag.Count -gt 0) {
        $confirmMessage += "You will also TAG $($toTag.Count) externally-disabled computer(s) for deletion.`n`nComputers to tag:`n"
        $confirmMessage += ($toTag | Select-Object -First 5 | ForEach-Object { "  - $($_.Name)" }) -join "`n"
        if ($toTag.Count -gt 5) { $confirmMessage += "`n  ... and $($toTag.Count - 5) more" }
        $confirmMessage += "`n`n"
    }
    $confirmMessage += "Do you want to continue?"

    if ([System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirm Disable/Tag", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -eq [System.Windows.Forms.DialogResult]::Yes) {
        $progressBar.Visible = $true
        $btnDisable.Enabled  = $false
        $form.Cursor         = [System.Windows.Forms.Cursors]::WaitCursor

        try {
            $stagingOU  = if ($txtStagingOU.Text.Trim()) { $txtStagingOU.Text.Trim() } else { $null }
            $allResults = @()

            if ($toDisable.Count -gt 0) { $allResults += Disable-SelectedComputers -Computers $toDisable -TargetOU $stagingOU -InactiveDays $numInactiveDays.Value }
            if ($toTag.Count -gt 0)     { $allResults += Set-ComputerDeletionTag   -Computers $toTag     -InactiveDays $numInactiveDays.Value }

            $disabledCount = @($allResults | Where-Object { $_.Action -eq "Disabled" -and $_.Status -eq "Success" }).Count
            $taggedCount   = @($allResults | Where-Object { $_.Action -eq "Tagged"   -and $_.Status -eq "Success" }).Count
            $failCount     = @($allResults | Where-Object { $_.Status -eq "Failed" }).Count

            Write-GUILog "Operation completed: $disabledCount disabled, $taggedCount tagged, $failCount failed"

            if ($allResults.Count -gt 0) {
                $exportDir = $txtExportPath.Text.Trim()
                if ($exportDir -and -not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir -Force | Out-Null }
                $dest = if ($exportDir) { $exportDir } else { $env:TEMP }
                $allResults | Export-Csv -Path (Join-Path $dest "DisableTagActions_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv") -NoTypeInformation
                Write-GUILog "Actions exported to: $dest"
            }

            [System.Windows.Forms.MessageBox]::Show("Operation completed.`n`nDisabled: $disabledCount`nTagged: $taggedCount`nFailed: $failCount", "Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $btnScan.PerformClick()
        }
        catch {
            Write-GUILog "Error during disable/tag: $_" -Level Error
            [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        finally {
            $btnDisable.Enabled  = $true
            $progressBar.Visible = $false
            $form.Cursor         = [System.Windows.Forms.Cursors]::Default
        }
    }
})

# Delete
$btnDelete.Add_Click({
    $selected = Get-SelectedComputers

    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No computers selected.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $minDays  = $numDeleteAfterDays.Value
    $toDelete = @($selected | Where-Object {
        -not $_.Enabled -and (
            ($_.MarkedForDeletion -and $_.DaysSinceDisabled -ge $minDays) -or
            (-not $_.MarkedForDeletion)
        )
    })

    if ($toDelete.Count -eq 0) {
        $notReady = @($selected | Where-Object { $_.MarkedForDeletion -and $_.DaysSinceDisabled -lt $minDays })
        if ($notReady.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show("Selected computers are marked for deletion but haven't been disabled long enough.`n`nRequired: $minDays days`nCurrent: $($notReady[0].DaysSinceDisabled) days", "Not Ready", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("No disabled computers in selection.`n`nOnly disabled computers can be deleted.", "Nothing to Delete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        return
    }

    $externalCount = ($toDelete | Where-Object { -not $_.MarkedForDeletion }).Count
    $msg  = "WARNING: You are about to PERMANENTLY DELETE $($toDelete.Count) computer account(s).`n`nTHIS ACTION CANNOT BE UNDONE!`n`n"
    if ($externalCount -gt 0) { $msg += "NOTE: $externalCount computer(s) were disabled externally (not by this script).`n`n" }
    $msg += "Computers to delete:`n"
    $msg += ($toDelete | Select-Object -First 10 | ForEach-Object {
        $days = if ($_.DaysSinceDisabled) { "disabled $($_.DaysSinceDisabled) days" } else { "disabled externally" }
        "  - $($_.Name) ($days)"
    }) -join "`n"
    if ($toDelete.Count -gt 10) { $msg += "`n  ... and $($toDelete.Count - 10) more" }
    $msg += "`n`nType 'DELETE' in the next dialog to confirm."

    if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning) -eq [System.Windows.Forms.DialogResult]::OK) {

        $inputForm                 = New-Object System.Windows.Forms.Form
        $inputForm.Text            = "Confirm Deletion"
        $inputForm.Size            = New-Object System.Drawing.Size(350, 150)
        $inputForm.StartPosition   = "CenterParent"
        $inputForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $inputForm.MaximizeBox     = $false
        $inputForm.MinimizeBox     = $false

        $lblConfirm          = New-Object System.Windows.Forms.Label
        $lblConfirm.Text     = "Type 'DELETE' to confirm permanent deletion:"
        $lblConfirm.Location = New-Object System.Drawing.Point(10, 20)
        $lblConfirm.AutoSize = $true

        $txtConfirm          = New-Object System.Windows.Forms.TextBox
        $txtConfirm.Location = New-Object System.Drawing.Point(10, 50)
        $txtConfirm.Size     = New-Object System.Drawing.Size(310, 25)

        $btnConfirmOK              = New-Object System.Windows.Forms.Button
        $btnConfirmOK.Text         = "Confirm"
        $btnConfirmOK.Location     = New-Object System.Drawing.Point(160, 80)
        $btnConfirmOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $btnConfirmCancel              = New-Object System.Windows.Forms.Button
        $btnConfirmCancel.Text         = "Cancel"
        $btnConfirmCancel.Location     = New-Object System.Drawing.Point(245, 80)
        $btnConfirmCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

        $inputForm.Controls.AddRange(@($lblConfirm, $txtConfirm, $btnConfirmOK, $btnConfirmCancel))
        $inputForm.AcceptButton = $btnConfirmOK
        $inputForm.CancelButton = $btnConfirmCancel

        if ($inputForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $txtConfirm.Text -eq "DELETE") {
            $progressBar.Visible = $true
            $btnDelete.Enabled   = $false
            $form.Cursor         = [System.Windows.Forms.Cursors]::WaitCursor

            try {
                $results      = Remove-SelectedComputers -Computers $toDelete -MinDaysDisabled $minDays
                $successCount = @($results | Where-Object { $_.Status -eq "Success" }).Count
                $failCount    = @($results | Where-Object { $_.Status -eq "Failed" }).Count

                Write-GUILog "Delete completed: $successCount succeeded, $failCount failed"

                if ($results.Count -gt 0) {
                    $exportDir = $txtExportPath.Text.Trim()
                    if ($exportDir -and -not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir -Force | Out-Null }
                    $dest = if ($exportDir) { $exportDir } else { $env:TEMP }
                    $results | Export-Csv -Path (Join-Path $dest "DeleteActions_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv") -NoTypeInformation
                    Write-GUILog "Actions exported to: $dest"
                }

                [System.Windows.Forms.MessageBox]::Show("Delete completed.`n`nSucceeded: $successCount`nFailed: $failCount", "Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $btnScan.PerformClick()
            }
            catch {
                Write-GUILog "Error during delete: $_" -Level Error
                [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            finally {
                $btnDelete.Enabled   = $true
                $progressBar.Visible = $false
                $form.Cursor         = [System.Windows.Forms.Cursors]::Default
            }
        } else {
            Write-GUILog "Delete cancelled - confirmation text not entered" -Level Warning
        }
    }
})

# Re-enable
$btnReEnable.Add_Click({
    $selected   = Get-SelectedComputers
    $toReEnable = @($selected | Where-Object { -not $_.Enabled })

    if ($toReEnable.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No disabled computers selected.`n`nRe-enable only applies to disabled computer accounts.", "Nothing to Re-enable", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $msg  = "You are about to RE-ENABLE $($toReEnable.Count) computer account(s).`n`nComputers to re-enable:`n"
    $msg += ($toReEnable | Select-Object -First 10 | ForEach-Object { "  - $($_.Name)" }) -join "`n"
    if ($toReEnable.Count -gt 10) { $msg += "`n  ... and $($toReEnable.Count - 10) more" }
    $msg += "`n`nOriginal descriptions will be restored where possible. Continue?"

    if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm Re-enable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -eq [System.Windows.Forms.DialogResult]::Yes) {
        $progressBar.Visible = $true
        $btnReEnable.Enabled = $false
        $form.Cursor         = [System.Windows.Forms.Cursors]::WaitCursor

        try {
            $results      = Restore-SelectedComputers -Computers $toReEnable
            $successCount = @($results | Where-Object { $_.Status -eq "Success" }).Count
            $failCount    = @($results | Where-Object { $_.Status -eq "Failed" }).Count

            Write-GUILog "Re-enable completed: $successCount succeeded, $failCount failed"

            if ($results.Count -gt 0) {
                $exportDir = $txtExportPath.Text.Trim()
                if ($exportDir -and -not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir -Force | Out-Null }
                $dest = if ($exportDir) { $exportDir } else { $env:TEMP }
                $results | Export-Csv -Path (Join-Path $dest "ReEnableActions_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv") -NoTypeInformation
                Write-GUILog "Actions exported to: $dest"
            }

            [System.Windows.Forms.MessageBox]::Show("Re-enable completed.`n`nSucceeded: $successCount`nFailed: $failCount", "Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $btnScan.PerformClick()
        }
        catch {
            Write-GUILog "Error during re-enable: $_" -Level Error
            [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        finally {
            $btnReEnable.Enabled = $true
            $progressBar.Visible = $false
            $form.Cursor         = [System.Windows.Forms.Cursors]::Default
        }
    }
})

# Export to CSV
$btnExport.Add_Click({
    if ($script:ComputerReport.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export. Please scan for stale computers first.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $exportFolder = $txtExportPath.Text.Trim()
    if (-not $exportFolder) {
        $dlg                    = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description        = "Select folder to save CSV report"
        $dlg.ShowNewFolderButton = $true
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportFolder       = $dlg.SelectedPath
            $txtExportPath.Text = $exportFolder
        } else { return }
    }

    $exportPath = Join-Path $exportFolder "StaleComputers_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

    try {
        if (-not (Test-Path $exportFolder)) { New-Item -ItemType Directory -Path $exportFolder -Force | Out-Null }

        $script:ComputerReport | Select-Object Name, Enabled, DistinguishedName, LastLogonDate,
            DaysSinceLogon, PasswordLastSet, WhenCreated, OperatingSystem,
            Description, MarkedForDeletion, DisableDate, DaysSinceDisabled |
            Export-Csv -Path $exportPath -NoTypeInformation

        Write-GUILog "Report exported to: $exportPath"

        $open = [System.Windows.Forms.MessageBox]::Show(
            "Report exported successfully to:`n$exportPath`n`nOpen export folder?",
            "Export Complete",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information)

        if ($open -eq [System.Windows.Forms.DialogResult]::Yes) {
            Start-Process explorer.exe $exportFolder
        }
    }
    catch {
        Write-GUILog "Failed to export report: $_" -Level Error
        [System.Windows.Forms.MessageBox]::Show("Failed to export report: $_", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Row color coding by lifecycle stage
$dataGridView.Add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }

    $row     = $dataGridView.Rows[$e.RowIndex]
    $enabled = $row.Cells["Enabled"].Value
    $marked  = $row.Cells["MarkedForDeletion"].Value
    $daysVal = $row.Cells["DaysSinceDisabled"].Value
    $days    = if ($daysVal -and $daysVal -ne "") { [int]$daysVal } else { 0 }
    $minDays = [int]$numDeleteAfterDays.Value

    $bg = if (-not $enabled -and $marked -and $days -ge $minDays) {
        [System.Drawing.Color]::FromArgb(255, 204, 188)   # salmon  - ready to delete
    } elseif (-not $enabled -and $marked) {
        [System.Drawing.Color]::FromArgb(255, 243, 205)   # yellow  - tagged, still waiting
    } elseif (-not $enabled -and -not $marked) {
        [System.Drawing.Color]::FromArgb(255, 224, 178)   # orange  - externally disabled
    } else {
        [System.Drawing.Color]::White
    }

    $e.CellStyle.BackColor          = $bg
    $e.CellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, [int]$bg.R - 40),
        [Math]::Max(0, [int]$bg.G - 40),
        [Math]::Max(0, [int]$bg.B - 40)
    )
})

# Right-click: if clicked row is not already checked, clear others and select just that row
$dataGridView.Add_CellMouseDown({
    param($sender, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0) {
        $clickedRow = $dataGridView.Rows[$e.RowIndex]
        if ($clickedRow.Cells["Select"].Value -ne $true) {
            foreach ($r in $dataGridView.Rows) { $r.Cells["Select"].Value = $false }
            $clickedRow.Cells["Select"].Value = $true
        }
    }
})

# Context menu items
$ctxDisable.Add_Click({  $btnDisable.PerformClick()  })
$ctxDelete.Add_Click({   $btnDelete.PerformClick()   })
$ctxReEnable.Add_Click({ $btnReEnable.PerformClick() })

$ctxCopyName.Add_Click({
    $sel = Get-SelectedComputers
    if ($sel.Count -gt 0) {
        [System.Windows.Forms.Clipboard]::SetText(($sel | ForEach-Object { $_.Name }) -join "`n")
        Write-GUILog "Copied $($sel.Count) computer name(s) to clipboard"
    }
})

$ctxCopyDN.Add_Click({
    $sel = Get-SelectedComputers
    if ($sel.Count -gt 0) {
        [System.Windows.Forms.Clipboard]::SetText(($sel | ForEach-Object { $_.DistinguishedName }) -join "`n")
        Write-GUILog "Copied $($sel.Count) distinguished name(s) to clipboard"
    }
})

#endregion Event Handlers

Write-GUILog "Stale AD Computer Management GUI v3.0 initialized"
[void]$form.ShowDialog()
