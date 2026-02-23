<#
.SYNOPSIS
    GUI version - Identifies, disables, and deletes stale computer objects in Active Directory.

.DESCRIPTION
    Provides a graphical interface for managing stale AD computer objects.
    Implements a 3-stage lifecycle for stale AD computer management:
    Stage 1: Identify workstations inactive for X days (default 365)
    Stage 2: Disable workstations and move to staging OU, tag with disable date
    Stage 3: Delete workstations that have been disabled for Y days (default 30)

    This script automatically targets only Windows workstation operating systems
    (Windows 10, Windows 11) and excludes Windows Server systems.

.NOTES
    Author: Alberto de la Torre
    Version: 1.0
    Date: February 2026
#>

#Requires -Modules ActiveDirectory

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Script Variables
$script:LogPath = "C:\Temp\StaleADComputerLogs"
$script:LogFile = $null
$script:StaleComputers = @()
$script:ComputerReport = @()
#endregion

#region Functions
function Write-GUILog {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Add to log textbox
    $txtLog.AppendText("$logMessage`r`n")
    $txtLog.ScrollToCaret()
    
    # Write to log file if exists
    if ($script:LogFile -and (Test-Path (Split-Path $script:LogFile -Parent))) {
        Add-Content -Path $script:LogFile -Value $logMessage
    }
    
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-StaleComputers {
    param(
        [int]$DaysInactive,
        [bool]$IncludeWorkstations,
        [bool]$IncludeServers
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
    Write-GUILog "Searching for computers inactive since $($cutoffDate.ToString('yyyy-MM-dd')) matching OS filter."
    
    # Get computers (filter by OS at query level for efficiency)
    $allComputers = Get-ADComputer -Filter $filterString -Properties `
        Name, 
    DistinguishedName, 
    LastLogonTimestamp, 
    LastLogonDate,
    pwdLastSet, 
    whenChanged, 
    whenCreated,
    Enabled, 
    Description, 
    OperatingSystem,
    OperatingSystemVersion
    
    Write-GUILog "Found $($allComputers.Count) computers matching filter in AD"
    
    # Filter for stale computers
    $staleComputers = $allComputers | Where-Object {
        # Check LastLogonTimestamp (most reliable for staleness)
        $lastLogon = if ($_.LastLogonTimestamp -and $_.LastLogonTimestamp -ne 0) {
            [DateTime]::FromFileTime($_.LastLogonTimestamp)
        }
        else {
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
        }
        else {
            $null
        }
        
        $pwdLastSet = if ($Computer.pwdLastSet -and $Computer.pwdLastSet -ne 0) {
            [DateTime]::FromFileTime($Computer.pwdLastSet)
        }
        else {
            $null
        }
        
        $daysSinceLogon = if ($lastLogon) {
            [math]::Round(((Get-Date) - $lastLogon).TotalDays, 0)
        }
        else {
            -1  # Use -1 for "Never" to allow sorting
        }
        
        # Check if marked for deletion (from description)
        $markedForDeletion = $false
        $disableDate = $null
        if ($Computer.Description -match "DISABLED:\s*(\d{4}-\d{2}-\d{2})") {
            $markedForDeletion = $true
            $disableDate = [DateTime]::Parse($Matches[1])
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
    $today = Get-Date -Format "yyyy-MM-dd"
    
    foreach ($computer in $Computers) {
        $computerName = $computer.Name
        
        try {
            $adComputer = Get-ADComputer -Identity $computerName -Properties Description
            $originalDescription = $adComputer.Description
            $newDescription = "DISABLED: $today | Original: $originalDescription | Stale $InactiveDays+ days"
            
            # Disable the computer account
            Set-ADComputer -Identity $adComputer.DistinguishedName -Enabled $false -Description $newDescription
            Write-GUILog "Disabled computer: $computerName"
            
            # Move to staging OU if specified
            if ($TargetOU) {
                Move-ADObject -Identity $adComputer.DistinguishedName -TargetPath $TargetOU
                Write-GUILog "Moved $computerName to staging OU"
            }
            
            $results += [PSCustomObject]@{
                ComputerName = $computerName
                Action       = "Disabled"
                Status       = "Success"
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
        
        # Check if computer has been disabled long enough (only if marked by this script)
        # Externally disabled computers (DaysSinceDisabled is null) can be deleted immediately
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
                Timestamp    = Get-Date
                Error        = $_.Exception.Message
            }
        }
    }
    
    return $results
}

function Tag-SelectedComputers {
    param(
        [array]$Computers,
        [int]$InactiveDays
    )
    
    $results = @()
    $today = Get-Date -Format "yyyy-MM-dd"
    
    foreach ($computer in $Computers) {
        $computerName = $computer.Name
        
        try {
            $adComputer = Get-ADComputer -Identity $computerName -Properties Description
            $originalDescription = $adComputer.Description
            
            # Add DISABLED tag to description
            $newDescription = "DISABLED: $today | Original: $originalDescription | Tagged for deletion (stale $InactiveDays+ days)"
            
            Set-ADComputer -Identity $adComputer.DistinguishedName -Description $newDescription
            Write-GUILog "Tagged computer for deletion: $computerName"
            
            $results += [PSCustomObject]@{
                ComputerName   = $computerName
                Action         = "Tagged"
                Status         = "Success"
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
                Timestamp    = Get-Date
                Error        = $_.Exception.Message
            }
        }
    }
    
    return $results
}

function Update-DataGridView {
    $dataGridView.DataSource = $null
    
    if ($script:ComputerReport.Count -gt 0) {
        # Create DataTable for binding
        $dataTable = New-Object System.Data.DataTable
        
        # Add columns
        $dataTable.Columns.Add("Select", [bool]) | Out-Null
        $dataTable.Columns.Add("Name", [string]) | Out-Null
        $dataTable.Columns.Add("Enabled", [bool]) | Out-Null
        $dataTable.Columns.Add("LastLogonDate", [string]) | Out-Null
        $dataTable.Columns.Add("DaysSinceLogon", [string]) | Out-Null
        $dataTable.Columns.Add("OperatingSystem", [string]) | Out-Null
        $dataTable.Columns.Add("MarkedForDeletion", [bool]) | Out-Null
        $dataTable.Columns.Add("DisableDate", [string]) | Out-Null
        $dataTable.Columns.Add("DaysSinceDisabled", [string]) | Out-Null
        $dataTable.Columns.Add("DistinguishedName", [string]) | Out-Null
        
        foreach ($computer in $script:ComputerReport) {
            $row = $dataTable.NewRow()
            $row["Select"] = $false
            $row["Name"] = $computer.Name
            $row["Enabled"] = $computer.Enabled
            $row["LastLogonDate"] = $computer.LastLogonDate
            $row["DaysSinceLogon"] = $computer.DaysSinceLogon
            $row["OperatingSystem"] = $computer.OperatingSystem
            $row["MarkedForDeletion"] = $computer.MarkedForDeletion
            $row["DisableDate"] = $computer.DisableDate
            $row["DaysSinceDisabled"] = if ($computer.DaysSinceDisabled) { $computer.DaysSinceDisabled.ToString() } else { "" }
            $row["DistinguishedName"] = $computer.DistinguishedName
            $dataTable.Rows.Add($row)
        }
        
        $dataGridView.DataSource = $dataTable
        
        # Format columns
        $dataGridView.Columns["Select"].Width = 50
        $dataGridView.Columns["Name"].Width = 120
        $dataGridView.Columns["Enabled"].Width = 60
        $dataGridView.Columns["LastLogonDate"].Width = 100
        $dataGridView.Columns["DaysSinceLogon"].Width = 80
        $dataGridView.Columns["OperatingSystem"].Width = 150
        $dataGridView.Columns["MarkedForDeletion"].Width = 100
        $dataGridView.Columns["DisableDate"].Width = 90
        $dataGridView.Columns["DaysSinceDisabled"].Width = 100
        $dataGridView.Columns["DistinguishedName"].Width = 300
        
        # Update summary labels
        $totalCount = $script:ComputerReport.Count
        $enabledCount = ($script:ComputerReport | Where-Object { $_.Enabled }).Count
        $disabledCount = ($script:ComputerReport | Where-Object { -not $_.Enabled }).Count
        $markedCount = ($script:ComputerReport | Where-Object { $_.MarkedForDeletion }).Count
        
        $lblSummary.Text = "Total: $totalCount | Enabled: $enabledCount | Disabled: $disabledCount | Marked for Deletion: $markedCount"
    }
}

function Get-SelectedComputers {
    $selected = @()
    
    foreach ($row in $dataGridView.Rows) {
        if ($row.Cells["Select"].Value -eq $true) {
            $computerName = $row.Cells["Name"].Value
            $computer = $script:ComputerReport | Where-Object { $_.Name -eq $computerName }
            if ($computer) {
                $selected += $computer
            }
        }
    }
    
    return $selected
}
#endregion Functions

#region Build Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Stale AD Computer Management"
$form.Size = New-Object System.Drawing.Size(1200, 800)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)



# Main TableLayoutPanel
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainLayout.ColumnCount = 1
$mainLayout.RowCount = 4
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 180))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 150))) | Out-Null
$mainLayout.Padding = New-Object System.Windows.Forms.Padding(10)

#region Settings Panel
$settingsGroup = New-Object System.Windows.Forms.GroupBox
$settingsGroup.Text = "Settings"
$settingsGroup.Dock = [System.Windows.Forms.DockStyle]::Fill

$settingsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$settingsLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$settingsLayout.ColumnCount = 6
$settingsLayout.RowCount = 4
$settingsLayout.Padding = New-Object System.Windows.Forms.Padding(5)

# Row 1: Inactive Days, Delete After Days, Staging OU
$lblInactiveDays = New-Object System.Windows.Forms.Label
$lblInactiveDays.Text = "Inactive Days:"
$lblInactiveDays.AutoSize = $true
$lblInactiveDays.Anchor = [System.Windows.Forms.AnchorStyles]::Left

$numInactiveDays = New-Object System.Windows.Forms.NumericUpDown
$numInactiveDays.Minimum = 30
$numInactiveDays.Maximum = 730
$numInactiveDays.Value = 365
$numInactiveDays.Width = 80

$lblDeleteAfterDays = New-Object System.Windows.Forms.Label
$lblDeleteAfterDays.Text = "Delete After Days:"
$lblDeleteAfterDays.AutoSize = $true
$lblDeleteAfterDays.Anchor = [System.Windows.Forms.AnchorStyles]::Left

$numDeleteAfterDays = New-Object System.Windows.Forms.NumericUpDown
$numDeleteAfterDays.Minimum = 7
$numDeleteAfterDays.Maximum = 365
$numDeleteAfterDays.Value = 30
$numDeleteAfterDays.Width = 80

$lblStagingOU = New-Object System.Windows.Forms.Label
$lblStagingOU.Text = "Staging OU (DN):"
$lblStagingOU.AutoSize = $true
$lblStagingOU.Anchor = [System.Windows.Forms.AnchorStyles]::Right

$txtStagingOU = New-Object System.Windows.Forms.TextBox
$txtStagingOU.Width = 350
$txtStagingOU.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$lblOSTypes = New-Object System.Windows.Forms.Label
$lblOSTypes.Text = "Include OS Types:"
$lblOSTypes.AutoSize = $true
$lblOSTypes.Anchor = [System.Windows.Forms.AnchorStyles]::Left

$chkWorkstations = New-Object System.Windows.Forms.CheckBox
$chkWorkstations.Text = "Workstations (Win 10/11)"
$chkWorkstations.Checked = $true
$chkWorkstations.AutoSize = $true

$chkServers = New-Object System.Windows.Forms.CheckBox
$chkServers.Text = "Servers"
$chkServers.Checked = $false
$chkServers.AutoSize = $true

# Row 2: Log Path, Export Folder
$lblLogPath = New-Object System.Windows.Forms.Label
$lblLogPath.Text = "Log Path:"
$lblLogPath.AutoSize = $true
$lblLogPath.Anchor = [System.Windows.Forms.AnchorStyles]::Left

$txtLogPath = New-Object System.Windows.Forms.TextBox
$txtLogPath.Text = "C:\Temp\StaleADComputerLogs"
$txtLogPath.Width = 250

$lblExportPath = New-Object System.Windows.Forms.Label
$lblExportPath.Text = "Export Folder:"
$lblExportPath.AutoSize = $true
$lblExportPath.Anchor = [System.Windows.Forms.AnchorStyles]::Left

$txtExportPath = New-Object System.Windows.Forms.TextBox
$txtExportPath.Width = 300
$txtExportPath.Text = "C:\Temp\StaleADComputerReports"

$btnBrowseExport = New-Object System.Windows.Forms.Button
$btnBrowseExport.Text = "..."
$btnBrowseExport.Width = 30

# Row 3: Action Buttons
$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan for Stale Computers"
$btnScan.Width = 180
$btnScan.Height = 30
$btnScan.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnScan.ForeColor = [System.Drawing.Color]::White
$btnScan.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

$btnDisable = New-Object System.Windows.Forms.Button
$btnDisable.Text = "Disable/Tag"
$btnDisable.Width = 100
$btnDisable.Height = 30
$btnDisable.BackColor = [System.Drawing.Color]::FromArgb(255, 185, 0)
$btnDisable.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnDisable.Enabled = $false

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = "Delete Selected"
$btnDelete.Width = 120
$btnDelete.Height = 30
$btnDelete.BackColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
$btnDelete.ForeColor = [System.Drawing.Color]::White
$btnDelete.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnDelete.Enabled = $false

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export to CSV"
$btnExport.Width = 120
$btnExport.Height = 30
$btnExport.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
$btnExport.ForeColor = [System.Drawing.Color]::White
$btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExport.Enabled = $false

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Select All"
$btnSelectAll.Width = 80
$btnSelectAll.Height = 30

$btnSelectNone = New-Object System.Windows.Forms.Button
$btnSelectNone.Text = "Select None"
$btnSelectNone.Width = 80
$btnSelectNone.Height = 30

# Add controls to settings layout
$settingsLayout.Controls.Add($lblInactiveDays, 0, 0)
$settingsLayout.Controls.Add($numInactiveDays, 1, 0)
$settingsLayout.Controls.Add($lblDeleteAfterDays, 2, 0)
$settingsLayout.Controls.Add($numDeleteAfterDays, 3, 0)
$settingsLayout.Controls.Add($lblStagingOU, 4, 0)
$settingsLayout.Controls.Add($txtStagingOU, 5, 0)

$settingsLayout.Controls.Add($lblLogPath, 0, 1)
$settingsLayout.Controls.Add($txtLogPath, 1, 1)
$settingsLayout.SetColumnSpan($txtLogPath, 2)
$settingsLayout.Controls.Add($lblOSTypes, 3, 1)
$osPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$osPanel.AutoSize = $true
$osPanel.Controls.Add($chkWorkstations)
$osPanel.Controls.Add($chkServers)
$settingsLayout.Controls.Add($osPanel, 4, 1)
$settingsLayout.SetColumnSpan($osPanel, 2)

# Row 3: Export Path
$settingsLayout.Controls.Add($lblExportPath, 0, 2)
$settingsLayout.Controls.Add($txtExportPath, 1, 2)
$settingsLayout.SetColumnSpan($txtExportPath, 4)
$settingsLayout.Controls.Add($btnBrowseExport, 5, 2)

# Row 4: Action Buttons
$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$buttonPanel.Controls.Add($btnScan)
$buttonPanel.Controls.Add($btnSelectAll)
$buttonPanel.Controls.Add($btnSelectNone)
$buttonPanel.Controls.Add($btnDisable)
$buttonPanel.Controls.Add($btnDelete)
$buttonPanel.Controls.Add($btnExport)
$settingsLayout.Controls.Add($buttonPanel, 0, 3)
$settingsLayout.SetColumnSpan($buttonPanel, 6)

$settingsGroup.Controls.Add($settingsLayout)
#endregion Settings Panel

#region Summary Label
$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Text = "Total: 0 | Enabled: 0 | Disabled: 0 | Marked for Deletion: 0"
$lblSummary.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblSummary.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblSummary.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
#endregion

#region DataGridView
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Dock = [System.Windows.Forms.DockStyle]::Fill
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.MultiSelect = $true
$dataGridView.ReadOnly = $false
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$dataGridView.RowHeadersVisible = $false
$dataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
#endregion DataGridView

#region Log TextBox
$logGroup = New-Object System.Windows.Forms.GroupBox
$logGroup.Text = "Activity Log"
$logGroup.Dock = [System.Windows.Forms.DockStyle]::Fill

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$txtLog.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)

$logGroup.Controls.Add($txtLog)
#endregion Log TextBox

# Add panels to main layout
$mainLayout.Controls.Add($settingsGroup, 0, 0)
$mainLayout.Controls.Add($lblSummary, 0, 1)
$mainLayout.Controls.Add($dataGridView, 0, 2)
$mainLayout.Controls.Add($logGroup, 0, 3)

$form.Controls.Add($mainLayout)
#endregion Build Form

#region Event Handlers
$btnBrowseExport.Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Select folder to save CSV reports"
        $folderDialog.ShowNewFolderButton = $true
    
        if ($txtExportPath.Text.Trim() -and (Test-Path $txtExportPath.Text.Trim())) {
            $folderDialog.SelectedPath = $txtExportPath.Text.Trim()
        }
    
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtExportPath.Text = $folderDialog.SelectedPath
        }
    })

$btnScan.Add_Click({
        $btnScan.Enabled = $false
        $btnDisable.Enabled = $false
        $btnDelete.Enabled = $false
        $btnExport.Enabled = $false
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
        try {
            # Initialize log file
            $script:LogPath = $txtLogPath.Text
            if (-not (Test-Path $script:LogPath)) {
                New-Item -ItemType Directory -Path $script:LogPath -Force | Out-Null
            }
            $timestamp = Get-Date -Format "yyyyMMdd"
            $script:LogFile = Join-Path $script:LogPath "StaleADComputers_$timestamp.log"
        
            Write-GUILog "=========================================="
            Write-GUILog "Starting scan for stale computers"
            Write-GUILog "Inactive Days Threshold: $($numInactiveDays.Value)"
            Write-GUILog "Included Workstations: $($chkWorkstations.Checked)"
            Write-GUILog "Included Servers: $($chkServers.Checked)"
            Write-GUILog "=========================================="
        
            # Get stale computers
            $script:StaleComputers = Get-StaleComputers -DaysInactive $numInactiveDays.Value `
                -IncludeWorkstations $chkWorkstations.Checked `
                -IncludeServers $chkServers.Checked
        
            if ($script:StaleComputers.Count -eq 0) {
                Write-GUILog "No stale computers found matching criteria."
                $script:ComputerReport = @()
            }
            else {
                # Generate report
                $script:ComputerReport = $script:StaleComputers | Get-ComputerReportData
                Write-GUILog "Generated report for $($script:ComputerReport.Count) computers"
            }
        
            # Update grid
            Update-DataGridView
        
            $btnDisable.Enabled = $true
            $btnDelete.Enabled = $true
            $btnExport.Enabled = $true
        
            Write-GUILog "Scan completed."
        }
        catch {
            Write-GUILog "Error during scan: $_" -Level Error
            [System.Windows.Forms.MessageBox]::Show("Error during scan: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        finally {
            $btnScan.Enabled = $true
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

$btnSelectAll.Add_Click({
        foreach ($row in $dataGridView.Rows) {
            $row.Cells["Select"].Value = $true
        }
    })

$btnSelectNone.Add_Click({
        foreach ($row in $dataGridView.Rows) {
            $row.Cells["Select"].Value = $false
        }
    })

$btnDisable.Add_Click({
        $selected = Get-SelectedComputers
    
        if ($selected.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No computers selected. Please check the 'Select' column for computers you want to process.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    
        # Filter to enabled computers (to disable) and externally-disabled untagged computers (to tag)
        $toDisable = @($selected | Where-Object { $_.Enabled -and -not $_.MarkedForDeletion })
        $toTag = @($selected | Where-Object { -not $_.Enabled -and -not $_.MarkedForDeletion })
    
        $totalToProcess = $toDisable.Count + $toTag.Count
    
        if ($totalToProcess -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No actionable computers in selection.`n`nAll selected computers are already tagged for deletion.", "Nothing to Process", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
    
        $confirmMessage = ""
        if ($toDisable.Count -gt 0) {
            $confirmMessage += "You are about to DISABLE $($toDisable.Count) computer account(s) in Active Directory.`n`n"
            $confirmMessage += "Computers to disable:`n"
            $confirmMessage += ($toDisable | Select-Object -First 5 | ForEach-Object { "  - $($_.Name)" }) -join "`n"
            if ($toDisable.Count -gt 5) {
                $confirmMessage += "`n  ... and $($toDisable.Count - 5) more"
            }
            $confirmMessage += "`n`n"
        }
        if ($toTag.Count -gt 0) {
            $confirmMessage += "You will also TAG $($toTag.Count) externally-disabled computer(s) for deletion.`n`n"
            $confirmMessage += "Computers to tag:`n"
            $confirmMessage += ($toTag | Select-Object -First 5 | ForEach-Object { "  - $($_.Name)" }) -join "`n"
            if ($toTag.Count -gt 5) {
                $confirmMessage += "`n  ... and $($toTag.Count - 5) more"
            }
            $confirmMessage += "`n`n"
        }
        $confirmMessage += "Do you want to continue?"
    
        $result = [System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirm Disable/Tag", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $btnDisable.Enabled = $false
        
            try {
                $stagingOU = if ($txtStagingOU.Text.Trim()) { $txtStagingOU.Text.Trim() } else { $null }
                $allResults = @()
            
                # Disable enabled computers
                if ($toDisable.Count -gt 0) {
                    $disableResults = Disable-SelectedComputers -Computers $toDisable -TargetOU $stagingOU -InactiveDays $numInactiveDays.Value
                    $allResults += $disableResults
                }
            
                # Tag externally-disabled computers
                if ($toTag.Count -gt 0) {
                    $tagResults = Tag-SelectedComputers -Computers $toTag -InactiveDays $numInactiveDays.Value
                    $allResults += $tagResults
                }
            
                $successCount = @($allResults | Where-Object { $_.Status -eq "Success" }).Count
                $failCount = @($allResults | Where-Object { $_.Status -eq "Failed" }).Count
            
                $disabledCount = @($allResults | Where-Object { $_.Action -eq "Disabled" -and $_.Status -eq "Success" }).Count
                $taggedCount = @($allResults | Where-Object { $_.Action -eq "Tagged" -and $_.Status -eq "Success" }).Count
            
                Write-GUILog "Operation completed: $disabledCount disabled, $taggedCount tagged, $failCount failed"
            
                # Export results
                if ($allResults.Count -gt 0) {
                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $actionFile = Join-Path $script:LogPath "DisableTagActions_$timestamp.csv"
                    $allResults | Export-Csv -Path $actionFile -NoTypeInformation
                    Write-GUILog "Actions logged to: $actionFile"
                }
            
                [System.Windows.Forms.MessageBox]::Show("Operation completed.`n`nDisabled: $disabledCount`nTagged: $taggedCount`nFailed: $failCount", "Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
                # Refresh the grid
                $btnScan.PerformClick()
            }
            catch {
                Write-GUILog "Error during operation: $_" -Level Error
                [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            finally {
                $btnDisable.Enabled = $true
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }
    })

$btnDelete.Add_Click({
        $selected = Get-SelectedComputers
    
        if ($selected.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No computers selected. Please check the 'Select' column for computers you want to delete.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    
        # Filter to only disabled computers
        # If MarkedForDeletion, check DaysSinceDisabled threshold
        # If not marked (disabled externally), still allow deletion
        $minDays = $numDeleteAfterDays.Value
        $toDelete = $selected | Where-Object { 
            -not $_.Enabled -and (
                # Either: marked for deletion and past threshold
                ($_.MarkedForDeletion -and $_.DaysSinceDisabled -ge $minDays) -or
                # Or: disabled externally (no tag) - allow deletion
                (-not $_.MarkedForDeletion)
            )
        }
    
        # Count how many are externally disabled (no tag)
        $externallyDisabled = ($toDelete | Where-Object { -not $_.MarkedForDeletion }).Count
        $markedAndReady = ($toDelete | Where-Object { $_.MarkedForDeletion }).Count
    
        if ($toDelete.Count -eq 0) {
            # Check if any are marked but not old enough
            $notReady = $selected | Where-Object { $_.MarkedForDeletion -and $_.DaysSinceDisabled -lt $minDays }
            if ($notReady.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("Selected computers are marked for deletion but haven't been disabled long enough.`n`nRequired: $minDays days`nCurrent: $($notReady[0].DaysSinceDisabled) days", "Not Ready", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("No disabled computers in selection.`n`nOnly disabled computers can be deleted.", "Nothing to Delete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            return
        }
    
        $confirmMessage = "WARNING: You are about to PERMANENTLY DELETE $($toDelete.Count) computer account(s) from Active Directory.`n`n"
        $confirmMessage += "THIS ACTION CANNOT BE UNDONE!`n`n"
        if ($externallyDisabled -gt 0) {
            $confirmMessage += "NOTE: $externallyDisabled computer(s) were disabled externally (not by this script).`n`n"
        }
        $confirmMessage += "Computers to delete:`n"
        $confirmMessage += ($toDelete | Select-Object -First 10 | ForEach-Object { 
                $days = if ($_.DaysSinceDisabled) { "disabled $($_.DaysSinceDisabled) days" } else { "disabled externally" }
                "  - $($_.Name) ($days)" 
            }) -join "`n"
        if ($toDelete.Count -gt 10) {
            $confirmMessage += "`n  ... and $($toDelete.Count - 10) more"
        }
        $confirmMessage += "`n`nType 'DELETE' in the next dialog to confirm."
    
        $result = [System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
    
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            # Second confirmation with text input
            $inputForm = New-Object System.Windows.Forms.Form
            $inputForm.Text = "Confirm Deletion"
            $inputForm.Size = New-Object System.Drawing.Size(350, 150)
            $inputForm.StartPosition = "CenterParent"
            $inputForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $inputForm.MaximizeBox = $false
            $inputForm.MinimizeBox = $false
        
            $lblConfirm = New-Object System.Windows.Forms.Label
            $lblConfirm.Text = "Type 'DELETE' to confirm permanent deletion:"
            $lblConfirm.Location = New-Object System.Drawing.Point(10, 20)
            $lblConfirm.AutoSize = $true
        
            $txtConfirm = New-Object System.Windows.Forms.TextBox
            $txtConfirm.Location = New-Object System.Drawing.Point(10, 50)
            $txtConfirm.Size = New-Object System.Drawing.Size(310, 25)
        
            $btnConfirmOK = New-Object System.Windows.Forms.Button
            $btnConfirmOK.Text = "Confirm"
            $btnConfirmOK.Location = New-Object System.Drawing.Point(160, 80)
            $btnConfirmOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
        
            $btnConfirmCancel = New-Object System.Windows.Forms.Button
            $btnConfirmCancel.Text = "Cancel"
            $btnConfirmCancel.Location = New-Object System.Drawing.Point(245, 80)
            $btnConfirmCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        
            $inputForm.Controls.AddRange(@($lblConfirm, $txtConfirm, $btnConfirmOK, $btnConfirmCancel))
            $inputForm.AcceptButton = $btnConfirmOK
            $inputForm.CancelButton = $btnConfirmCancel
        
            if ($inputForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $txtConfirm.Text -eq "DELETE") {
                $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
                $btnDelete.Enabled = $false
            
                try {
                    $results = Remove-SelectedComputers -Computers $toDelete -MinDaysDisabled $minDays
                
                    $successCount = @($results | Where-Object { $_.Status -eq "Success" }).Count
                    $failCount = @($results | Where-Object { $_.Status -eq "Failed" }).Count
                
                    Write-GUILog "Delete operation completed: $successCount succeeded, $failCount failed"
                
                    # Export results
                    if ($results.Count -gt 0) {
                        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                        $actionFile = Join-Path $script:LogPath "DeleteActions_$timestamp.csv"
                        $results | Export-Csv -Path $actionFile -NoTypeInformation
                        Write-GUILog "Actions logged to: $actionFile"
                    }
                
                    [System.Windows.Forms.MessageBox]::Show("Delete operation completed.`n`nSucceeded: $successCount`nFailed: $failCount", "Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                
                    # Refresh the grid
                    $btnScan.PerformClick()
                }
                catch {
                    Write-GUILog "Error during delete operation: $_" -Level Error
                    [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
                finally {
                    $btnDelete.Enabled = $true
                    $form.Cursor = [System.Windows.Forms.Cursors]::Default
                }
            }
            else {
                Write-GUILog "Delete operation cancelled - confirmation text not entered" -Level Warning
            }
        }
    })

$btnExport.Add_Click({
        if ($script:ComputerReport.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No data to export. Please scan for stale computers first.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    
        $exportFolder = $txtExportPath.Text.Trim()
    
        # If no folder specified, prompt user to select one
        if (-not $exportFolder) {
            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderDialog.Description = "Select folder to save CSV report"
            $folderDialog.ShowNewFolderButton = $true
        
            if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $exportFolder = $folderDialog.SelectedPath
                $txtExportPath.Text = $exportFolder
            }
            else {
                return
            }
        }
    
        # Auto-generate filename with timestamp
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $fileName = "StaleComputers_Report_$timestamp.csv"
        $exportPath = Join-Path $exportFolder $fileName
    
        try {
            # Ensure directory exists
            if (-not (Test-Path $exportFolder)) {
                New-Item -ItemType Directory -Path $exportFolder -Force | Out-Null
            }
        
            # Export without the sort column
            $script:ComputerReport | Select-Object Name, Enabled, DistinguishedName, LastLogonDate, DaysSinceLogon, PasswordLastSet, WhenCreated, OperatingSystem, Description, MarkedForDeletion, DisableDate, DaysSinceDisabled | 
            Export-Csv -Path $exportPath -NoTypeInformation
        
            Write-GUILog "Report exported to: $exportPath"
            [System.Windows.Forms.MessageBox]::Show("Report exported successfully to:`n$exportPath", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            Write-GUILog "Failed to export report: $_" -Level Error
            [System.Windows.Forms.MessageBox]::Show("Failed to export report: $_", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
#endregion Event Handlers

# Show form
Write-GUILog "Stale AD Computer Management GUI initialized"
[void]$form.ShowDialog()
