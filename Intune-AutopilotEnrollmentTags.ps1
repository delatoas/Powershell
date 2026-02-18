#Requires -RunAsAdministrator
#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Search and manage Autopilot device Group Tags in Microsoft Intune.

.DESCRIPTION
    This script allows administrators to search for Autopilot-enrolled devices by serial number
    or group tag, view device details, and update group tags as needed.

    Features:
    - Search by serial number (exact or partial match)
    - Search by group tag (lists all devices with matching tag)
    - Search for untagged devices (use 'NONE' keyword)
    - View all assigned group tags with device counts
    - View device details (serial, model, manufacturer, group tag, enrollment state)
    - Update group tag with manual entry (single device or bulk update)
    - Export search results and summaries to CSV reports
    - Session logging with timestamped log files

    OUTPUT FILES:
    - Log files:        AutopilotTags_YYYYMMDD_HHmmss.log (session activity log)
    - Device reports:   AutopilotTags_Report_YYYYMMDD_HHmmss.csv (search results)
    - Summary reports:  AutopilotTags_Summary_YYYYMMDD_HHmmss.csv (all group tags)

    PREREQUISITES:
    - Microsoft.Graph.Authentication module installed
    - Appropriate permissions: DeviceManagementServiceConfig.ReadWrite.All
    - Run as Administrator (required for proper Entra authentication)

.NOTES
    Author:         Alberto de la Torre
    Created:        February 5, 2026
    Repository:     https://github.com/smpa-it/Intune-EndUserComputing/tree/main/Scripts

.EXAMPLE
    .\Intune-AutopilotEnrollmentTags.ps1

    Runs the script interactively. Example session:

    ====================================================
    Autopilot Device Group Tag Management
    ====================================================
    1 - Search by Serial Number
    2 - Search by Group Tag
    3 - View All Group Tags
    4 - Exit

    Select option: 1
    Enter serial number (full or partial): PF2K7X

    Found 1 device(s):
    ----------------------------------------------------
    Serial Number : PF2K7XH3
    Manufacturer  : Lenovo
    Model         : ThinkPad X1 Carbon
    Group Tag     : SMPA
    Enrollment    : Enrolled
    ----------------------------------------------------

    Would you like to change the Group Tag? (Y/N): Y
    Enter new Group Tag: SMPAIT

    Group Tag updated successfully!

.EXAMPLE
    # Option 2 - Search for untagged devices using 'NONE' keyword:

    Enter Group Tag to search (or 'NONE' for untagged devices): NONE
    Searching for devices with no Group Tag
    Searching for untagged devices...

    Found 9 device(s):
    ----------------------------------------------------
    Serial Number : ABC12345
    Group Tag     : (None)
    ...

.EXAMPLE
    # Option 3 - View All Group Tags displays:

    ====================================================
              All Assigned Group Tags
    ====================================================
      Total Devices: 2,538
      Unique Tags:   12
    ----------------------------------------------------
    Group Tag                  Device Count
    ----------------------------------------------------
    (No Tag)                             9
    ENZMFG                              83
    SMPA                             1,784
    ----------------------------------------------------

.CHANGELOG
    1.0 - 2026-02-05 - Initial version
    1.1 - 2026-02-05 - Added CSV export for Group Tag search results, session logging
    1.2 - 2026-02-06 - Added Requires RunAsAdministrator, warm-up connection call
    1.3 - 2026-02-06 - Added View All Group Tags with device counts and summary export
    1.4 - 2026-02-06 - Added 'NONE' keyword to search for untagged devices
#>

#Region Functions

# Script-level variables for logging
$script:SessionStartTime = Get-Date -Format "yyyyMMdd_HHmmss"
$script:LogFileName = "AutopilotTags_$($script:SessionStartTime).log"
$script:LogFilePath = Join-Path $PSScriptRoot $script:LogFileName

function Write-Log {
    <#
    .SYNOPSIS
        Writes output to both console and log file.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Action')]
        [string]$Level = 'Info',
        
        [Parameter(Mandatory = $false)]
        [switch]$NoNewLine
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    # Determine console color based on level
    $color = switch ($Level) {
        'Info'    { 'White' }
        'Warning' { 'Yellow' }
        'Error'   { 'Red' }
        'Success' { 'Green' }
        'Action'  { 'Cyan' }
    }

    # Write to console
    if ($NoNewLine) {
        Write-Host $Message -ForegroundColor $color -NoNewline
    }
    else {
        Write-Host $Message -ForegroundColor $color
    }

    # Append to log file
    Add-Content -Path $script:LogFilePath -Value $logEntry -ErrorAction SilentlyContinue
}

function Initialize-SessionLog {
    <#
    .SYNOPSIS
        Creates the session log file with header information.
    #>
    $header = @"
================================================================================
Autopilot Device Group Tag Management - Session Log
================================================================================
Session Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
User:            $env:USERNAME
Computer:        $env:COMPUTERNAME
Log File:        $script:LogFilePath
================================================================================

"@
    Set-Content -Path $script:LogFilePath -Value $header -ErrorAction SilentlyContinue
}

function Close-SessionLog {
    <#
    .SYNOPSIS
        Appends closing information to the session log.
    #>
    $footer = @"

================================================================================
Session Ended: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
================================================================================
"@
    Add-Content -Path $script:LogFilePath -Value $footer -ErrorAction SilentlyContinue
    Write-Host "`nSession log saved to: $script:LogFilePath" -ForegroundColor Cyan
}

function Export-DevicesToCsv {
    <#
    .SYNOPSIS
        Exports device search results to a CSV file.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [array]$Devices,
        
        [Parameter(Mandatory = $false)]
        [string]$SearchType = "GroupTag"
    )

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportFileName = "AutopilotTags_Report_$timestamp.csv"
    $reportFilePath = Join-Path $PSScriptRoot $reportFileName

    try {
        $exportData = $Devices | Select-Object @(
            @{Name='SerialNumber'; Expression={$_.serialNumber}}
            @{Name='Manufacturer'; Expression={$_.manufacturer}}
            @{Name='Model'; Expression={$_.model}}
            @{Name='GroupTag'; Expression={$_.groupTag}}
            @{Name='EnrollmentState'; Expression={$_.enrollmentState}}
            @{Name='DeviceId'; Expression={$_.id}}
            @{Name='LastContactedDateTime'; Expression={$_.lastContactedDateTime}}
            @{Name='ExportedOn'; Expression={Get-Date -Format "yyyy-MM-dd HH:mm:ss"}}
        )

        $exportData | Export-Csv -Path $reportFilePath -NoTypeInformation -Encoding UTF8
        
        Write-Log "Report exported successfully: $reportFilePath" -Level Success
        Write-Log "  Total devices exported: $($Devices.Count)" -Level Info
        return $true
    }
    catch {
        Write-Log "ERROR: Failed to export report - $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Connect-ToGraph {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph with required permissions.
    #>
    try {
        $context = Get-MgContext
        if ($null -eq $context) {
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
            Connect-MgGraph -Scopes "DeviceManagementServiceConfig.ReadWrite.All" -NoWelcome
            Write-Host "Connected successfully." -ForegroundColor Green
        }
        else {
            Write-Host "Already connected to Microsoft Graph as $($context.Account)" -ForegroundColor Green
        }
        
        # Warm-up call to establish token for beta endpoint (prevents second auth prompt)
        Write-Host "Validating connection..." -ForegroundColor Yellow
        $null = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?`$top=1" -Method GET
        Write-Host "Connection validated." -ForegroundColor Green
        
        return $true
    }
    catch {
        Write-Host "ERROR: Failed to connect to Microsoft Graph." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        return $false
    }
}

function Get-AutopilotDeviceBySerial {
    <#
    .SYNOPSIS
        Searches for Autopilot devices by serial number.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$SerialNumber
    )

    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities"
        $allDevices = @()
        
        do {
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            $allDevices += $response.value
            $uri = $response.'@odata.nextLink'
        } while ($uri)

        # Filter by serial number (case-insensitive partial match)
        $devices = $allDevices | Where-Object { $_.serialNumber -like "*$SerialNumber*" }
        
        return $devices
    }
    catch {
        Write-Host "ERROR: Failed to search for devices." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        return $null
    }
}

function Get-AutopilotDeviceByGroupTag {
    <#
    .SYNOPSIS
        Searches for Autopilot devices by group tag.
        Pass an empty string to search for untagged devices.
    #>
    param (
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$GroupTag = ''
    )

    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities"
        $allDevices = @()
        
        do {
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            $allDevices += $response.value
            $uri = $response.'@odata.nextLink'
        } while ($uri)

        # Filter by group tag - empty/null GroupTag searches for untagged devices
        if ([string]::IsNullOrWhiteSpace($GroupTag)) {
            $devices = $allDevices | Where-Object { [string]::IsNullOrWhiteSpace($_.groupTag) }
        }
        else {
            $devices = $allDevices | Where-Object { $_.groupTag -eq $GroupTag }
        }
        
        return $devices
    }
    catch {
        Write-Host "ERROR: Failed to search for devices." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        return $null
    }
}

function Show-DeviceDetails {
    <#
    .SYNOPSIS
        Displays device details in a formatted table.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [array]$Devices
    )

    if ($Devices.Count -eq 0) {
        Write-Host "`nNo devices found matching your search criteria." -ForegroundColor Yellow
        return
    }

    Write-Host "`nFound $($Devices.Count) device(s):" -ForegroundColor Cyan
    Write-Host ("-" * 80)
    
    $Devices | ForEach-Object {
        Write-Host "Serial Number : " -NoNewline; Write-Host $_.serialNumber -ForegroundColor White
        Write-Host "Manufacturer  : " -NoNewline; Write-Host $_.manufacturer -ForegroundColor White
        Write-Host "Model         : " -NoNewline; Write-Host $_.model -ForegroundColor White
        Write-Host "Group Tag     : " -NoNewline
        if ([string]::IsNullOrWhiteSpace($_.groupTag)) {
            Write-Host "(None)" -ForegroundColor DarkGray
        }
        else {
            Write-Host $_.groupTag -ForegroundColor Green
        }
        Write-Host "Enrollment    : " -NoNewline; Write-Host $_.enrollmentState -ForegroundColor White
        Write-Host "Device ID     : " -NoNewline; Write-Host $_.id -ForegroundColor DarkGray
        Write-Host ("-" * 80)
    }
}

function Update-DeviceGroupTag {
    <#
    .SYNOPSIS
        Updates the group tag for a specific Autopilot device.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$DeviceId,
        
        [Parameter(Mandatory = $true)]
        [string]$NewGroupTag,
        
        [Parameter(Mandatory = $false)]
        [string]$SerialNumber = "Unknown"
    )

    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities/$DeviceId/updateDeviceProperties"
        
        $body = @{
            groupTag = $NewGroupTag
        } | ConvertTo-Json

        Invoke-MgGraphRequest -Uri $uri -Method POST -Body $body -ContentType "application/json"
        
        Write-Log "Group Tag updated: Device=$SerialNumber, NewTag=$NewGroupTag" -Level Success
        Write-Host "Note: Changes may take a few minutes to reflect in Intune." -ForegroundColor Yellow
        return $true
    }
    catch {
        Write-Log "FAILED to update Group Tag: Device=$SerialNumber, Error=$($_.Exception.Message)" -Level Error
        return $false
    }
}

function Show-MainMenu {
    <#
    .SYNOPSIS
        Displays the main menu and handles user selection.
    #>
    Write-Host "`n===================================================="
    Write-Host "       Autopilot Device Group Tag Management" -ForegroundColor Cyan
    Write-Host "===================================================="
    Write-Host "  1 - Search by Serial Number"
    Write-Host "  2 - Search by Group Tag"
    Write-Host "  3 - View All Group Tags"
    Write-Host "  4 - Exit"
    Write-Host "====================================================`n"
}

function Get-AllGroupTags {
    <#
    .SYNOPSIS
        Retrieves all unique group tags with device counts.
    #>
    try {
        Write-Host "`nRetrieving all Autopilot devices..." -ForegroundColor Yellow
        $uri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities"
        $allDevices = @()
        
        do {
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            $allDevices += $response.value
            $uri = $response.'@odata.nextLink'
        } while ($uri)

        Write-Log "Retrieved $($allDevices.Count) total Autopilot devices" -Level Info

        # Group by GroupTag and count (use script block to explicitly access property from hashtable)
        $groupTagSummary = $allDevices | Group-Object -Property { $_.groupTag } | 
            Select-Object @(
                @{Name='GroupTag'; Expression={ if ([string]::IsNullOrWhiteSpace($_.Name)) { '(No Tag)' } else { $_.Name } }}
                @{Name='DeviceCount'; Expression={ $_.Count }}
            ) | Sort-Object GroupTag

        $uniqueTagCount = ($groupTagSummary | Where-Object { $_.GroupTag -ne '(No Tag)' }).Count

        Write-Host "`n===================================================="
        Write-Host "          All Assigned Group Tags" -ForegroundColor Cyan
        Write-Host "===================================================="
        Write-Host "  Total Devices: $($allDevices.Count)"
        Write-Host "  Unique Tags:   $uniqueTagCount"
        Write-Host "----------------------------------------------------"
        Write-Host ("{0,-25} {1,12}" -f "Group Tag", "Device Count") -ForegroundColor Yellow
        Write-Host "----------------------------------------------------"
        
        foreach ($tag in $groupTagSummary) {
            if ($tag.GroupTag -eq '(No Tag)') {
                Write-Host ("{0,-25} {1,12}" -f $tag.GroupTag, $tag.DeviceCount) -ForegroundColor DarkGray
            }
            else {
                Write-Host ("{0,-25} {1,12}" -f $tag.GroupTag, $tag.DeviceCount)
            }
        }
        
        Write-Host "----------------------------------------------------"
        Write-Host "  Total: $($allDevices.Count) devices"
        Write-Host "===================================================="

        # Offer to export summary
        Write-Host "`nWould you like to export this summary to CSV?"
        do {
            $exportChoice = Read-Host "(Y)es or (N)o"
        } until ($exportChoice -match '^[YN]$')

        if ($exportChoice -eq 'Y') {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $reportFileName = "AutopilotTags_Summary_$timestamp.csv"
            $reportFilePath = Join-Path $PSScriptRoot $reportFileName
            
            $groupTagSummary | Export-Csv -Path $reportFilePath -NoTypeInformation -Encoding UTF8
            Write-Log "Group Tag summary exported: $reportFilePath" -Level Success
        }

        return $groupTagSummary
    }
    catch {
        Write-Log "ERROR: Failed to retrieve group tags - $($_.Exception.Message)" -Level Error
        return $null
    }
}

function Start-GroupTagUpdate {
    <#
    .SYNOPSIS
        Prompts user to update group tag for selected device(s).
    #>
    param (
        [Parameter(Mandatory = $true)]
        [array]$Devices
    )

    if ($Devices.Count -eq 0) { return }

    # Single device - simple flow
    if ($Devices.Count -eq 1) {
        Write-Host "`nWould you like to change the Group Tag?"
        do {
            $updateChoice = Read-Host "(Y)es or (N)o"
        } until ($updateChoice -match '^[YN]$')

        if ($updateChoice -eq 'N') { return }

        Update-SingleDevice -Device $Devices[0]
        return
    }

    # Multiple devices - show clearer menu
    $currentTag = $Devices[0].groupTag
    Write-Host "`nFound $($Devices.Count) device(s) with Group Tag: " -NoNewline
    Write-Host $currentTag -ForegroundColor Green
    Write-Host "`nWhat would you like to do?" -ForegroundColor Yellow
    Write-Host "  1 - Update a single device"
    Write-Host "  2 - Update ALL $($Devices.Count) devices (bulk)"
    Write-Host "  0 - Return to main menu"

    do {
        $actionChoice = Read-Host "`nSelect option"
    } until ($actionChoice -match '^[0-2]$')

    switch ($actionChoice) {
        '0' { return }
        '1' {
            # Single device selection
            Write-Host "`nSelect device to update:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $Devices.Count; $i++) {
                Write-Host "  $($i + 1) - $($Devices[$i].serialNumber) (Current Tag: $($Devices[$i].groupTag))"
            }
            Write-Host "  0 - Cancel"

            $selectedDevice = $null
            do {
                $selection = Read-Host "`nEnter selection"
                $selectionNum = $selection -as [int]
                if ($selectionNum -eq 0) { return }
                if ($selectionNum -ge 1 -and $selectionNum -le $Devices.Count) {
                    $selectedDevice = $Devices[$selectionNum - 1]
                }
                else {
                    Write-Host "Invalid selection. Please enter a number between 0 and $($Devices.Count)." -ForegroundColor Red
                }
            } until ($null -ne $selectedDevice)

            Update-SingleDevice -Device $selectedDevice
        }
        '2' {
            # Bulk update all devices
            Update-BulkDevices -Devices $Devices
        }
    }
}

function Update-SingleDevice {
    <#
    .SYNOPSIS
        Updates group tag for a single device with confirmation.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [object]$Device
    )

    Write-Host "`nSelected device: $($Device.serialNumber)" -ForegroundColor Cyan
    Write-Host "Current Group Tag: " -NoNewline
    if ([string]::IsNullOrWhiteSpace($Device.groupTag)) {
        Write-Host "(None)" -ForegroundColor DarkGray
    }
    else {
        Write-Host $Device.groupTag -ForegroundColor Green
    }

    # Get new group tag
    do {
        $newGroupTag = Read-Host "`nEnter new Group Tag"
        if ([string]::IsNullOrWhiteSpace($newGroupTag)) {
            Write-Host "Group Tag cannot be empty." -ForegroundColor Red
        }
    } until (-not [string]::IsNullOrWhiteSpace($newGroupTag))

    $newGroupTag = $newGroupTag.ToUpper().Trim()

    # Confirm change
    Write-Host "`nConfirm change:" -ForegroundColor Yellow
    Write-Host "  Device:        $($Device.serialNumber)"
    Write-Host "  Current Tag:   $($Device.groupTag)"
    Write-Host "  New Tag:       $newGroupTag"

    do {
        $confirm = Read-Host "`nProceed with update? (Y/N)"
    } until ($confirm -match '^[YN]$')

    if ($confirm -eq 'Y') {
        Write-Log "User confirmed single device update: $($Device.serialNumber) from '$($Device.groupTag)' to '$newGroupTag'" -Level Action
        Update-DeviceGroupTag -DeviceId $Device.id -NewGroupTag $newGroupTag -SerialNumber $Device.serialNumber
    }
    else {
        Write-Log "User cancelled update for device: $($Device.serialNumber)" -Level Warning
        Write-Host "`nUpdate cancelled." -ForegroundColor Yellow
    }
}

function Update-BulkDevices {
    <#
    .SYNOPSIS
        Updates group tag for multiple devices with strong confirmation.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [array]$Devices
    )

    $currentTag = $Devices[0].groupTag

    # Get new group tag
    do {
        $newGroupTag = Read-Host "`nEnter new Group Tag for ALL $($Devices.Count) devices"
        if ([string]::IsNullOrWhiteSpace($newGroupTag)) {
            Write-Host "Group Tag cannot be empty." -ForegroundColor Red
        }
    } until (-not [string]::IsNullOrWhiteSpace($newGroupTag))

    $newGroupTag = $newGroupTag.ToUpper().Trim()

    # Strong warning and confirmation
    Write-Host "`n" + ("=" * 60) -ForegroundColor Red
    Write-Host "                    *** WARNING ***" -ForegroundColor Red
    Write-Host ("=" * 60) -ForegroundColor Red
    Write-Host "`nYou are about to update $($Devices.Count) devices." -ForegroundColor Yellow
    Write-Host "  Current Tag:   $currentTag"
    Write-Host "  New Tag:       $newGroupTag"
    Write-Host "`nDevices to be updated:" -ForegroundColor Yellow
    $Devices | ForEach-Object { Write-Host "  - $($_.serialNumber)" }
    Write-Host "`n" + ("=" * 60) -ForegroundColor Red

    Write-Host "`nType 'CONFIRM' to proceed with bulk update: " -NoNewline -ForegroundColor Yellow
    $confirmation = Read-Host

    if ($confirmation -ceq 'CONFIRM') {
        Write-Log "User confirmed BULK update: $($Devices.Count) devices from '$currentTag' to '$newGroupTag'" -Level Action
        Write-Host "`nUpdating $($Devices.Count) devices..." -ForegroundColor Yellow
        $successCount = 0
        $failCount = 0
        $failedDevices = @()

        foreach ($device in $Devices) {
            Write-Host "  Updating $($device.serialNumber)... " -NoNewline
            try {
                $uri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities/$($device.id)/updateDeviceProperties"
                $body = @{ groupTag = $newGroupTag } | ConvertTo-Json
                Invoke-MgGraphRequest -Uri $uri -Method POST -Body $body -ContentType "application/json"
                Write-Host "OK" -ForegroundColor Green
                $successCount++
            }
            catch {
                Write-Host "FAILED" -ForegroundColor Red
                $failCount++
                $failedDevices += $device.serialNumber
            }
        }

        Write-Log "Bulk update complete: $successCount successful, $failCount failed" -Level Info
        if ($failCount -gt 0) {
            Write-Log "Failed devices: $($failedDevices -join ', ')" -Level Error
        }
        
        Write-Host "`nBulk update complete:" -ForegroundColor Cyan
        Write-Host "  Successful: $successCount" -ForegroundColor Green
        if ($failCount -gt 0) {
            Write-Host "  Failed:     $failCount" -ForegroundColor Red
        }
        Write-Host "Note: Changes may take a few minutes to reflect in Intune." -ForegroundColor Yellow
    }
    else {
        Write-Log "User cancelled bulk update for $($Devices.Count) devices" -Level Warning
        Write-Host "`nBulk update cancelled. Confirmation text did not match." -ForegroundColor Yellow
    }
}

#EndRegion Functions

#Region Main Script

# Initialize session log
Initialize-SessionLog
Write-Log "Session started" -Level Info

# Verify Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Log "ERROR: Microsoft.Graph.Authentication module is not installed." -Level Error
    Write-Log "Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser" -Level Warning
    exit 1
}

# Connect to Graph
if (-not (Connect-ToGraph)) {
    exit 1
}
Write-Log "Connected to Microsoft Graph" -Level Success

# Main loop
do {
    Show-MainMenu
    
    do {
        $menuChoice = Read-Host "Select option"
    } until ($menuChoice -match '^[1-4]$')

    switch ($menuChoice) {
        '1' {
            # Search by Serial Number
            $serialNumber = Read-Host "`nEnter serial number (full or partial)"
            if (-not [string]::IsNullOrWhiteSpace($serialNumber)) {
                Write-Log "Searching for serial number: $serialNumber" -Level Action
                $devices = Get-AutopilotDeviceBySerial -SerialNumber $serialNumber
                if ($null -ne $devices) {
                    Write-Log "Found $($devices.Count) device(s) matching serial: $serialNumber" -Level Info
                    Show-DeviceDetails -Devices $devices
                    Start-GroupTagUpdate -Devices $devices
                }
            }
            else {
                Write-Log "Serial number cannot be empty." -Level Error
            }
        }
        '2' {
            # Search by Group Tag
            $groupTag = Read-Host "`nEnter Group Tag to search (or 'NONE' for untagged devices)"
            if (-not [string]::IsNullOrWhiteSpace($groupTag)) {
                $groupTag = $groupTag.ToUpper().Trim()
                
                # Check if searching for untagged devices
                if ($groupTag -eq 'NONE') {
                    Write-Log "Searching for devices with no Group Tag" -Level Action
                    Write-Host "Searching for untagged devices..." -ForegroundColor Yellow
                    $devices = Get-AutopilotDeviceByGroupTag -GroupTag ''
                    $displayTag = '(No Tag)'
                }
                else {
                    Write-Log "Searching for Group Tag: $groupTag" -Level Action
                    Write-Host "Searching..." -ForegroundColor Yellow
                    $devices = Get-AutopilotDeviceByGroupTag -GroupTag $groupTag
                    $displayTag = $groupTag
                }
                
                if ($null -ne $devices) {
                    Write-Log "Found $($devices.Count) device(s) with Group Tag: $displayTag" -Level Info
                    Show-DeviceDetails -Devices $devices
                    
                    # Offer export option for Group Tag searches
                    if ($devices.Count -gt 0) {
                        Write-Host "`nWould you like to export these results to CSV?"
                        do {
                            $exportChoice = Read-Host "(Y)es or (N)o"
                        } until ($exportChoice -match '^[YN]$')
                        
                        if ($exportChoice -eq 'Y') {
                            Export-DevicesToCsv -Devices $devices -SearchType "GroupTag"
                        }
                    }
                    
                    Start-GroupTagUpdate -Devices $devices
                }
            }
            else {
                Write-Log "Group Tag cannot be empty. Use 'NONE' to search for untagged devices." -Level Error
            }
        }
        '3' {
            # View All Group Tags
            Write-Log "User requested view all group tags" -Level Action
            $null = Get-AllGroupTags
        }
        '4' {
            Write-Log "User requested disconnect" -Level Info
            Write-Host "`nDisconnecting from Microsoft Graph..." -ForegroundColor Yellow
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            Write-Log "Disconnected from Microsoft Graph" -Level Info
            Close-SessionLog
            Write-Host "Goodbye!" -ForegroundColor Green
        }
    }

} until ($menuChoice -eq '4')

#EndRegion Main Script
