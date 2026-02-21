<#
.SYNOPSIS
    Identifies, disables, and deletes stale computer objects in Active Directory.

.DESCRIPTION
    Implements a 3-stage lifecycle for stale AD computer management:
    Stage 1: Identify workstations inactive for X days (default 365)
    Stage 2: Disable workstations and move to staging OU, tag with disable date
    Stage 3: Delete workstations that have been disabled for Y days (default 30)

    This script automatically targets only Windows workstation operating systems
    (Windows 10, Windows 11) and excludes Windows Server systems.

    For Hybrid Azure AD Join environments, deletions will sync to Azure AD/Intune
    via Azure AD Connect.

.PARAMETER InactiveDays
    Number of days of inactivity before a computer is considered stale. Default: 365

.PARAMETER DeleteAfterDays
    Number of days after disabling before a computer is deleted. Default: 30

.PARAMETER StagingOU
    Distinguished Name of the OU where disabled computers are moved.

.PARAMETER ReportOnly
    Generate reports without making any changes.

.PARAMETER DisableStale
    Disable stale computers and move to staging OU.

.PARAMETER DeleteDisabled
    Delete computers that have been disabled for the specified period.

.PARAMETER LogPath
    Path for log files. Default: C:\Temp\StaleADComputerLogs

.PARAMETER ExportPath
    Path for exporting the CSV report. If not specified, exports to LogPath.

.PARAMETER Force
    Skip confirmation prompts when disabling or deleting computers.

.EXAMPLE
    .\Manage-StaleADComputers.ps1 -ReportOnly
    Generate a report of all stale Windows workstations without making changes.

.EXAMPLE
    .\Manage-StaleADComputers.ps1 -ReportOnly -ExportPath "C:\Reports\StaleDevices.csv"
    Export stale workstations to a specific CSV file for review.

.EXAMPLE
    .\Manage-StaleADComputers.ps1 -DisableStale -StagingOU "OU=ToBeDeleted,DC=contoso,DC=com"
    Disable stale workstations and move them to the staging OU.

.EXAMPLE
    .\Manage-StaleADComputers.ps1 -DeleteDisabled
    Delete workstations that have been in the staging OU for the deletion period.

.NOTES
    Author: Alberto de la Torre
    Version: 1.0
    Date: February 2026

    Industry Best Practices Implemented:
    - Multi-stage approach (Identify → Disable → Delete)
    - Configurable thresholds
    - OS-based filtering (Windows 10/11 workstations only)
    - Automatic exclusion of Windows Server systems
    - Comprehensive logging
    - Description tagging with timestamps
    - Staging OU for review before deletion
#>

#Requires -Modules ActiveDirectory

[CmdletBinding(DefaultParameterSetName = 'Report')]
param(
    [Parameter()]
    [int]$InactiveDays = 365,

    [Parameter()]
    [int]$DeleteAfterDays = 30,

    [Parameter()]
    [string]$StagingOU,

    [Parameter(ParameterSetName = 'Report')]
    [switch]$ReportOnly,

    [Parameter(ParameterSetName = 'Disable')]
    [switch]$DisableStale,

    [Parameter(ParameterSetName = 'Delete')]
    [switch]$DeleteDisabled,

    [Parameter()]
    [string]$LogPath = "C:\Temp\StaleADComputerLogs",

    [Parameter()]
    [string]$ExportPath,

    [Parameter()]
    [switch]$IncludeWorkstations,

    [Parameter()]
    [switch]$IncludeServers,

    [Parameter()]
    [switch]$Force
)

# Handle default OS filter if neither is specified
if (-not $IncludeWorkstations -and -not $IncludeServers) {
    $IncludeWorkstations = $true
}

#region Functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        'Info' { Write-Host $logMessage -ForegroundColor Green }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error' { Write-Host $logMessage -ForegroundColor Red }
    }
    
    Add-Content -Path $script:LogFile -Value $logMessage
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
        Write-Log "No OS types selected for search." -Level Warning
        return @()
    }
    
    $filterString = $osFilters -join " -or "
    Write-Log "Searching for computers inactive since $($cutoffDate.ToString('yyyy-MM-dd')) matching OS filter."
    
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
    
    Write-Log "Found $($allComputers.Count) computers matching filter in AD"
    
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
    
    Write-Log "Found $($staleComputers.Count) stale computers"
    return $staleComputers
}

function Get-ComputerReport {
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
            -1
        }
        
        # Check if marked for deletion (from description)
        $markedForDeletion = $false
        $disableDate = $null
        if ($Computer.Description -match "DISABLED:\s*(\d{4}-\d{2}-\d{2})") {
            $markedForDeletion = $true
            $disableDate = [DateTime]::Parse($Matches[1])
        }
        
        [PSCustomObject]@{
            Name              = $Computer.Name
            Enabled           = $Computer.Enabled
            DistinguishedName = $Computer.DistinguishedName
            LastLogonDate     = if ($lastLogon) { $lastLogon.ToString("yyyy-MM-dd") } else { "Never" }
            DaysSinceLogon    = if ($daysSinceLogon -eq -1) { "Never" } else { $daysSinceLogon }
            PasswordLastSet   = if ($pwdLastSet) { $pwdLastSet.ToString("yyyy-MM-dd") } else { "" }
            WhenCreated       = $Computer.whenCreated.ToString("yyyy-MM-dd")
            OperatingSystem   = $Computer.OperatingSystem
            Description       = $Computer.Description
            MarkedForDeletion = $markedForDeletion
            DisableDate       = if ($disableDate) { $disableDate.ToString("yyyy-MM-dd") } else { "" }
            DaysSinceDisabled = if ($disableDate) { [math]::Round(((Get-Date) - $disableDate).TotalDays, 0) } else { $null }
        }
    }
}

function Disable-StaleComputer {
    param(
        [Parameter(ValueFromPipeline)]
        $Computer,
        [string]$TargetOU
    )
    
    process {
        $computerName = $Computer.Name
        $today = Get-Date -Format "yyyy-MM-dd"
        $originalDescription = $Computer.Description
        $newDescription = "DISABLED: $today | Original: $originalDescription | Stale $InactiveDays+ days"
        
        try {
            # Disable the computer account
            Set-ADComputer -Identity $Computer.DistinguishedName -Enabled $false -Description $newDescription
            Write-Log "Disabled computer: $computerName"
            
            # Move to staging OU if specified
            if ($TargetOU) {
                Move-ADObject -Identity $Computer.DistinguishedName -TargetPath $TargetOU
                Write-Log "Moved $computerName to staging OU: $TargetOU"
            }
            
            [PSCustomObject]@{
                ComputerName = $computerName
                Action       = "Disabled"
                Status       = "Success"
                Timestamp    = Get-Date
                NewLocation  = if ($TargetOU) { $TargetOU } else { "Not moved" }
            }
        }
        catch {
            Write-Log "Failed to disable $computerName : $_" -Level Error
            
            [PSCustomObject]@{
                ComputerName = $computerName
                Action       = "Disable"
                Status       = "Failed"
                Timestamp    = Get-Date
                Error        = $_.Exception.Message
            }
        }
    }
}

function Remove-DisabledComputer {
    param(
        [Parameter(ValueFromPipeline)]
        $Computer,
        [int]$MinDaysDisabled
    )
    
    process {
        $computerName = $Computer.Name
        
        # Check if computer has been disabled long enough
        if ($Computer.DaysSinceDisabled -lt $MinDaysDisabled) {
            Write-Log "Skipping $computerName - only disabled for $($Computer.DaysSinceDisabled) days (min: $MinDaysDisabled)" -Level Warning
            return
        }
        
        try {
            # Get fresh AD object for deletion
            $adComputer = Get-ADComputer -Identity $computerName
            
            # Remove the computer object
            Remove-ADComputer -Identity $adComputer -Confirm:$false
            Write-Log "Deleted computer: $computerName (was disabled for $($Computer.DaysSinceDisabled) days)"
            
            [PSCustomObject]@{
                ComputerName = $computerName
                Action       = "Deleted"
                Status       = "Success"
                Timestamp    = Get-Date
                DaysDisabled = $Computer.DaysSinceDisabled
            }
        }
        catch {
            Write-Log "Failed to delete $computerName : $_" -Level Error
            
            [PSCustomObject]@{
                ComputerName = $computerName
                Action       = "Delete"
                Status       = "Failed"
                Timestamp    = Get-Date
                Error        = $_.Exception.Message
            }
        }
    }
}

function Set-DeletionTag {
    <#
    .SYNOPSIS
        Tags an already-disabled computer for deletion by adding the DISABLED timestamp
    #>
    param(
        [Parameter(ValueFromPipeline)]
        $Computer
    )
    
    process {
        $computerName = $Computer.Name
        $today = Get-Date -Format "yyyy-MM-dd"
        
        try {
            $adComputer = Get-ADComputer -Identity $computerName -Properties Description
            $originalDescription = $adComputer.Description
            $newDescription = "DISABLED: $today | Original: $originalDescription | Tagged for deletion (stale $InactiveDays+ days)"
            
            Set-ADComputer -Identity $adComputer.DistinguishedName -Description $newDescription
            Write-Log "Tagged computer for deletion: $computerName"
            
            [PSCustomObject]@{
                ComputerName   = $computerName
                Action         = "Tagged"
                Status         = "Success"
                Timestamp      = Get-Date
                NewDescription = $newDescription
            }
        }
        catch {
            Write-Log "Failed to tag $computerName : $_" -Level Error
            
            [PSCustomObject]@{
                ComputerName = $computerName
                Action       = "Tag"
                Status       = "Failed"
                Timestamp    = Get-Date
                Error        = $_.Exception.Message
            }
        }
    }
}
#endregion Functions

#region Main Script
# Ensure log directory exists
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd"
$script:LogFile = Join-Path $LogPath "StaleADComputers_$timestamp.log"

Write-Log "=========================================="
Write-Log "Stale AD Computer Management Script"
Write-Log "=========================================="
Write-Log "Parameters:"
Write-Log "  Inactive Days Threshold: $InactiveDays"
Write-Log "  Delete After Days: $DeleteAfterDays"
Write-Log "  Staging OU: $(if ($StagingOU) { $StagingOU } else { 'Not specified' })"
Write-Log "  Included Workstations: $IncludeWorkstations"
Write-Log "  Included Servers: $IncludeServers"
Write-Log "  Mode: $(if ($DisableStale) { 'Disable Stale' } elseif ($DeleteDisabled) { 'Delete Disabled' } else { 'Report Only' })"
Write-Log "=========================================="

# Get stale computers
$staleComputers = Get-StaleComputers -DaysInactive $InactiveDays `
    -IncludeWorkstations $IncludeWorkstations `
    -IncludeServers $IncludeServers

if ($staleComputers.Count -eq 0) {
    Write-Log "No stale computers found matching criteria."
    exit 0
}

# Generate detailed report
$report = $staleComputers | Get-ComputerReport

# Export full report
if ($ExportPath) {
    # Use specified export path
    $reportFile = $ExportPath
    # Ensure directory exists
    $exportDir = Split-Path -Path $ExportPath -Parent
    if ($exportDir -and -not (Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }
}
else {
    $reportFile = Join-Path $LogPath "StaleComputers_Report_$timestamp.csv"
}
$report | Export-Csv -Path $reportFile -NoTypeInformation
Write-Log "Full report exported to: $reportFile"
Write-Host ""
Write-Host "=== CSV EXPORT ==="  -ForegroundColor Cyan
Write-Host "Review stale devices at: $reportFile" -ForegroundColor Cyan
Write-Host "==================" -ForegroundColor Cyan
Write-Host ""

# Create summary statistics
$summary = [PSCustomObject]@{
    TotalStaleComputers    = $report.Count
    EnabledStaleComputers  = @($report | Where-Object { $_.Enabled }).Count
    DisabledStaleComputers = @($report | Where-Object { -not $_.Enabled }).Count
    MarkedForDeletion      = @($report | Where-Object { $_.MarkedForDeletion }).Count
    ReadyForDeletion       = @($report | Where-Object { $_.MarkedForDeletion -and $_.DaysSinceDisabled -ge $DeleteAfterDays }).Count
    AverageInactiveDays    = [math]::Round(($report | Where-Object { $_.DaysSinceLogon -ne "Never" } | Measure-Object -Property DaysSinceLogon -Average).Average, 0)
    OldestInactiveDays     = ($report | Where-Object { $_.DaysSinceLogon -ne "Never" } | Measure-Object -Property DaysSinceLogon -Maximum).Maximum
}

Write-Log "=========================================="
Write-Log "Summary:"
Write-Log "  Total Stale Computers: $($summary.TotalStaleComputers)"
Write-Log "  Still Enabled: $($summary.EnabledStaleComputers)"
Write-Log "  Already Disabled: $($summary.DisabledStaleComputers)"
Write-Log "  Marked for Deletion: $($summary.MarkedForDeletion)"
Write-Log "  Ready for Deletion (>$DeleteAfterDays days): $($summary.ReadyForDeletion)"
Write-Log "  Average Inactive Days: $($summary.AverageInactiveDays)"
Write-Log "  Oldest Inactive Days: $($summary.OldestInactiveDays)"
Write-Log "=========================================="

# Default to report mode if no action parameter specified
if ($ReportOnly -or (-not $DisableStale -and -not $DeleteDisabled)) {
    Write-Log "Report mode - no changes made."
    Write-Log "Review the report at: $reportFile"
}
elseif ($DisableStale) {
    if (-not $StagingOU) {
        Write-Log "Warning: No staging OU specified. Computers will be disabled but not moved." -Level Warning
    }
    
    # Get enabled stale computers that haven't been marked yet
    $toDisable = @($report | Where-Object { $_.Enabled -and -not $_.MarkedForDeletion })
    
    # Get disabled stale computers that haven't been tagged (externally disabled)
    $toTag = @($report | Where-Object { -not $_.Enabled -and -not $_.MarkedForDeletion })
    
    $totalToProcess = $toDisable.Count + $toTag.Count
    
    if ($totalToProcess -eq 0) {
        Write-Log "No stale computers found to process. All are already tagged for deletion."
    }
    else {
        if ($toDisable.Count -gt 0) {
            Write-Log "Found $($toDisable.Count) enabled stale computers to disable."
        }
        if ($toTag.Count -gt 0) {
            Write-Log "Found $($toTag.Count) externally-disabled computers to tag for deletion."
        }
        
        # Confirmation prompt unless -Force is specified
        if (-not $Force) {
            Write-Host ""
            if ($toDisable.Count -gt 0) {
                Write-Host "WARNING: You are about to DISABLE $($toDisable.Count) computer account(s) in Active Directory." -ForegroundColor Yellow
            }
            if ($toTag.Count -gt 0) {
                Write-Host "INFO: You will also TAG $($toTag.Count) externally-disabled computer(s) for deletion." -ForegroundColor Cyan
            }
            Write-Host "Review the exported CSV before proceeding: $reportFile" -ForegroundColor Yellow
            Write-Host ""
            $confirmation = Read-Host "Type 'YES' to confirm or any other key to cancel"
            if ($confirmation -ne 'YES') {
                Write-Log "Operation cancelled by user." -Level Warning
                exit 0
            }
        }
    }
    
    $disableResults = @()
    
    # Disable enabled computers
    if ($toDisable.Count -gt 0) {
        Write-Log "Disabling $($toDisable.Count) stale computers..."
        foreach ($computer in $toDisable) {
            $adComputer = Get-ADComputer -Identity $computer.Name -Properties Description
            $result = $adComputer | Disable-StaleComputer -TargetOU $StagingOU
            $disableResults += $result
        }
    }
    
    # Tag externally-disabled computers
    if ($toTag.Count -gt 0) {
        Write-Log "Tagging $($toTag.Count) externally-disabled computers for deletion..."
        foreach ($computer in $toTag) {
            $result = $computer | Set-DeletionTag
            $disableResults += $result
        }
    }
    
    # Export action results
    if ($disableResults.Count -gt 0) {
        $actionFile = Join-Path $LogPath "DisableActions_$timestamp.csv"
        $disableResults | Export-Csv -Path $actionFile -NoTypeInformation
        Write-Log "Actions logged to: $actionFile"
    }
}
elseif ($DeleteDisabled) {
    # Get computers ready for deletion
    $toDelete = $report | Where-Object { 
        $_.MarkedForDeletion -and 
        $_.DaysSinceDisabled -ge $DeleteAfterDays -and
        -not $_.Enabled
    }
    
    if ($toDelete.Count -eq 0) {
        Write-Log "No computers are ready for deletion (must be disabled for $DeleteAfterDays+ days)."
    }
    else {
        Write-Log "Found $($toDelete.Count) computers ready for deletion (disabled for $DeleteAfterDays+ days)."
        
        # Confirmation prompt unless -Force is specified
        if (-not $Force) {
            Write-Host ""
            Write-Host "WARNING: You are about to PERMANENTLY DELETE $($toDelete.Count) computer account(s) from Active Directory." -ForegroundColor Red
            Write-Host "This action CANNOT be undone!" -ForegroundColor Red
            Write-Host "Review the exported CSV before proceeding: $reportFile" -ForegroundColor Yellow
            Write-Host ""
            $confirmation = Read-Host "Type 'DELETE' to confirm or any other key to cancel"
            if ($confirmation -ne 'DELETE') {
                Write-Log "Operation cancelled by user." -Level Warning
                exit 0
            }
        }
        
        Write-Log "Deleting $($toDelete.Count) computers..."
        
        $deleteResults = @()
        foreach ($computer in $toDelete) {
            $result = $computer | Remove-DisabledComputer -MinDaysDisabled $DeleteAfterDays
            if ($result) {
                $deleteResults += $result
            }
        }
        
        # Export action results
        $actionFile = Join-Path $LogPath "DeleteActions_$timestamp.csv"
        $deleteResults | Export-Csv -Path $actionFile -NoTypeInformation
        Write-Log "Delete actions logged to: $actionFile"
    }
}

Write-Log "Script completed."
Write-Log "Log file: $script:LogFile"
#endregion Main Script
