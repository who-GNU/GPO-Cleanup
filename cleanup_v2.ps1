# Script to find unlinked GPOs and those that haven't been modified in X days
param (
    [Parameter(Mandatory=$true)]
    [int]$DaysThreshold
)

# Import required modules
Import-Module GroupPolicy
Import-Module ActiveDirectory

# Get the current date to compare with
$CurrentDate = Get-Date

# Function to check if a GPO is linked anywhere in the domain
function Test-GPOIsLinked {
    param (
        [Parameter(Mandatory=$true)]
        [string]$GpoId
    )
    
    # Format the GPO ID in the same format as it appears in LinkedGroupPolicyObjects
    $FormattedGpoId = "[{0}]" -f $GpoId
    
    # Check domain root
    $DomainRoot = Get-ADDomain
    if ($DomainRoot.LinkedGroupPolicyObjects -match $FormattedGpoId) {
        return $true
    }
    
    # Check all OUs
    $AllOUs = Get-ADOrganizationalUnit -Filter * -Properties LinkedGroupPolicyObjects
    foreach ($OU in $AllOUs) {
        if ($OU.LinkedGroupPolicyObjects -match $FormattedGpoId) {
            return $true
        }
    }
    
    return $false
}

# Get all GPOs in the domain
Write-Host "Retrieving all GPOs in the domain..." -ForegroundColor Cyan
$AllGPOs = Get-GPO -All

# Initialize results arrays
$UnlinkedGPOs = @()
$OldGPOs = @()
$TotalCount = $AllGPOs.Count
$CurrentCount = 0

Write-Host "Processing $TotalCount GPOs..." -ForegroundColor Cyan
foreach ($GPO in $AllGPOs) {
    $CurrentCount++
    Write-Progress -Activity "Checking GPO links" -Status "Processing $($GPO.DisplayName)" -PercentComplete (($CurrentCount / $TotalCount) * 100)
    
    $IsLinked = Test-GPOIsLinked -GpoId $GPO.Id.Guid
    $ModificationTime = $GPO.ModificationTime
    $DaysSinceModified = ($CurrentDate - $ModificationTime).Days
    
    # Check if GPO is unlinked
    if (-not $IsLinked) {
        $UnlinkedGPOs += [PSCustomObject]@{
            Name = $GPO.DisplayName
            ID = $GPO.Id.Guid
            CreationTime = $GPO.CreationTime
            ModificationTime = $ModificationTime
            DaysSinceModified = $DaysSinceModified
        }
    }
    
    # Check if GPO hasn't been modified within threshold
    if ($DaysSinceModified -gt $DaysThreshold) {
        $OldGPOs += [PSCustomObject]@{
            Name = $GPO.DisplayName
            ID = $GPO.Id.Guid
            IsLinked = $IsLinked
            CreationTime = $GPO.CreationTime
            ModificationTime = $ModificationTime
            DaysSinceModified = $DaysSinceModified
        }
    }
}

Write-Progress -Activity "Checking GPO links" -Completed

# Display results
Write-Host "`n=========================" -ForegroundColor Yellow
Write-Host "UNLINKED GPOs:" -ForegroundColor Yellow
Write-Host "=========================" -ForegroundColor Yellow
$UnlinkedGPOs | Format-Table -AutoSize

Write-Host "`n=========================" -ForegroundColor Yellow
Write-Host "GPOs not modified in the last $DaysThreshold days:" -ForegroundColor Yellow
Write-Host "=========================" -ForegroundColor Yellow
$OldGPOs | Format-Table -AutoSize

# Create a combined report of unlinked GPOs that are also old
$UnlinkedAndOld = $UnlinkedGPOs | Where-Object { $_.DaysSinceModified -gt $DaysThreshold }

Write-Host "`n=========================" -ForegroundColor Green
Write-Host "UNLINKED GPOs not modified in the last $DaysThreshold days:" -ForegroundColor Green
Write-Host "=========================" -ForegroundColor Green
$UnlinkedAndOld | Format-Table -AutoSize

# Summary statistics
Write-Host "`n=========================" -ForegroundColor Cyan
Write-Host "SUMMARY:" -ForegroundColor Cyan
Write-Host "=========================" -ForegroundColor Cyan
Write-Host "Total GPOs: $($AllGPOs.Count)"
Write-Host "Unlinked GPOs: $($UnlinkedGPOs.Count)"
Write-Host "GPOs not modified in the last $DaysThreshold days: $($OldGPOs.Count)"
Write-Host "Unlinked GPOs not modified in the last $DaysThreshold days: $($UnlinkedAndOld.Count)"

# Export results to CSV
$ExportPath = Join-Path -Path $PWD -ChildPath "GPO_Audit_Report_$(Get-Date -Format 'yyyyMMdd').csv"
$UnlinkedAndOld | Export-Csv -Path $ExportPath -NoTypeInformation

Write-Host "`nReport exported to: $ExportPath" -ForegroundColor Cyan
