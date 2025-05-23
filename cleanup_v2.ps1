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

# Function to collect all linked GPO GUIDs in the domain
function Get-LinkedGPOGuids {
    $LinkedGPOGuids = @{}
    
    Write-Host "Collecting all linked GPOs..." -ForegroundColor Cyan
    
    # Check domain root
    try {
        $DomainRoot = Get-ADDomain
        if ($DomainRoot.LinkedGroupPolicyObjects) {
            foreach ($Link in $DomainRoot.LinkedGroupPolicyObjects) {
                # Extract GUID from the link string
                if ($Link -match '\{([0-9a-fA-F-]+)\}') {
                    $Guid = $Matches[1]
                    $LinkedGPOGuids[$Guid] = $true
                }
            }
        }
    } catch {
        Write-Warning "Error checking domain root links: $_"
    }
    
    # Check all OUs
    try {
        $AllOUs = Get-ADOrganizationalUnit -Filter * -Properties LinkedGroupPolicyObjects
        foreach ($OU in $AllOUs) {
            if ($OU.LinkedGroupPolicyObjects) {
                foreach ($Link in $OU.LinkedGroupPolicyObjects) {
                    # Extract GUID from the link string
                    if ($Link -match '\{([0-9a-fA-F-]+)\}') {
                        $Guid = $Matches[1]
                        $LinkedGPOGuids[$Guid] = $true
                    }
                }
            }
        }
    } catch {
        Write-Warning "Error checking OU links: $_"
    }
    
    return $LinkedGPOGuids
}

# Collect all linked GPO GUIDs
$LinkedGPOGuids = Get-LinkedGPOGuids

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
    
    # Check if GPO is linked by looking up its GUID in the hashtable
    $IsLinked = $LinkedGPOGuids.ContainsKey($GPO.Id.Guid.ToString())
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
