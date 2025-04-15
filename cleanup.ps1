Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "GPO Cleanup Tool"
$form.Size = New-Object System.Drawing.Size(800, 500)
$form.StartPosition = "CenterScreen"

# Create a variable for days threshold (default 90)
$global:daysThreshold = 90

# Create controls
$lblInstructions = New-Object System.Windows.Forms.Label
$lblInstructions.Text = "Select actions to perform on Group Policies:"
$lblInstructions.Location = New-Object System.Drawing.Point(20, 20)
$lblInstructions.Size = New-Object System.Drawing.Size(400, 20)

# Add threshold configuration
$lblThreshold = New-Object System.Windows.Forms.Label
$lblThreshold.Text = "Days Threshold:"
$lblThreshold.Location = New-Object System.Drawing.Point(20, 50)
$lblThreshold.Size = New-Object System.Drawing.Size(100, 20)

$txtThreshold = New-Object System.Windows.Forms.TextBox
$txtThreshold.Text = $global:daysThreshold
$txtThreshold.Location = New-Object System.Drawing.Point(125, 50)
$txtThreshold.Size = New-Object System.Drawing.Size(50, 20)

$btnFlagGPOs = New-Object System.Windows.Forms.Button
$btnFlagGPOs.Text = "Flag GPOs for Review"
$btnFlagGPOs.Location = New-Object System.Drawing.Point(20, 80)
$btnFlagGPOs.Size = New-Object System.Drawing.Size(200, 30)

# Create a CheckedListBox with owner-draw enabled for custom colors
$lstGPOs = New-Object System.Windows.Forms.CheckedListBox
$lstGPOs.Location = New-Object System.Drawing.Point(20, 130)
$lstGPOs.Size = New-Object System.Drawing.Size(350, 250)
$lstGPOs.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
# Set DisplayMember so that each item shows its DisplayText property
$lstGPOs.DisplayMember = "DisplayText"

# Owner-draw event handler: customize item background based on modification age.
$lstGPOs.Add_DrawItem({
    param($sender, $e)

    # Ensure a valid index
    if ($e.Index -lt 0) { return }

    # Retrieve the current item.
    $item = $sender.Items[$e.Index]

    # Default text if for some reason we don't have a custom object.
    $text = $item.ToString()

    # Determine days since modification if the object has the LastModified property.
    $daysSinceModified = 0
    if ($item -is [PSCustomObject] -and $item.PSObject.Properties.Name -contains "LastModified") {
        $lastModified = $item.LastModified
        $daysSinceModified = (New-TimeSpan -Start $lastModified -End (Get-Date)).Days
    }

    # Choose background color based on age
    if ($e.State -band [System.Windows.Forms.DrawItemState]::Selected) {
        $bgColor = [System.Drawing.SystemColors]::Highlight
        $fgColor = [System.Drawing.SystemColors]::HighlightText
    }
    else {
        if ($daysSinceModified -ge 90) {
            $bgColor = [System.Drawing.Color]::Red
        }
        elseif ($daysSinceModified -ge 60) {
            $bgColor = [System.Drawing.Color]::Orange
        }
        elseif ($daysSinceModified -ge 30) {
            $bgColor = [System.Drawing.Color]::Yellow
        }
        else {
            $bgColor = [System.Drawing.Color]::White
        }
        $fgColor = [System.Drawing.Color]::Black
    }

    # Fill the background
    $e.Graphics.FillRectangle((New-Object System.Drawing.SolidBrush($bgColor)), $e.Bounds)
    # Draw the text
    $e.Graphics.DrawString($text, $e.Font, (New-Object System.Drawing.SolidBrush($fgColor)), $e.Bounds.Location)
    # Draw focus rectangle if needed
    $e.DrawFocusRectangle()
})

$btnDeleteGPOs = New-Object System.Windows.Forms.Button
$btnDeleteGPOs.Text = "Delete Selected GPOs"
$btnDeleteGPOs.Location = New-Object System.Drawing.Point(400, 80)
$btnDeleteGPOs.Size = New-Object System.Drawing.Size(200, 30)

$btnSysVolCleanup = New-Object System.Windows.Forms.Button
$btnSysVolCleanup.Text = "Clean Up SysVol"
$btnSysVolCleanup.Location = New-Object System.Drawing.Point(400, 130)
$btnSysVolCleanup.Size = New-Object System.Drawing.Size(200, 30)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 400)
$progressBar.Size = New-Object System.Drawing.Size(760, 20)
$progressBar.Step = 1  # Set the step value

$textBoxLogs = New-Object System.Windows.Forms.TextBox
$textBoxLogs.Multiline = $true
$textBoxLogs.ScrollBars = "Vertical"
$textBoxLogs.Location = New-Object System.Drawing.Point(400, 170)
$textBoxLogs.Size = New-Object System.Drawing.Size(380, 210)

# Add controls to the form
$form.Controls.Add($lblInstructions)
$form.Controls.Add($lblThreshold)
$form.Controls.Add($txtThreshold)
$form.Controls.Add($btnFlagGPOs)
$form.Controls.Add($lstGPOs)
$form.Controls.Add($btnDeleteGPOs)
$form.Controls.Add($btnSysVolCleanup)
$form.Controls.Add($progressBar)
$form.Controls.Add($textBoxLogs)

# Function to get domain information
function Get-DomainInfo {
    $domain = Get-ADDomain
    return @{
        DomainDN = $domain.DistinguishedName
        NetBIOSName = $domain.NetBIOSName
        DNSRoot = $domain.DNSRoot
    }
}

# Define the Flag GPOs button click event
$btnFlagGPOs.Add_Click({
    $textBoxLogs.AppendText("Flagging GPOs for review..." + [Environment]::NewLine)
    
    # Get threshold from textbox
    try {
        $global:daysThreshold = [int]$txtThreshold.Text
    } catch {
        $textBoxLogs.AppendText("Invalid threshold value. Using default of 90 days." + [Environment]::NewLine)
        $global:daysThreshold = 90
        $txtThreshold.Text = "90"
    }
    
    # Import the GroupPolicy module (ensure it's available)
    try {
        Import-Module GroupPolicy -ErrorAction Stop
    } catch {
        $textBoxLogs.AppendText("Error: Failed to import GroupPolicy module. Is it installed?" + [Environment]::NewLine)
        return
    }

    # Get domain info
    try {
        $domainInfo = Get-DomainInfo
        $domainDN = $domainInfo.DomainDN
        $textBoxLogs.AppendText("Working with domain: $($domainInfo.DNSRoot)" + [Environment]::NewLine)
    } catch {
        $textBoxLogs.AppendText("Error: Failed to get domain information. $_" + [Environment]::NewLine)
        return
    }

    # Retrieve all GPOs and apply criteria:
    # Fetch all GPOs
    try {
        $allGPOs = Get-GPO -All
    } catch {
        $textBoxLogs.AppendText("Error retrieving GPOs: $_" + [Environment]::NewLine)
        return
    }
    
    # Initialize progress bar
    $progressBar.Value = 0
    $progressBar.Maximum = 2  # Two major steps
    
    # Step 1: Get linked GPOs
    $textBoxLogs.AppendText("Finding linked GPOs..." + [Environment]::NewLine)
    $linkedGPOGuids = @()
    
    try {
        # Retrieve all linked GPOs from Active Directory
        $adObjects = Get-ADObject -LDAPFilter "(gPLink=*)" -SearchBase $domainDN -SearchScope Subtree -Properties gPLink
        
        # Process each object with GPO links
        foreach ($obj in $adObjects) {
            if ($obj.gPLink) {
                # Extract GUIDs from the gPLink attribute
                # gPLink format: [LDAP://cn={guid},cn=policies,cn=system,DC=domain,DC=com;0]
                $matches = [regex]::Matches($obj.gPLink, "cn=\{(.*?)\}")
                foreach ($match in $matches) {
                    if ($match.Groups.Count -ge 2) {
                        $linkedGPOGuids += $match.Groups[1].Value
                    }
                }
            }
        }
        
        $textBoxLogs.AppendText("Found $($linkedGPOGuids.Count) linked GPOs." + [Environment]::NewLine)
    } catch {
        $textBoxLogs.AppendText("Error finding linked GPOs: $_" + [Environment]::NewLine)
    }
    
    $progressBar.PerformStep()
    
    # Step 2: Filter GPOs based on modification time and link status
    $textBoxLogs.AppendText("Filtering GPOs based on age ($global:daysThreshold days) and link status..." + [Environment]::NewLine)
    
    $flaggedGPOs = $allGPOs | Where-Object {
        $daysSinceModified = (New-TimeSpan -Start $_.ModificationTime -End (Get-Date)).Days
        $isOld = $daysSinceModified -gt $global:daysThreshold
        $isUnlinked = -not ($_.Id.Guid -in $linkedGPOGuids)
        $isOld -and $isUnlinked
    }
    
    $progressBar.PerformStep()
    
    # Clear the list and populate with flagged GPOs as custom objects
    $lstGPOs.Items.Clear()
    foreach ($gpo in $flaggedGPOs) {
        $displayText = "$($gpo.DisplayName) | Last Modified: $($gpo.ModificationTime.ToShortDateString()) | ID: $($gpo.Id.Guid)"
        # Create a custom object with the text and the LastModified property
        $lstGPOs.Items.Add([PSCustomObject]@{
            DisplayText = $displayText
            LastModified = $gpo.ModificationTime
            Id = $gpo.Id
            Name = $gpo.DisplayName
            PSPath = $gpo.PSPath
        })
    }
    
    # Export the flagged GPOs to CSV for review
    try {
        $flaggedGPOs | Select-Object DisplayName, Id, CreationTime, ModificationTime, Description |
            Export-Csv -Path "FlaggedGPOsForReview.csv" -NoTypeInformation
        $textBoxLogs.AppendText("Flagged $($flaggedGPOs.Count) GPOs exported to FlaggedGPOsForReview.csv" + [Environment]::NewLine)
    } catch {
        $textBoxLogs.AppendText("Error exporting to CSV: $_" + [Environment]::NewLine)
    }
})

# Define the Delete Selected GPOs button click event
$btnDeleteGPOs.Add_Click({
    $textBoxLogs.AppendText("Deleting selected GPOs..." + [Environment]::NewLine)
    $progressBar.Value = 0
    $progressBar.Maximum = $lstGPOs.CheckedItems.Count
    
    if ($lstGPOs.CheckedItems.Count -eq 0) {
        $textBoxLogs.AppendText("No GPOs selected for deletion." + [Environment]::NewLine)
        return
    }
    
    # Create a list to store deletion results
    $deletionResults = @()
    
    foreach ($selectedItem in $lstGPOs.CheckedItems) {
        # Get GPO ID from the item
        $gpoGuid = $selectedItem.Id.Guid
        $gpoName = $selectedItem.Name
        
        try {
            $textBoxLogs.AppendText("Deleting GPO: $gpoName ($gpoGuid)" + [Environment]::NewLine)
            # Uncomment the following line to actually delete
            # Remove-GPO -Guid $gpoGuid -Confirm:$false
            
            # For now, just log what would be deleted
            $textBoxLogs.AppendText("Would delete GPO: $gpoName (Simulation)" + [Environment]::NewLine)
            
            # Add to results
            $deletionResults += [PSCustomObject]@{
                Name = $gpoName
                GUID = $gpoGuid
                Status = "Deleted (Simulated)"
                Timestamp = Get-Date
            }
        } catch {
            $textBoxLogs.AppendText("Error deleting GPO: $gpoName - $_" + [Environment]::NewLine)
            $deletionResults += [PSCustomObject]@{
                Name = $gpoName
                GUID = $gpoGuid
                Status = "Error: $_"
                Timestamp = Get-Date
            }
        }
        $progressBar.PerformStep()
    }
    
    # Export results to CSV
    try {
        $deletionResults | Export-Csv -Path "DeletedGPOs.csv" -NoTypeInformation
        $textBoxLogs.AppendText("Selected GPOs processed. Report saved to DeletedGPOs.csv" + [Environment]::NewLine)
    } catch {
        $textBoxLogs.AppendText("Error saving deletion report: $_" + [Environment]::NewLine)
    }
})

# Define the SysVol Cleanup button click event
$btnSysVolCleanup.Add_Click({
    $textBoxLogs.AppendText("Cleaning up orphaned SysVol objects..." + [Environment]::NewLine)
    
    # Get domain info for SYSVOL path
    try {
        $domainInfo = Get-DomainInfo
        $domainNetBIOS = $domainInfo.NetBIOSName
        $domainDNSRoot = $domainInfo.DNSRoot
        
        # Set SYSVOL path based on domain info
        $sysVolPath = "\\$domainDNSRoot\SYSVOL\$domainDNSRoot\Policies"
        $textBoxLogs.AppendText("Using SYSVOL path: $sysVolPath" + [Environment]::NewLine)
    } catch {
        $textBoxLogs.AppendText("Error determining SYSVOL path: $_" + [Environment]::NewLine)
        return
    }
    
    # Get all GPO GUIDs from AD
    try {
        $gpoGuids = (Get-GPO -All).Id.Guid
        $textBoxLogs.AppendText("Found $($gpoGuids.Count) GPOs in Active Directory." + [Environment]::NewLine)
    } catch {
        $textBoxLogs.AppendText("Error retrieving GPO GUIDs: $_" + [Environment]::NewLine)
        return
    }
    
    # Get folders in SYSVOL\Policies
    try {
        $sysVolFolders = Get-ChildItem -Path $sysVolPath -Directory -ErrorAction Stop
        $textBoxLogs.AppendText("Found $($sysVolFolders.Count) folders in SYSVOL." + [Environment]::NewLine)
    } catch {
        $textBoxLogs.AppendText("Error accessing SYSVOL path: $_" + [Environment]::NewLine)
        return
    }
    
    # Find orphaned folders (those in SYSVOL but not in AD)
    $orphanedFolders = $sysVolFolders | Where-Object {
        # Remove curly braces if present in folder name
        $folderName = $_.Name -replace '[{}]', ''
        -not ($folderName -in $gpoGuids)
    }
    
    $textBoxLogs.AppendText("Found $($orphanedFolders.Count) orphaned folders in SYSVOL." + [Environment]::NewLine)
    
    $progressBar.Value = 0
    $progressBar.Maximum = $orphanedFolders.Count
    
    if ($orphanedFolders.Count -eq 0) {
        $textBoxLogs.AppendText("No orphaned folders found." + [Environment]::NewLine)
        return
    }
    
    # Create a list to store cleanup results
    $cleanupResults = @()
    
    foreach ($folder in $orphanedFolders) {
        try {
            $textBoxLogs.AppendText("Processing orphaned folder: $($folder.Name)" + [Environment]::NewLine)
            # Uncomment the following line to actually delete
            # Remove-Item -Path $folder.FullName -Recurse -Force
            
            # For now, just log what would be deleted
            $textBoxLogs.AppendText("Would delete folder: $($folder.FullName) (Simulation)" + [Environment]::NewLine)
            
            # Add to results
            $cleanupResults += [PSCustomObject]@{
                Name = $folder.Name
                Path = $folder.FullName
                Status = "Deleted (Simulated)"
                Timestamp = Get-Date
            }
        } catch {
            $textBoxLogs.AppendText("Error processing folder: $($folder.Name) - $_" + [Environment]::NewLine)
            $cleanupResults += [PSCustomObject]@{
                Name = $folder.Name
                Path = $folder.FullName
                Status = "Error: $_"
                Timestamp = Get-Date
            }
        }
        $progressBar.PerformStep()
    }
    
    # Export cleanup results to CSV
    try {
        $cleanupResults | Export-Csv -Path "OrphanedSysVolFolders.csv" -NoTypeInformation
        $textBoxLogs.AppendText("SysVol cleanup complete. Report saved to OrphanedSysVolFolders.csv" + [Environment]::NewLine)
    } catch {
        $textBoxLogs.AppendText("Error saving cleanup report: $_" + [Environment]::NewLine)
    }
})

# Show the form
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
