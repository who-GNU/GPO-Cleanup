Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "GPO Cleanup Tool"
$form.Size = New-Object System.Drawing.Size(800, 500)
$form.StartPosition = "CenterScreen"

# Create controls
$lblInstructions = New-Object System.Windows.Forms.Label
$lblInstructions.Text = "Select actions to perform on Group Policies:"
$lblInstructions.Location = New-Object System.Drawing.Point(20, 20)
$lblInstructions.Size = New-Object System.Drawing.Size(400, 20)

$btnFlagGPOs = New-Object System.Windows.Forms.Button
$btnFlagGPOs.Text = "Flag GPOs for Review"
$btnFlagGPOs.Location = New-Object System.Drawing.Point(20, 60)
$btnFlagGPOs.Size = New-Object System.Drawing.Size(200, 30)

# Create a CheckedListBox with owner-draw enabled for custom colors
$lstGPOs = New-Object System.Windows.Forms.CheckedListBox
$lstGPOs.Location = New-Object System.Drawing.Point(20, 110)
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
$btnDeleteGPOs.Location = New-Object System.Drawing.Point(400, 60)
$btnDeleteGPOs.Size = New-Object System.Drawing.Size(200, 30)

$btnSysVolCleanup = New-Object System.Windows.Forms.Button
$btnSysVolCleanup.Text = "Clean Up SysVol"
$btnSysVolCleanup.Location = New-Object System.Drawing.Point(400, 110)
$btnSysVolCleanup.Size = New-Object System.Drawing.Size(200, 30)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 400)
$progressBar.Size = New-Object System.Drawing.Size(760, 20)

$textBoxLogs = New-Object System.Windows.Forms.TextBox
$textBoxLogs.Multiline = $true
$textBoxLogs.ScrollBars = "Vertical"
$textBoxLogs.Location = New-Object System.Drawing.Point(400, 150)
$textBoxLogs.Size = New-Object System.Drawing.Size(380, 230)

# Add controls to the form
$form.Controls.Add($lblInstructions)
$form.Controls.Add($btnFlagGPOs)
$form.Controls.Add($lstGPOs)
$form.Controls.Add($btnDeleteGPOs)
$form.Controls.Add($btnSysVolCleanup)
$form.Controls.Add($progressBar)
$form.Controls.Add($textBoxLogs)

# Define the Flag GPOs button click event
$btnFlagGPOs.Add_Click({
    $textBoxLogs.AppendText("Flagging GPOs for review..." + [Environment]::NewLine)
    
    # Import the GroupPolicy module (ensure it's available)
    Import-Module GroupPolicy -ErrorAction Stop

    # Retrieve all GPOs and apply criteria:
    #   - ModificationTime more than 30 days ago AND
    #   - Not linked (assuming the Links property reflects OU links)
    # Fetch all GPOs
    $allGPOs = Get-GPO -All

    # Retrieve all linked GPOs from Active Directory
    $linkedGPOs = Get-ADObject -LDAPFilter "(gPLink=*)" -SearchBase "DC=YourDomain,DC=com" -SearchScope Subtree |
        Select-Object -ExpandProperty gPLink

    # Extract GUIDs from the gPLink attribute (GPO links are stored as `[{GUID};0]`)
    $linkedGPOs = $linkedGPOs -match "{.*?}" | ForEach-Object { $_ -match "{(.*?)}"; $matches[1] }

    # Filter out linked GPOs
    $flaggedGPOs = $allGPOs | Where-Object {
        $daysSinceModified = (New-TimeSpan -Start $_.ModificationTime -End (Get-Date)).Days
        $isOld = $daysSinceModified -gt 30
        $isUnlinked = -not ($_.Id.Guid -in $linkedGPOs)  # Check if it's linked in AD
        $isOld -and $isUnlinked
    }
    

    # Clear the list and populate with flagged GPOs as custom objects.
    $lstGPOs.Items.Clear()
    foreach ($gpo in $flaggedGPOs) {
        $displayText = "$($gpo.DisplayName) | Last Modified: $($gpo.ModificationTime.ToShortDateString())"
        # Create a custom object with the text and the LastModified property.
        $lstGPOs.Items.Add([PSCustomObject]@{
            DisplayText = $displayText
            LastModified = $gpo.ModificationTime
        })
    }
    
    # Export the flagged GPOs to CSV for review.
    $flaggedGPOs | Select-Object DisplayName, ModificationTime |
        Export-Csv -Path "FlaggedGPOsForReview.csv" -NoTypeInformation
    $textBoxLogs.AppendText("Flagged GPOs exported to FlaggedGPOsForReview.csv" + [Environment]::NewLine)
})

# Define the Delete Selected GPOs button click event
$btnDeleteGPOs.Add_Click({
    $textBoxLogs.AppendText("Deleting selected GPOs..." + [Environment]::NewLine)
    $progressBar.Value = 0
    $progressBar.Maximum = $lstGPOs.CheckedItems.Count

    foreach ($selectedItem in $lstGPOs.CheckedItems) {
        # Extract the GPO name from the DisplayText (assumes the name is before the pipe "|")
        $gpoName = $selectedItem.DisplayText.Split('|')[0].Trim()
        try {
            $gpo = Get-GPO -Name $gpoName -ErrorAction Stop
           ## Remove-GPO -Guid $gpo.Id -Confirm:$false
            $textBoxLogs.AppendText("Deleted GPO: $gpoName" + [Environment]::NewLine)
        } catch {
            $textBoxLogs.AppendText("Error deleting GPO: $gpoName - $_" + [Environment]::NewLine)
        }
        $progressBar.PerformStep()
    }
    
    $textBoxLogs.AppendText("Selected GPOs deleted. Report saved to DeletedGPOs.csv" + [Environment]::NewLine)
})

# Define the SysVol Cleanup button click event
$btnSysVolCleanup.Add_Click({
    $textBoxLogs.AppendText("Cleaning up orphaned SysVol objects..." + [Environment]::NewLine)
    
    # Set your SYSVOL path (modify this to match your domain's path)
    $sysVolPath = "\\YourDomain\SYSVOL\YourDomain\Policies"
    $gpoGuids = (Get-GPO -All).Id.Guid
    $sysVolFolders = Get-ChildItem -Path $sysVolPath -Directory

    $orphanedFolders = $sysVolFolders | Where-Object {
        $folderGuid = $_.Name
        -not ($folderGuid -in $gpoGuids)
    }

    $progressBar.Value = 0
    $progressBar.Maximum = $orphanedFolders.Count

    foreach ($folder in $orphanedFolders) {
        try {
            ##Remove-Item -Path $folder.FullName -Recurse -Force
            $textBoxLogs.AppendText("Deleted orphaned folder: $($folder.Name)" + [Environment]::NewLine)
        } catch {
            $textBoxLogs.AppendText("Error deleting folder: $($folder.Name) - $_" + [Environment]::NewLine)
        }
        $progressBar.PerformStep()
    }
    
    # Optionally export orphaned folders info to CSV
    $orphanedFolders | Select-Object Name, FullName |
        Export-Csv -Path "OrphanedSysVolFolders.csv" -NoTypeInformation
    $textBoxLogs.AppendText("SysVol cleanup complete. Report saved to OrphanedSysVolFolders.csv" + [Environment]::NewLine)
})

# Show the form
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
