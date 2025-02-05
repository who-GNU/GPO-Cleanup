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

$lstGPOs = New-Object System.Windows.Forms.CheckedListBox
$lstGPOs.Location = New-Object System.Drawing.Point(20, 110)
$lstGPOs.Size = New-Object System.Drawing.Size(350, 250)

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

# Define actions
$btnFlagGPOs.Add_Click({
    $textBoxLogs.AppendText("Flagging GPOs for review..." + [Environment]::NewLine)
    $allGPOs = Get-GPO -All
    $flaggedGPOs = $allGPOs | Where-Object {
        $_.ModificationTime -lt (Get-Date).AddDays(-30) -or
        $_.Links -eq $null
    }
    
    $lstGPOs.Items.Clear()
    foreach ($gpo in $flaggedGPOs) {
        $lstGPOs.Items.Add("$($gpo.DisplayName) | Last Modified: $($gpo.ModificationTime)")
    }
    
    $flaggedGPOs | Select-Object DisplayName, ModificationTime | Export-Csv -Path "FlaggedGPOsForReview.csv" -NoTypeInformation
    $textBoxLogs.AppendText("Flagged GPOs exported to FlaggedGPOsForReview.csv" + [Environment]::NewLine)
})

$btnDeleteGPOs.Add_Click({
    $textBoxLogs.AppendText("Deleting selected GPOs..." + [Environment]::NewLine)
    $progressBar.Value = 0
    $progressBar.Maximum = $lstGPOs.CheckedItems.Count

    foreach ($selectedItem in $lstGPOs.CheckedItems) {
        $gpoName = $selectedItem.Split('|')[0].Trim()
        $gpo = Get-GPO -Name $gpoName -ErrorAction SilentlyContinue
        if ($gpo) {
            Remove-GPO -Guid $gpo.Id -Confirm:$false
            $textBoxLogs.AppendText("Deleted GPO: $gpoName" + [Environment]::NewLine)
        }
        $progressBar.PerformStep()
    }

    $textBoxLogs.AppendText("Selected GPOs deleted. Report saved to DeletedGPOs.csv" + [Environment]::NewLine)
})

$btnSysVolCleanup.Add_Click({
    $textBoxLogs.AppendText("Cleaning up orphaned SysVol objects..." + [Environment]::NewLine)
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
        Remove-Item -Path $folder.FullName -Recurse -Force
        $textBoxLogs.AppendText("Deleted orphaned folder: $($folder.Name)" + [Environment]::NewLine)
        $progressBar.PerformStep()
    }

    $textBoxLogs.AppendText("SysVol cleanup complete. Report saved to OrphanedSysVolFolders.csv" + [Environment]::NewLine)
})

# Show the form
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
