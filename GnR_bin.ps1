Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

$form = New-Object System.Windows.Forms.Form
$form.Text = "GnR: Grep and Replace"
$form.Size = New-Object System.Drawing.Size(500, 620)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.AllowDrop = $true

$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Text = "1. Select Target Folder (Drag & Drop Supported):"
$lblPath.Location = New-Object System.Drawing.Point(20, 20)
$lblPath.AutoSize = $true
$form.Controls.Add($lblPath)

$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Location = New-Object System.Drawing.Point(20, 40)
$txtPath.Size = New-Object System.Drawing.Size(350, 20)
$txtPath.ReadOnly = $true
$form.Controls.Add($txtPath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse"
$btnBrowse.Location = New-Object System.Drawing.Point(380, 38)
$btnBrowse.Size = New-Object System.Drawing.Size(80, 23)
$btnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select a folder to process."
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtPath.Text = $dialog.SelectedPath
    }
})
$form.Controls.Add($btnBrowse)

$dragEnterHandler = {
    if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
}

$dragDropHandler = {
    $files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($files.Count -gt 0) {
        $path = $files[0]
        if ((Get-Item -LiteralPath $path) -is [System.IO.DirectoryInfo]) {
            $txtPath.Text = $path
        } else {
            [System.Windows.Forms.MessageBox]::Show("Only folders can be dragged and dropped.", "Info", 0, 48) | Out-Null
        }
    }
}

$form.Add_DragEnter($dragEnterHandler)
$form.Add_DragDrop($dragDropHandler)
$txtPath.AllowDrop = $true
$txtPath.Add_DragEnter($dragEnterHandler)
$txtPath.Add_DragDrop($dragDropHandler)

$lblExt = New-Object System.Windows.Forms.Label
$lblExt.Text = "2. Target Extensions for Content (Comma separated):"
$lblExt.Location = New-Object System.Drawing.Point(20, 80)
$lblExt.AutoSize = $true
$form.Controls.Add($lblExt)

$txtExt = New-Object System.Windows.Forms.TextBox
$txtExt.Location = New-Object System.Drawing.Point(20, 100)
$txtExt.Size = New-Object System.Drawing.Size(440, 20)
$txtExt.Text = "*.*"
$form.Controls.Add($txtExt)

$lblEncoding = New-Object System.Windows.Forms.Label
$lblEncoding.Text = "3. File Encoding:"
$lblEncoding.Location = New-Object System.Drawing.Point(20, 135)
$lblEncoding.AutoSize = $true
$form.Controls.Add($lblEncoding)

$cmbEncoding = New-Object System.Windows.Forms.ComboBox
$cmbEncoding.Location = New-Object System.Drawing.Point(20, 155)
$cmbEncoding.Size = New-Object System.Drawing.Size(200, 20)
$cmbEncoding.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbEncoding.Items.Add("Default")
$cmbEncoding.Items.Add("UTF-8")
$cmbEncoding.Items.Add("UTF-8 BOM")
$cmbEncoding.Items.Add("ANSI")
$cmbEncoding.Items.Add("SHIFT-JIS")
$cmbEncoding.SelectedIndex = 0
$form.Controls.Add($cmbEncoding)

$grpScope = New-Object System.Windows.Forms.GroupBox
$grpScope.Text = "4. Replace Scope"
$grpScope.Location = New-Object System.Drawing.Point(20, 190)
$grpScope.Size = New-Object System.Drawing.Size(440, 50)
$form.Controls.Add($grpScope)

$optContentOnly = New-Object System.Windows.Forms.RadioButton
$optContentOnly.Text = "Content Only"
$optContentOnly.Location = New-Object System.Drawing.Point(20, 20)
$optContentOnly.AutoSize = $true
$grpScope.Controls.Add($optContentOnly)

$optAll = New-Object System.Windows.Forms.RadioButton
$optAll.Text = "Content + File/Folder Names"
$optAll.Location = New-Object System.Drawing.Point(150, 20)
$optAll.AutoSize = $true
$optAll.Checked = $true
$grpScope.Controls.Add($optAll)

$lblRules = New-Object System.Windows.Forms.Label
$lblRules.Text = "5. Replace Rules (Format: Old|New):"
$lblRules.Location = New-Object System.Drawing.Point(20, 255)
$lblRules.AutoSize = $true
$form.Controls.Add($lblRules)

$btnClearRules = New-Object System.Windows.Forms.Button
$btnClearRules.Text = "Clear"
$btnClearRules.Location = New-Object System.Drawing.Point(400, 250)
$btnClearRules.Size = New-Object System.Drawing.Size(60, 22)
$btnClearRules.Add_Click({
    $txtRules.Text = ""
})
$form.Controls.Add($btnClearRules)

$txtRules = New-Object System.Windows.Forms.TextBox
$txtRules.Location = New-Object System.Drawing.Point(20, 275)
$txtRules.Size = New-Object System.Drawing.Size(440, 150)
$txtRules.Multiline = $true
$txtRules.ScrollBars = "Vertical"
$txtRules.Text = "OldKeyword1|NewKeyword1`r`nOldKeyword2|NewKeyword2"
$form.Controls.Add($txtRules)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Waiting..."
$lblStatus.Location = New-Object System.Drawing.Point(20, 440)
$lblStatus.Size = New-Object System.Drawing.Size(440, 40)
$lblStatus.ForeColor = "Blue"
$form.Controls.Add($lblStatus)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run & Create ZIP"
$btnRun.Location = New-Object System.Drawing.Point(20, 490)
$btnRun.Size = New-Object System.Drawing.Size(440, 50)
$btnRun.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$btnRun.Add_Click({
    $sourcePath = $txtPath.Text
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid folder.", "Error", 0, 16) | Out-Null
        return
    }

    $rules = $txtRules.Lines | Where-Object { $_ -match '\|' }
    if ($rules.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please enter at least one rule.`n(e.g., OldWord|NewWord)", "Error", 0, 16) | Out-Null
        return
    }

    $btnRun.Enabled = $false
    $lblStatus.Text = "Preparing (Copying to temp folder)..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "GrepReplace_$([guid]::NewGuid())"
        Copy-Item -Path $sourcePath -Destination $tempDir -Recurse -Force

        $extensions = $txtExt.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
        
            $encObj = switch ($cmbEncoding.SelectedItem.ToString()) {
                "Default" { [System.Text.Encoding]::Default }
                "UTF-8" { New-Object System.Text.UTF8Encoding($false) }
                "UTF-8 BOM" { New-Object System.Text.UTF8Encoding($true) }
                "ANSI" { [System.Text.Encoding]::Default }
                "SHIFT-JIS" { [System.Text.Encoding]::GetEncoding("shift_jis") }
                default { [System.Text.Encoding]::Default }
            }

            $lblStatus.Text = "Replacing file contents..."
            [System.Windows.Forms.Application]::DoEvents()
        
            $filesToContentReplace = Get-ChildItem -Path $tempDir -Recurse -File -Include $extensions
            foreach ($file in $filesToContentReplace) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
            
                if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { 
                    $readEnc = New-Object System.Text.UTF8Encoding($true) 
                }
                elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { 
                    $readEnc = [System.Text.Encoding]::Unicode 
                }
                else {
                    $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
                    try {
                        $null = $strictUtf8.GetString($bytes)
                        $readEnc = New-Object System.Text.UTF8Encoding($false)
                    }
                    catch {
                        $readEnc = [System.Text.Encoding]::GetEncoding("shift_jis")
                    }
                }

                $content = [System.IO.File]::ReadAllText($file.FullName, $readEnc)
                $isModified = $false

                foreach ($rule in $rules) {
                    $split = $rule -split '\|', 2
                    $search = $split[0]
                    $replace = $split[1]

                    if ($content -cmatch [regex]::Escape($search)) {
                        $content = $content.Replace($search, $replace)
                        $isModified = $true
                    }
                }

                if ($isModified) {
                    [System.IO.File]::WriteAllText($file.FullName, $content, $encObj)
                }
            }

        if ($optAll.Checked) {
            $lblStatus.Text = "Replacing file and folder names..."
            [System.Windows.Forms.Application]::DoEvents()

            $itemsToRename = Get-ChildItem -Path $tempDir -Recurse | Sort-Object -Property @{Expression={($_.FullName.Length - $_.FullName.Replace('\','').Length)}; Descending=$true}
            foreach ($item in $itemsToRename) {
                $newName = $item.Name
                $isModified = $false

                foreach ($rule in $rules) {
                    $split = $rule -split '\|', 2
                    $search = $split[0]
                    $replace = $split[1]

                    if ($newName.Contains($search)) {
                        $newName = $newName.Replace($search, $replace)
                        $isModified = $true
                    }
                }

                if ($isModified) {
                    Rename-Item -LiteralPath $item.FullName -NewName $newName
                }
            }
        }

        $lblStatus.Text = "Creating ZIP file..."
        [System.Windows.Forms.Application]::DoEvents()

        $parentFolder = Split-Path $sourcePath -Parent
        $folderName = Split-Path $sourcePath -Leaf
        $zipPath = Join-Path $parentFolder "${folderName}_Modified_$((Get-Date).ToString('yyyyMMdd_HHmmss')).zip"
        
        if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $zipPath)

        $lblStatus.Text = "Cleaning up temp files..."
        [System.Windows.Forms.Application]::DoEvents()
        Remove-Item -LiteralPath $tempDir -Recurse -Force

        $lblStatus.Text = "Done."
        [System.Windows.Forms.MessageBox]::Show("Process completed!`n`nSaved at: $zipPath", "Complete", 0, 64) | Out-Null

    } catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred:`n$_", "Error", 0, 16) | Out-Null
        $lblStatus.Text = "Error occurred"
    } finally {
        $btnRun.Enabled = $true
    }
})
$form.Controls.Add($btnRun)

$lblSignature = New-Object System.Windows.Forms.Label
$lblSignature.Text = "Created by D5 Kan: github.com/hyonimix"
$lblSignature.Location = New-Object System.Drawing.Point(250, 550)
$lblSignature.AutoSize = $true
$lblSignature.ForeColor = "DarkGray"
$form.Controls.Add($lblSignature)

$form.ShowDialog() | Out-Null