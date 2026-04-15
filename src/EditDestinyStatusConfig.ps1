Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$dataDir = Join-Path $env:APPDATA 'DestinyStatusDesktopData'
$legacyDataDir = Join-Path $env:APPDATA 'KickStatusAppData'
if (-not (Test-Path $dataDir)) {
    New-Item -ItemType Directory -Path $dataDir | Out-Null
}
$configPath = Join-Path $dataDir 'config.json'
$startupFolder = [Environment]::GetFolderPath('Startup')
$startupShortcut = Join-Path $startupFolder 'Destiny Status Desktop.lnk'
$launcherPath = Join-Path $PSScriptRoot 'DestinyStatusDesktop.exe'

if (-not (Test-Path $configPath) -and (Test-Path (Join-Path $legacyDataDir 'config.json'))) {
    Copy-Item (Join-Path $legacyDataDir 'config.json') $configPath -Force
}

if (-not (Test-Path $configPath)) {
    $legacyConfig = Join-Path $PSScriptRoot 'config.json'
    if (Test-Path $legacyConfig) {
        Copy-Item $legacyConfig $configPath -Force
    } else {
        @'
{
  "primaryChannel": "destiny",
  "secondaryChannel": "anythingelse",
  "priorityChannels": [],
  "primaryImagePath": "",
  "secondaryImagePath": "",
  "priorityImagePath": "",
  "priorityGifPath": ""
}
'@ | Set-Content -Path $configPath -Encoding UTF8
    }
}

$config = Get-Content $configPath -Raw | ConvertFrom-Json
function Normalize-ImagePath {
    param([string]$Path)

    if ($null -eq $Path) {
        return ''
    }

    $clean = $Path.Trim()
    if ($clean.Length -ge 2) {
        $startsWithQuote = $clean.StartsWith('"') -or $clean.StartsWith("'")
        $endsWithQuote = $clean.EndsWith('"') -or $clean.EndsWith("'")
        if ($startsWithQuote -and $endsWithQuote) {
            $clean = $clean.Substring(1, $clean.Length - 2).Trim()
        }
    }

    return $clean
}

function Get-StringArray {
    param([object]$Value)

    return @(
        $Value |
        ForEach-Object { [string]$_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_.Trim() }
    )
}

function Set-ConfigProperty {
    param(
        [object]$Target,
        [string]$Name,
        [object]$Value
    )

    if ($Target.PSObject.Properties.Name -contains $Name) {
        $Target.$Name = $Value
    } else {
        $Target | Add-Member -NotePropertyName $Name -NotePropertyValue $Value
    }
}

function Get-RuleDisplayValue {
    param(
        [object]$Config,
        [string[]]$Channels
    )

    $targetChannels = Get-StringArray $Channels
    if ($targetChannels.Count -eq 0) {
        return ''
    }

    $rules = @($Config.priorityRules)
    foreach ($rule in $rules) {
        $ruleChannels = Get-StringArray $rule.channels
        if ($ruleChannels.Count -eq 1 -and $targetChannels.Count -eq 1 -and
            [string]::Equals($ruleChannels[0], $targetChannels[0], [System.StringComparison]::OrdinalIgnoreCase)) {
            return Normalize-ImagePath ([string]$rule.displayValue)
        }
    }

    foreach ($rule in $rules) {
        $ruleChannels = Get-StringArray $rule.channels
        foreach ($ruleChannel in $ruleChannels) {
            if ($targetChannels -icontains $ruleChannel) {
                return Normalize-ImagePath ([string]$rule.displayValue)
            }
        }
    }

    return ''
}

function Update-RuleDisplayValue {
    param(
        [object]$Config,
        [string[]]$Channels,
        [string]$DisplayValue
    )

    $targetChannels = Get-StringArray $Channels
    if ($targetChannels.Count -eq 0 -or $null -eq $Config.priorityRules) {
        return
    }

    foreach ($rule in @($Config.priorityRules)) {
        $ruleChannels = Get-StringArray $rule.channels
        if ($ruleChannels.Count -ne $targetChannels.Count) {
            continue
        }

        $matchesAll = $true
        foreach ($targetChannel in $targetChannels) {
            if ($ruleChannels -inotcontains $targetChannel) {
                $matchesAll = $false
                break
            }
        }

        if (-not $matchesAll) {
            continue
        }

        Set-ConfigProperty -Target $rule -Name 'displayMode' -Value 'Gif'
        Set-ConfigProperty -Target $rule -Name 'displayValue' -Value $DisplayValue
        return
    }
}

Set-ConfigProperty -Target $config -Name 'primaryImagePath' -Value ([string]$config.primaryImagePath)
Set-ConfigProperty -Target $config -Name 'secondaryImagePath' -Value ([string]$config.secondaryImagePath)
Set-ConfigProperty -Target $config -Name 'priorityImagePath' -Value ([string]$config.priorityImagePath)
Set-ConfigProperty -Target $config -Name 'priorityGifPath' -Value ([string]$config.priorityGifPath)

$config.primaryImagePath = Normalize-ImagePath ([string]$config.primaryImagePath)
$config.secondaryImagePath = Normalize-ImagePath ([string]$config.secondaryImagePath)
$config.priorityImagePath = Normalize-ImagePath ([string]$config.priorityImagePath)
$config.priorityGifPath = Normalize-ImagePath ([string]$config.priorityGifPath)

if ([string]::IsNullOrWhiteSpace($config.primaryImagePath)) {
    $config.primaryImagePath = Get-RuleDisplayValue -Config $config -Channels @([string]$config.primaryChannel)
}
if ([string]::IsNullOrWhiteSpace($config.secondaryImagePath)) {
    $config.secondaryImagePath = Get-RuleDisplayValue -Config $config -Channels @([string]$config.secondaryChannel)
}
if ([string]::IsNullOrWhiteSpace($config.priorityImagePath)) {
    $config.priorityImagePath = Get-RuleDisplayValue -Config $config -Channels (Get-StringArray $config.priorityChannels)
}
if ([string]::IsNullOrWhiteSpace($config.priorityImagePath)) {
    $config.priorityImagePath = $config.priorityGifPath
}
if ([string]::IsNullOrWhiteSpace($config.priorityGifPath)) {
    $config.priorityGifPath = $config.priorityImagePath
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Edit Destiny Status Channels'
$form.StartPosition = 'CenterScreen'
$form.ClientSize = New-Object System.Drawing.Size(520, 620)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $false

$labelPrimary = New-Object System.Windows.Forms.Label
$labelPrimary.Text = 'Primary channel'
$labelPrimary.Location = New-Object System.Drawing.Point(20, 20)
$labelPrimary.AutoSize = $true
$form.Controls.Add($labelPrimary)

$tbPrimary = New-Object System.Windows.Forms.TextBox
$tbPrimary.Location = New-Object System.Drawing.Point(20, 42)
$tbPrimary.Size = New-Object System.Drawing.Size(220, 24)
$tbPrimary.Text = [string]$config.primaryChannel
$form.Controls.Add($tbPrimary)

$labelSecondary = New-Object System.Windows.Forms.Label
$labelSecondary.Text = 'Secondary channel'
$labelSecondary.Location = New-Object System.Drawing.Point(270, 20)
$labelSecondary.AutoSize = $true
$form.Controls.Add($labelSecondary)

$tbSecondary = New-Object System.Windows.Forms.TextBox
$tbSecondary.Location = New-Object System.Drawing.Point(270, 42)
$tbSecondary.Size = New-Object System.Drawing.Size(220, 24)
$tbSecondary.Text = [string]$config.secondaryChannel
$form.Controls.Add($tbSecondary)

function Update-ImagePreview {
    param(
        [pscustomobject]$Row,
        [string]$FallbackText
    )

    if ($null -ne $Row.PreviewImage) {
        $Row.Picture.Image = $null
        $Row.PreviewImage.Dispose()
        $Row.PreviewImage = $null
    }

    $path = Normalize-ImagePath $Row.TextBox.Text
    if ($path -ne $Row.TextBox.Text) {
        $Row.TextBox.Text = $path
        return
    }

    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path $path)) {
        $Row.Picture.Visible = $false
        $Row.Fallback.Visible = $true
        $Row.Fallback.Text = $FallbackText
        return
    }

    try {
        $Row.PreviewImage = [System.Drawing.Image]::FromFile($path)
        $Row.Picture.Image = $Row.PreviewImage
        $Row.Picture.Visible = $true
        $Row.Fallback.Visible = $false
    }
    catch {
        $Row.Picture.Visible = $false
        $Row.Fallback.Visible = $true
        $Row.Fallback.Text = 'ERR'
    }
}

function New-ImageSelectorRow {
    param(
        [string]$LabelText,
        [string]$InitialValue,
        [string]$FallbackText,
        [int]$Y
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $LabelText
    $label.Location = New-Object System.Drawing.Point(20, $Y)
    $label.AutoSize = $true
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(20, ($Y + 22))
    $textBox.Size = New-Object System.Drawing.Size(300, 24)
    $textBox.Text = $InitialValue
    $form.Controls.Add($textBox)

    $browse = New-Object System.Windows.Forms.Button
    $browse.Text = 'Browse'
    $browse.Location = New-Object System.Drawing.Point(330, ($Y + 20))
    $browse.Size = New-Object System.Drawing.Size(70, 28)
    $form.Controls.Add($browse)

    $previewPanel = New-Object System.Windows.Forms.Panel
    $previewPanel.Location = New-Object System.Drawing.Point(420, $Y)
    $previewPanel.Size = New-Object System.Drawing.Size(64, 64)
    $previewPanel.BorderStyle = 'FixedSingle'
    $previewPanel.BackColor = [System.Drawing.Color]::LimeGreen
    $form.Controls.Add($previewPanel)

    $picture = New-Object System.Windows.Forms.PictureBox
    $picture.Dock = 'Fill'
    $picture.BackColor = [System.Drawing.Color]::Transparent
    $picture.SizeMode = 'Zoom'
    $picture.Visible = $false
    $previewPanel.Controls.Add($picture)

    $fallback = New-Object System.Windows.Forms.Label
    $fallback.Dock = 'Fill'
    $fallback.BackColor = [System.Drawing.Color]::Transparent
    $fallback.ForeColor = [System.Drawing.Color]::White
    $fallback.TextAlign = 'MiddleCenter'
    $fallback.Font = if ($FallbackText.Length -gt 1) {
        New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    } else {
        New-Object System.Drawing.Font('Segoe UI', 18, [System.Drawing.FontStyle]::Bold)
    }
    $fallback.Text = $FallbackText
    $previewPanel.Controls.Add($fallback)
    $fallback.BringToFront()

    $row = [pscustomobject]@{
        TextBox = $textBox
        Browse = $browse
        Picture = $picture
        Fallback = $fallback
        PreviewImage = $null
    }

    $browse.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Image files (*.gif;*.png;*.jpg;*.jpeg;*.bmp;*.webp)|*.gif;*.png;*.jpg;*.jpeg;*.bmp;*.webp|All files (*.*)|*.*'
        $dialog.FileName = Normalize-ImagePath $row.TextBox.Text
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $row.TextBox.Text = $dialog.FileName
        }
    })

    $textBox.Add_TextChanged({
        Update-ImagePreview -Row $row -FallbackText $FallbackText
    })

    return $row
}

$primaryImageRow = New-ImageSelectorRow -LabelText 'Primary image path' -InitialValue ([string]$config.primaryImagePath) -FallbackText 'D' -Y 82
$secondaryImageRow = New-ImageSelectorRow -LabelText 'Secondary image path' -InitialValue ([string]$config.secondaryImagePath) -FallbackText 'AE' -Y 160
$priorityImageRow = New-ImageSelectorRow -LabelText 'Priority image path' -InitialValue ([string]$config.priorityImagePath) -FallbackText 'P' -Y 238

$labelPriority = New-Object System.Windows.Forms.Label
$labelPriority.Text = 'Priority channels, one slug per line'
$labelPriority.Location = New-Object System.Drawing.Point(20, 330)
$labelPriority.AutoSize = $true
$form.Controls.Add($labelPriority)

$tbPriority = New-Object System.Windows.Forms.TextBox
$tbPriority.Location = New-Object System.Drawing.Point(20, 353)
$tbPriority.Size = New-Object System.Drawing.Size(470, 170)
$tbPriority.Multiline = $true
$tbPriority.ScrollBars = 'Vertical'
$tbPriority.AcceptsReturn = $true
$tbPriority.AcceptsTab = $true
$tbPriority.Text = (($config.priorityChannels | ForEach-Object { [string]$_ }) -join [Environment]::NewLine)
$form.Controls.Add($tbPriority)

$cbStartup = New-Object System.Windows.Forms.CheckBox
$cbStartup.Text = 'Launch square on Windows startup'
$cbStartup.Location = New-Object System.Drawing.Point(20, 545)
$cbStartup.AutoSize = $true
$cbStartup.Checked = Test-Path $startupShortcut
$form.Controls.Add($cbStartup)

$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = 'Save'
$saveButton.Location = New-Object System.Drawing.Point(330, 543)
$saveButton.Size = New-Object System.Drawing.Size(75, 30)
$form.Controls.Add($saveButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = 'Cancel'
$cancelButton.Location = New-Object System.Drawing.Point(415, 543)
$cancelButton.Size = New-Object System.Drawing.Size(75, 30)
$form.Controls.Add($cancelButton)

$form.Add_Shown({
    Update-ImagePreview -Row $primaryImageRow -FallbackText 'D'
    Update-ImagePreview -Row $secondaryImageRow -FallbackText 'AE'
    Update-ImagePreview -Row $priorityImageRow -FallbackText 'P'
})

$form.Add_FormClosed({
    foreach ($row in @($primaryImageRow, $secondaryImageRow, $priorityImageRow)) {
        if ($null -ne $row.PreviewImage) {
            $row.PreviewImage.Dispose()
            $row.PreviewImage = $null
        }
    }
})

$saveButton.Add_Click({
    $priorityChannels = @(
        $tbPriority.Lines |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ }
    )

    $newConfig = $config.PSObject.Copy()
    $newPrimaryChannel = $tbPrimary.Text.Trim()
    $newSecondaryChannel = $tbSecondary.Text.Trim()
    $newPrimaryImagePath = Normalize-ImagePath $primaryImageRow.TextBox.Text
    $newSecondaryImagePath = Normalize-ImagePath $secondaryImageRow.TextBox.Text
    $newPriorityImagePath = Normalize-ImagePath $priorityImageRow.TextBox.Text

    Set-ConfigProperty -Target $newConfig -Name 'primaryChannel' -Value $newPrimaryChannel
    Set-ConfigProperty -Target $newConfig -Name 'secondaryChannel' -Value $newSecondaryChannel
    Set-ConfigProperty -Target $newConfig -Name 'priorityChannels' -Value $priorityChannels
    Set-ConfigProperty -Target $newConfig -Name 'primaryImagePath' -Value $newPrimaryImagePath
    Set-ConfigProperty -Target $newConfig -Name 'secondaryImagePath' -Value $newSecondaryImagePath
    Set-ConfigProperty -Target $newConfig -Name 'priorityImagePath' -Value $newPriorityImagePath
    Set-ConfigProperty -Target $newConfig -Name 'priorityGifPath' -Value $newPriorityImagePath

    Update-RuleDisplayValue -Config $newConfig -Channels @($newPrimaryChannel) -DisplayValue $newPrimaryImagePath
    Update-RuleDisplayValue -Config $newConfig -Channels @($newSecondaryChannel) -DisplayValue $newSecondaryImagePath
    Update-RuleDisplayValue -Config $newConfig -Channels $priorityChannels -DisplayValue $newPriorityImagePath

    $newConfig | ConvertTo-Json -Depth 4 | Set-Content -Path $configPath -Encoding UTF8

    if ($cbStartup.Checked) {
        $wsh = New-Object -ComObject WScript.Shell
        $shortcut = $wsh.CreateShortcut($startupShortcut)
        $shortcut.TargetPath = $launcherPath
        $shortcut.WorkingDirectory = $PSScriptRoot
        $shortcut.Save()
    } elseif (Test-Path $startupShortcut) {
        Remove-Item $startupShortcut -Force
    }

    [System.Windows.Forms.MessageBox]::Show('Saved. Relaunch Destiny Status Desktop to use the new channels.', 'Destiny Status Desktop')
    $form.Close()
})

$cancelButton.Add_Click({
    $form.Close()
})

[void]$form.ShowDialog()
