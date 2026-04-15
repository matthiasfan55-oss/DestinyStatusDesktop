Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

[System.Windows.Forms.Application]::EnableVisualStyles()

$dataDir = Join-Path $env:APPDATA 'KickStatusAppData'
if (-not (Test-Path $dataDir)) {
    New-Item -ItemType Directory -Path $dataDir | Out-Null
}

$configPath = Join-Path $dataDir 'config.json'
$startupFolder = [Environment]::GetFolderPath('Startup')
$startupShortcut = Join-Path $startupFolder 'Kick Status.lnk'
$launcherPath = Join-Path $PSScriptRoot 'KickStatusSquare.exe'

function Normalize-ImagePath($path) {
    if ($null -eq $path) {
        return ''
    }

    $clean = ([string]$path).Trim()
    if ($clean.Length -ge 2) {
        $startsWithQuote = $clean.StartsWith('"') -or $clean.StartsWith("'")
        $endsWithQuote = $clean.EndsWith('"') -or $clean.EndsWith("'")
        if ($startsWithQuote -and $endsWithQuote) {
            $clean = $clean.Substring(1, $clean.Length - 2).Trim()
        }
    }

    $markdownStart = $clean.IndexOf('](')
    if ($markdownStart -ge 0 -and $clean.EndsWith(')')) {
        $clean = $clean.Substring($markdownStart + 2, $clean.Length - $markdownStart - 3).Trim()
    }

    return $clean
}

function Load-PreviewImage($path) {
    $normalizedPath = Normalize-ImagePath $path
    if ([string]::IsNullOrWhiteSpace($normalizedPath) -or -not (Test-Path $normalizedPath)) {
        return $null
    }

    try {
        $stream = New-Object System.IO.FileStream($normalizedPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $loaded = [System.Drawing.Image]::FromStream($stream)
            try {
                return New-Object System.Drawing.Bitmap -ArgumentList $loaded
            }
            finally {
                $loaded.Dispose()
            }
        }
        finally {
            $stream.Dispose()
        }
    }
    catch {
    }

    try {
        $uri = New-Object System.Uri($normalizedPath, [System.UriKind]::Absolute)
        $decoder = [System.Windows.Media.Imaging.BitmapDecoder]::Create(
            $uri,
            [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat,
            [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        )
        $frame = $decoder.Frames[0]
        if ($null -eq $frame) {
            return $null
        }

        $encoder = New-Object System.Windows.Media.Imaging.BmpBitmapEncoder
        [void]$encoder.Frames.Add($frame)

        $memoryStream = New-Object System.IO.MemoryStream
        try {
            $encoder.Save($memoryStream)
            $memoryStream.Position = 0
            $loaded = [System.Drawing.Image]::FromStream($memoryStream)
            try {
                return New-Object System.Drawing.Bitmap -ArgumentList $loaded
            }
            finally {
                $loaded.Dispose()
            }
        }
        finally {
            $memoryStream.Dispose()
        }
    }
    catch {
        return $null
    }
}

function New-DefaultPriorityRules {
    @(
        [pscustomobject]@{
            channels = @('destiny')
            displayMode = 'Text'
            displayValue = 'D'
        },
        [pscustomobject]@{
            channels = @('anythingelse')
            displayMode = 'Text'
            displayValue = 'AE'
        },
        [pscustomobject]@{
            channels = @()
            displayMode = 'Gif'
            displayValue = ''
        }
    )
}

function Normalize-Rule($rule) {
    if ($null -eq $rule) {
        $rule = [pscustomobject]@{}
    }

    $channels = @()
    if ($rule.PSObject.Properties['channels']) {
        $channels = @($rule.channels | ForEach-Object { [string]$_ } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    $displayMode = 'Text'
    if ($rule.PSObject.Properties['displayMode'] -and $rule.displayMode -eq 'Gif') {
        $displayMode = 'Gif'
    }

    $displayValue = ''
    if ($rule.PSObject.Properties['displayValue'] -and $null -ne $rule.displayValue) {
        $displayValue = [string]$rule.displayValue
    }
    if ($displayMode -eq 'Gif') {
        $displayValue = Normalize-ImagePath $displayValue
    }

    [pscustomobject]@{
        channels = $channels
        displayMode = $displayMode
        displayValue = $displayValue
    }
}

function Normalize-Config($cfg) {
    if ($null -eq $cfg) {
        $cfg = [pscustomobject]@{}
    }

    $rules = @()
    if ($cfg.PSObject.Properties['priorityRules'] -and $cfg.priorityRules) {
        $rules = @($cfg.priorityRules | ForEach-Object { Normalize-Rule $_ })
    } else {
        $rules = New-DefaultPriorityRules

        if ($cfg.PSObject.Properties['primaryChannel'] -and $cfg.primaryChannel) {
            $rules[0].channels = @([string]$cfg.primaryChannel)
        }
        if ($cfg.PSObject.Properties['secondaryChannel'] -and $cfg.secondaryChannel) {
            $rules[1].channels = @([string]$cfg.secondaryChannel)
        }
        if ($cfg.PSObject.Properties['priorityChannels'] -and $cfg.priorityChannels) {
            $rules[2].channels = @($cfg.priorityChannels | ForEach-Object { [string]$_ } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
        if ($cfg.PSObject.Properties['priorityGifPath'] -and $cfg.priorityGifPath) {
            $rules[2].displayMode = 'Gif'
            $rules[2].displayValue = [string]$cfg.priorityGifPath
        }
    }

    while ($rules.Count -lt 3) {
        $rules += [pscustomobject]@{
            channels = @()
            displayMode = 'Text'
            displayValue = ''
        }
    }

    if (-not $cfg.PSObject.Properties['openKickOnClick']) {
        $cfg | Add-Member -NotePropertyName openKickOnClick -NotePropertyValue $true
    }
    if (-not $cfg.PSObject.Properties['openBigscreenInChromeOnClick']) {
        $cfg | Add-Member -NotePropertyName openBigscreenInChromeOnClick -NotePropertyValue $false
    }
    if (-not $cfg.PSObject.Properties['refreshMinutes']) {
        $cfg | Add-Member -NotePropertyName refreshMinutes -NotePropertyValue 1.0
    }
    if ([double]$cfg.refreshMinutes -le 0) {
        $cfg.refreshMinutes = 1.0
    }

    [pscustomobject]@{
        priorityRules = $rules
        openKickOnClick = [bool]$cfg.openKickOnClick
        openBigscreenInChromeOnClick = [bool]$cfg.openBigscreenInChromeOnClick
        refreshMinutes = [double]$cfg.refreshMinutes
    }
}

if (-not (Test-Path $configPath)) {
    $defaultConfig = [pscustomobject]@{
        priorityRules = New-DefaultPriorityRules
        openKickOnClick = $true
        openBigscreenInChromeOnClick = $false
        refreshMinutes = 1.0
    }
    $defaultConfig | ConvertTo-Json -Depth 6 | Set-Content -Path $configPath -Encoding UTF8
}

$config = Normalize-Config (Get-Content $configPath -Raw | ConvertFrom-Json)
$config | ConvertTo-Json -Depth 6 | Set-Content -Path $configPath -Encoding UTF8

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Edit Kick Status Channels'
$form.StartPosition = 'CenterScreen'
$form.ClientSize = New-Object System.Drawing.Size(1100, 1100)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.Font = New-Object System.Drawing.Font('Segoe UI', 11)

$priorityPanel = New-Object System.Windows.Forms.Panel
$priorityPanel.Location = New-Object System.Drawing.Point(20, 20)
$priorityPanel.Size = New-Object System.Drawing.Size(1060, 820)
$priorityPanel.AutoScroll = $true
$priorityPanel.BorderStyle = 'FixedSingle'
$form.Controls.Add($priorityPanel)

$ruleControls = New-Object System.Collections.ArrayList

function Dispose-RulePreview($entry) {
    if ($null -eq $entry -or $null -eq $entry.previewImage) {
        return
    }

    $entry.previewPicture.Image = $null
    $entry.previewImage.Dispose()
    $entry.previewImage = $null
}

function Update-RulePreview($entry) {
    if ($null -eq $entry) {
        return
    }

    Dispose-RulePreview $entry

    $mode = [string]$entry.modeCombo.SelectedItem
    $isGif = $mode -eq 'Gif'
    $entry.previewPanel.Visible = $isGif
    if (-not $isGif) {
        return
    }

    $path = Normalize-ImagePath $entry.displayTextBox.Text
    if ($path -ne $entry.displayTextBox.Text) {
        $entry.displayTextBox.Text = $path
        return
    }

    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path $path)) {
        $entry.previewPicture.Visible = $false
        $entry.previewFallback.Visible = $true
        $entry.previewFallback.Text = 'IMG'
        return
    }

    $entry.previewImage = Load-PreviewImage $path
    if ($null -ne $entry.previewImage) {
        $entry.previewPicture.Image = $entry.previewImage
        $entry.previewPicture.Visible = $true
        $entry.previewFallback.Visible = $false
    }
    else {
        $entry.previewPicture.Visible = $false
        $entry.previewFallback.Visible = $true
        $entry.previewFallback.Text = 'ERR'
    }
}

function Update-RuleLayout {
    $y = 10
    for ($i = 0; $i -lt $ruleControls.Count; $i++) {
        $entry = $ruleControls[$i]
        $entry.container.Location = New-Object System.Drawing.Point(10, $y)
        if ($i -eq 0) {
            $entry.titleLabel.Text = 'Primary'
        } elseif ($i -eq 1) {
            $entry.titleLabel.Text = 'Secondary'
        } else {
            $entry.titleLabel.Text = "Priority #$($i - 1)"
        }
        $entry.moveUpButton.Enabled = ($i -gt 0)
        $entry.moveDownButton.Enabled = ($i -lt ($ruleControls.Count - 1))
        $entry.removeButton.Enabled = ($ruleControls.Count -gt 1)
        $mode = [string]$entry.modeCombo.SelectedItem
        if ([string]::IsNullOrWhiteSpace($mode)) {
            $mode = 'Text'
        }
        $entry.browseButton.Visible = ($mode -eq 'Gif')
        $entry.displayLabel.Text = if ($mode -eq 'Gif') { 'Image path' } else { 'Display text' }
        if ($mode -eq 'Gif') {
            $entry.displayTextBox.Size = New-Object System.Drawing.Size(300, 32)
            $entry.previewPanel.Visible = $true
        } else {
            $entry.displayTextBox.Size = New-Object System.Drawing.Size(470, 32)
            $entry.previewPanel.Visible = $false
        }
        Update-RulePreview $entry
        $y += $entry.container.Height + 10
    }
}

function Move-RuleEntry($entry, $offset) {
    if ($null -eq $entry) {
        return
    }

    $currentIndex = $ruleControls.IndexOf($entry)
    if ($currentIndex -lt 0) {
        return
    }

    $targetIndex = $currentIndex + $offset
    if ($targetIndex -lt 0 -or $targetIndex -ge $ruleControls.Count) {
        return
    }

    $ruleControls.RemoveAt($currentIndex)
    $ruleControls.Insert($targetIndex, $entry)
    Update-RuleLayout
    $entry.container.Focus()
}

function Add-RuleEditor($rule) {
    $rule = Normalize-Rule $rule

    $container = New-Object System.Windows.Forms.Panel
    $container.Size = New-Object System.Drawing.Size(1020, 320)
    $container.BorderStyle = 'FixedSingle'

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(12, 10)
    $titleLabel.AutoSize = $true
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $titleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
    $container.Controls.Add($titleLabel)

    $moveUpButton = New-Object System.Windows.Forms.Button
    $moveUpButton.Text = 'Up'
    $moveUpButton.Size = New-Object System.Drawing.Size(78, 36)
    $moveUpButton.Location = New-Object System.Drawing.Point(670, 8)
    $container.Controls.Add($moveUpButton)

    $moveDownButton = New-Object System.Windows.Forms.Button
    $moveDownButton.Text = 'Down'
    $moveDownButton.Size = New-Object System.Drawing.Size(96, 36)
    $moveDownButton.Location = New-Object System.Drawing.Point(754, 8)
    $container.Controls.Add($moveDownButton)

    $removeButton = New-Object System.Windows.Forms.Button
    $removeButton.Text = 'Remove'
    $removeButton.Size = New-Object System.Drawing.Size(112, 36)
    $removeButton.Location = New-Object System.Drawing.Point(860, 8)
    $container.Controls.Add($removeButton)

    $channelsLabel = New-Object System.Windows.Forms.Label
    $channelsLabel.Text = 'Kick channel'
    $channelsLabel.Location = New-Object System.Drawing.Point(12, 60)
    $channelsLabel.AutoSize = $true
    $channelsLabel.BackColor = [System.Drawing.Color]::Transparent
    $container.Controls.Add($channelsLabel)

    $channelsBox = New-Object System.Windows.Forms.TextBox
    $channelsBox.Location = New-Object System.Drawing.Point(12, 112)
    $channelsBox.Size = New-Object System.Drawing.Size(470, 160)
    $channelsBox.Multiline = $true
    $channelsBox.ScrollBars = 'Vertical'
    $channelsBox.AcceptsReturn = $true
    $channelsBox.AcceptsTab = $true
    $channelsBox.Font = New-Object System.Drawing.Font('Segoe UI', 11)
    $channelsBox.Text = (($rule.channels | ForEach-Object { [string]$_ }) -join [Environment]::NewLine)
    $container.Controls.Add($channelsBox)

    $modeLabel = New-Object System.Windows.Forms.Label
    $modeLabel.Text = 'Display'
    $modeLabel.Location = New-Object System.Drawing.Point(520, 60)
    $modeLabel.AutoSize = $true
    $modeLabel.BackColor = [System.Drawing.Color]::Transparent
    $container.Controls.Add($modeLabel)

    $modeCombo = New-Object System.Windows.Forms.ComboBox
    $modeCombo.Location = New-Object System.Drawing.Point(520, 112)
    $modeCombo.Size = New-Object System.Drawing.Size(170, 34)
    $modeCombo.DropDownStyle = 'DropDownList'
    $modeCombo.Font = New-Object System.Drawing.Font('Segoe UI', 11)
    $modeCombo.IntegralHeight = $false
    [void]$modeCombo.Items.Add('Text')
    [void]$modeCombo.Items.Add('Gif')
    $modeCombo.SelectedItem = if ($rule.displayMode -eq 'Gif') { 'Gif' } else { 'Text' }
    $container.Controls.Add($modeCombo)

    $displayLabel = New-Object System.Windows.Forms.Label
    $displayLabel.Location = New-Object System.Drawing.Point(520, 176)
    $displayLabel.AutoSize = $true
    $displayLabel.BackColor = [System.Drawing.Color]::Transparent
    $container.Controls.Add($displayLabel)

    $displayTextBox = New-Object System.Windows.Forms.TextBox
    $displayTextBox.Location = New-Object System.Drawing.Point(520, 228)
    $displayTextBox.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $displayTextBox.Text = [string]$rule.displayValue
    $container.Controls.Add($displayTextBox)

    $previewPanel = New-Object System.Windows.Forms.Panel
    $previewPanel.Location = New-Object System.Drawing.Point(830, 220)
    $previewPanel.Size = New-Object System.Drawing.Size(66, 66)
    $previewPanel.BorderStyle = 'FixedSingle'
    $previewPanel.BackColor = [System.Drawing.Color]::LimeGreen
    $container.Controls.Add($previewPanel)

    $previewPicture = New-Object System.Windows.Forms.PictureBox
    $previewPicture.Dock = 'Fill'
    $previewPicture.BackColor = [System.Drawing.Color]::Transparent
    $previewPicture.SizeMode = 'Zoom'
    $previewPicture.Visible = $false
    $previewPanel.Controls.Add($previewPicture)

    $previewFallback = New-Object System.Windows.Forms.Label
    $previewFallback.Dock = 'Fill'
    $previewFallback.BackColor = [System.Drawing.Color]::Transparent
    $previewFallback.ForeColor = [System.Drawing.Color]::White
    $previewFallback.TextAlign = 'MiddleCenter'
    $previewFallback.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $previewFallback.Text = 'IMG'
    $previewPanel.Controls.Add($previewFallback)
    $previewFallback.BringToFront()

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Text = 'Browse'
    $browseButton.Size = New-Object System.Drawing.Size(104, 36)
    $browseButton.Location = New-Object System.Drawing.Point(900, 230)
    $container.Controls.Add($browseButton)

    $browseButton.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Image files (*.gif;*.png;*.jpg;*.jpeg;*.bmp;*.webp)|*.gif;*.png;*.jpg;*.jpeg;*.bmp;*.webp|All files (*.*)|*.*'
        $dialog.FileName = Normalize-ImagePath $displayTextBox.Text
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $displayTextBox.Text = $dialog.FileName
        }
    })

    $entry = $null

    $moveUpButton.Add_Click({
        Move-RuleEntry $entry -1
    })

    $moveDownButton.Add_Click({
        Move-RuleEntry $entry 1
    })

    $removeButton.Add_Click({
        Dispose-RulePreview $entry
        $priorityPanel.Controls.Remove($container)
        [void]$ruleControls.Remove($entry)
        Update-RuleLayout
    })

    $entry = [pscustomobject]@{
        container = $container
        titleLabel = $titleLabel
        moveUpButton = $moveUpButton
        moveDownButton = $moveDownButton
        removeButton = $removeButton
        channelsBox = $channelsBox
        modeCombo = $modeCombo
        displayLabel = $displayLabel
        displayTextBox = $displayTextBox
        previewPanel = $previewPanel
        previewPicture = $previewPicture
        previewFallback = $previewFallback
        previewImage = $null
        browseButton = $browseButton
    }

    $titleLabel.BringToFront()
    $channelsLabel.BringToFront()
    $modeLabel.BringToFront()
    $displayLabel.BringToFront()

    $modeCombo.Add_SelectedIndexChanged({
        Update-RuleLayout
        Update-RulePreview $entry
    })
    $displayTextBox.Add_TextChanged({
        Update-RulePreview $entry
    })

    [void]$ruleControls.Add($entry)
    $priorityPanel.Controls.Add($container)
    Update-RulePreview $entry
    Update-RuleLayout
}

foreach ($rule in $config.priorityRules) {
    Add-RuleEditor $rule
}

$addPriorityButton = New-Object System.Windows.Forms.Button
$addPriorityButton.Text = 'Add Priority'
$addPriorityButton.Location = New-Object System.Drawing.Point(20, 860)
$addPriorityButton.Size = New-Object System.Drawing.Size(170, 42)
$form.Controls.Add($addPriorityButton)

$addPriorityButton.Add_Click({
    Add-RuleEditor ([pscustomobject]@{
        channels = @()
        displayMode = 'Text'
        displayValue = ''
    })
})

$labelRefresh = New-Object System.Windows.Forms.Label
$labelRefresh.Text = 'Refresh interval (minutes, decimals allowed)'
$labelRefresh.Location = New-Object System.Drawing.Point(94, 930)
$labelRefresh.AutoSize = $true
$form.Controls.Add($labelRefresh)

$tbRefresh = New-Object System.Windows.Forms.TextBox
$tbRefresh.Location = New-Object System.Drawing.Point(20, 926)
$tbRefresh.Size = New-Object System.Drawing.Size(62, 32)
$tbRefresh.Text = ([string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0:0.###}', [double]$config.refreshMinutes))
$form.Controls.Add($tbRefresh)

$cbStartup = New-Object System.Windows.Forms.CheckBox
$cbStartup.Text = 'Launch square on Windows startup'
$cbStartup.Location = New-Object System.Drawing.Point(20, 972)
$cbStartup.AutoSize = $true
$cbStartup.Checked = Test-Path $startupShortcut
$form.Controls.Add($cbStartup)

$cbOpenKick = New-Object System.Windows.Forms.CheckBox
$cbOpenKick.Text = 'Open Kick when square is clicked'
$cbOpenKick.Location = New-Object System.Drawing.Point(20, 1004)
$cbOpenKick.AutoSize = $true
$cbOpenKick.Checked = [bool]$config.openKickOnClick
$form.Controls.Add($cbOpenKick)

$cbBigscreen = New-Object System.Windows.Forms.CheckBox
$cbBigscreen.Text = 'Also open destiny.gg/bigscreen in a new Chrome tab'
$cbBigscreen.Location = New-Object System.Drawing.Point(20, 1036)
$cbBigscreen.AutoSize = $true
$cbBigscreen.Checked = [bool]$config.openBigscreenInChromeOnClick
$form.Controls.Add($cbBigscreen)

$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = 'Save'
$saveButton.Location = New-Object System.Drawing.Point(858, 1026)
$saveButton.Size = New-Object System.Drawing.Size(96, 38)
$form.Controls.Add($saveButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = 'Cancel'
$cancelButton.Location = New-Object System.Drawing.Point(964, 1026)
$cancelButton.Size = New-Object System.Drawing.Size(96, 38)
$form.Controls.Add($cancelButton)

$saveButton.Add_Click({
    $priorityRules = @()

    foreach ($entry in $ruleControls) {
        $channels = @(
            $entry.channelsBox.Lines |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ }
        )

        $displayMode = if ([string]$entry.modeCombo.SelectedItem -eq 'Gif') { 'Gif' } else { 'Text' }
        $displayValue = if ($displayMode -eq 'Gif') {
            Normalize-ImagePath $entry.displayTextBox.Text
        } else {
            $entry.displayTextBox.Text.Trim()
        }

        $priorityRules += [pscustomobject]@{
            channels = $channels
            displayMode = $displayMode
            displayValue = $displayValue
        }
    }

    $refreshMinutes = 1.0
    if (-not [double]::TryParse($tbRefresh.Text.Trim(), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$refreshMinutes)) {
        [System.Windows.Forms.MessageBox]::Show('Refresh interval must be a number in minutes.', 'Kick Status')
        return
    }
    if ($refreshMinutes -le 0) {
        [System.Windows.Forms.MessageBox]::Show('Refresh interval must be greater than 0.', 'Kick Status')
        return
    }

    $newConfig = [pscustomobject]@{
        priorityRules = $priorityRules
        openKickOnClick = $cbOpenKick.Checked
        openBigscreenInChromeOnClick = $cbBigscreen.Checked
        refreshMinutes = $refreshMinutes
    }

    $newConfig | ConvertTo-Json -Depth 6 | Set-Content -Path $configPath -Encoding UTF8

    if ($cbStartup.Checked) {
        $wsh = New-Object -ComObject WScript.Shell
        $shortcut = $wsh.CreateShortcut($startupShortcut)
        $shortcut.TargetPath = $launcherPath
        $shortcut.WorkingDirectory = $PSScriptRoot
        $shortcut.Save()
    } elseif (Test-Path $startupShortcut) {
        Remove-Item $startupShortcut -Force
    }

    [System.Windows.Forms.MessageBox]::Show('Saved. Relaunch the status app to use the new channels.', 'Kick Status')
    $form.Close()
})

$cancelButton.Add_Click({
    $form.Close()
})

$form.Add_FormClosed({
    foreach ($entry in $ruleControls) {
        Dispose-RulePreview $entry
    }
})

[void]$form.ShowDialog()
