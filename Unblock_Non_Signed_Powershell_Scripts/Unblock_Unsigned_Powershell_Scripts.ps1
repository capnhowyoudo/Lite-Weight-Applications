<#
.SYNOPSIS
    Modern GUI (light/dark) to unblock .ps1 files, but only when they are
    NOT digitally signed.

.DESCRIPTION
    Run this script to open a window. Use "Add Files..." / "Add Folder..."
    to build a list of .ps1 files, then click "Unblock Unsigned" to process
    everything. Toggle the sun/moon button in the title bar to switch
    between dark and light mode.

    For each file:
      - Validly signed (Authenticode)   -> left alone, never touched.
      - Unsigned AND currently blocked  -> unblocked (Zone.Identifier removed).
      - Unsigned AND already unblocked  -> no action needed.
#>

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $psExe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $psExe -ArgumentList @(
        '-NoProfile','-STA','-ExecutionPolicy','Bypass','-File', "`"$PSCommandPath`""
    )
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------------------------------------------------------
# Themes
# ---------------------------------------------------------------------------
function C($r,$g,$b) { [System.Drawing.Color]::FromArgb(255,$r,$g,$b) }

$script:Themes = @{
    Dark = @{
        Bg          = C 24 26 32
        Panel       = C 32 34 42
        PanelAlt    = C 38 40 50
        TitleBar    = C 20 21 27
        Border      = C 55 58 70
        Text        = C 230 230 235
        SubText     = C 150 152 165
        Accent      = C 124 92 252
        AccentHover = C 145 116 253
        AccentText  = [System.Drawing.Color]::White
        Green       = C 88 214 141
        Magenta     = C 214 130 245
        Gray        = C 150 152 165
        Red         = C 235 110 110
        RedHoverText= [System.Drawing.Color]::White
        ToggleIcon  = ([char]0x2600)   # sun - shown while dark, click to go light
    }
    Light = @{
        Bg          = C 245 246 249
        Panel       = C 255 255 255
        PanelAlt    = C 235 236 241
        TitleBar    = C 255 255 255
        Border      = C 205 208 217
        Text        = C 30 32 38
        SubText     = C 105 108 122
        Accent      = C 108 74 240
        AccentHover = C 90 58 214
        AccentText  = [System.Drawing.Color]::White
        Green       = C 30 140 90
        Magenta     = C 160 50 190
        Gray        = C 105 108 122
        Red         = C 205 45 55
        RedHoverText= [System.Drawing.Color]::White
        ToggleIcon  = ([char]0x263E)   # crescent moon - shown while light, click to go dark
    }
}

$script:CurrentTheme = 'Dark'

$fontUI         = New-Object System.Drawing.Font("Segoe UI", 9.5)
$fontUIBold     = New-Object System.Drawing.Font("Segoe UI Semibold", 9.5)
$fontTitle      = New-Object System.Drawing.Font("Segoe UI Semibold", 10.5)
$fontBig        = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
$fontToggle     = New-Object System.Drawing.Font("Segoe UI Symbol", 12)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Get-Ps1FilesFromPath {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $item = Get-Item -LiteralPath $Path -Force
    if ($item.PSIsContainer) {
        return Get-ChildItem -LiteralPath $Path -Filter *.ps1 -Recurse -File -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty FullName
    } elseif ($item.Extension -ieq '.ps1') {
        return @($item.FullName)
    } else {
        return @()
    }
}

function Set-RoundedRegion {
    param($Control, [int]$Radius)
    $w = $Control.Width
    $h = $Control.Height
    $d = $Radius * 2
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc(0, 0, $d, $d, 180, 90)
    $path.AddArc(($w - $d), 0, $d, $d, 270, 90)
    $path.AddArc(($w - $d), ($h - $d), $d, $d, 0, 90)
    $path.AddArc(0, ($h - $d), $d, $d, 90, 90)
    $path.CloseFigure()
    $Control.Region = New-Object System.Drawing.Region($path)
}

function New-FlatButton {
    param([string]$Text, [System.Drawing.Font]$Font)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderSize = 0
    $b.Font = $Font
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    $b.UseVisualStyleBackColor = $false
    return $b
}

# All themed flat buttons get tracked here so Set-AppTheme can restyle them
# (base color + explicit hover/press colors, so .NET's own automatic
# hover-darkening never stacks with a manual color change).
$script:ThemedButtons = @()

# ---------------------------------------------------------------------------
# Form (borderless, custom title bar)
# ---------------------------------------------------------------------------
$form                = New-Object System.Windows.Forms.Form
$form.Text           = "PS1 Unblocker"
$form.Size           = New-Object System.Drawing.Size(800, 600)
$form.MinimumSize    = New-Object System.Drawing.Size(600, 420)
$form.StartPosition  = 'CenterScreen'
$form.FormBorderStyle = 'None'
$form.Padding        = New-Object System.Windows.Forms.Padding(1)

$form.Add_Load({ Set-RoundedRegion -Control $form -Radius 10 })
$form.Add_Resize({ Set-RoundedRegion -Control $form -Radius 10 })

# --- Custom title bar ---
$titleBar            = New-Object System.Windows.Forms.Panel
$titleBar.Dock        = 'Top'
$titleBar.Height       = 40

$iconLabel            = New-Object System.Windows.Forms.Label
$iconLabel.Text       = [System.Char]::ConvertFromUtf32(0x1F512)   # lock glyph (astral char needs a string)
$iconLabel.Font       = New-Object System.Drawing.Font("Segoe UI Emoji", 11)
$iconLabel.AutoSize   = $true
$iconLabel.Location   = New-Object System.Drawing.Point(14, 9)

$titleLabel           = New-Object System.Windows.Forms.Label
$titleLabel.Text      = "PS1 Unblocker"
$titleLabel.Font      = $fontTitle
$titleLabel.AutoSize  = $true
$titleLabel.Location  = New-Object System.Drawing.Point(38, 10)

$btnClose             = New-FlatButton -Text ([char]0x2715) -Font $fontUI
$btnClose.Size         = New-Object System.Drawing.Size(40, 40)
$btnClose.Dock         = 'Right'
$btnClose.Add_MouseEnter({ $btnClose.ForeColor = $script:Themes[$script:CurrentTheme].RedHoverText })
$btnClose.Add_MouseLeave({ $btnClose.ForeColor = $script:Themes[$script:CurrentTheme].SubText })
$btnClose.Add_Click({ $form.Close() })
$script:ThemedButtons += @{ Button = $btnClose; Base = 'TitleBar'; Hover = 'Red'; Text = 'SubText' }

$btnMin               = New-FlatButton -Text ([char]0x2212) -Font $fontUI
$btnMin.Size           = New-Object System.Drawing.Size(40, 40)
$btnMin.Dock           = 'Right'
$btnMin.Add_MouseEnter({ $btnMin.ForeColor = $script:Themes[$script:CurrentTheme].Text })
$btnMin.Add_MouseLeave({ $btnMin.ForeColor = $script:Themes[$script:CurrentTheme].SubText })
$btnMin.Add_Click({ $form.WindowState = 'Minimized' })
$script:ThemedButtons += @{ Button = $btnMin; Base = 'TitleBar'; Hover = 'PanelAlt'; Text = 'SubText' }

$btnTheme             = New-FlatButton -Text ($script:Themes[$script:CurrentTheme].ToggleIcon) -Font $fontToggle
$btnTheme.Size         = New-Object System.Drawing.Size(40, 40)
$btnTheme.Dock         = 'Right'
$btnTheme.Add_Click({
    $script:CurrentTheme = if ($script:CurrentTheme -eq 'Dark') { 'Light' } else { 'Dark' }
    Set-AppTheme -Name $script:CurrentTheme
})
$script:ThemedButtons += @{ Button = $btnTheme; Base = 'TitleBar'; Hover = 'PanelAlt'; Text = 'SubText' }

$titleBar.Controls.AddRange(@($iconLabel, $titleLabel, $btnTheme, $btnMin, $btnClose))

# Draggable title bar (move borderless form)
$script:dragging  = $false
$script:dragPoint = New-Object System.Drawing.Point(0, 0)
$titleBarDragHandler_Down = { $script:dragging = $true; $script:dragPoint = New-Object System.Drawing.Point($_.X, $_.Y) }
$titleBarDragHandler_Move = {
    if ($script:dragging) {
        $p = $form.PointToScreen($_.Location)
        $newX = $p.X - $script:dragPoint.X
        $newY = $p.Y - $script:dragPoint.Y
        $form.Location = New-Object System.Drawing.Point($newX, $newY)
    }
}
$titleBarDragHandler_Up = { $script:dragging = $false }
$titleBar.Add_MouseDown($titleBarDragHandler_Down)
$titleBar.Add_MouseMove($titleBarDragHandler_Move)
$titleBar.Add_MouseUp($titleBarDragHandler_Up)
$titleLabel.Add_MouseDown($titleBarDragHandler_Down)
$titleLabel.Add_MouseMove($titleBarDragHandler_Move)
$titleLabel.Add_MouseUp($titleBarDragHandler_Up)

# --- DataGridView (file list) ---
$grid                = New-Object System.Windows.Forms.DataGridView
$grid.Dock           = 'Fill'
$grid.BorderStyle    = 'None'
$grid.RowHeadersVisible = $false
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.ReadOnly       = $true
$grid.SelectionMode  = 'FullRowSelect'
$grid.MultiSelect    = $true
$grid.RowTemplate.Height = 30
$grid.ColumnHeadersHeightSizeMode = 'DisableResizing'
$grid.ColumnHeadersHeight = 34
$grid.EnableHeadersVisualStyles = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.CellBorderStyle = 'None'
$grid.RowHeadersWidthSizeMode = 'DisableResizing'

$colFile   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colFile.Name = "File"; $colFile.HeaderText = "File"; $colFile.FillWeight = 55
$colSigned = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSigned.Name = "Signed"; $colSigned.HeaderText = "SIGNED"; $colSigned.FillWeight = 15
$colBlocked = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colBlocked.Name = "Blocked"; $colBlocked.HeaderText = "BLOCKED"; $colBlocked.FillWeight = 15
$colAction = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colAction.Name = "Action"; $colAction.HeaderText = "STATUS"; $colAction.FillWeight = 20
$grid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colFile, $colSigned, $colBlocked, $colAction))
foreach ($c in @($colSigned, $colBlocked, $colAction)) { $c.SortMode = 'NotSortable' }

$gridWrapper          = New-Object System.Windows.Forms.Panel
$gridWrapper.Dock     = 'Fill'
$gridWrapper.Padding  = New-Object System.Windows.Forms.Padding(16, 0, 16, 0)
$gridWrapper.Controls.Add($grid)

# --- Bottom bar ---
$bottomBar            = New-Object System.Windows.Forms.Panel
$bottomBar.Dock        = 'Bottom'
$bottomBar.Height      = 70

$btnAddFiles  = New-FlatButton -Text "Add Files"  -Font $fontUI
$btnAddFiles.Size = New-Object System.Drawing.Size(110, 36)
$btnAddFiles.Location = New-Object System.Drawing.Point(16, 17)

$btnAddFolder = New-FlatButton -Text "Add Folder" -Font $fontUI
$btnAddFolder.Size = New-Object System.Drawing.Size(110, 36)
$btnAddFolder.Location = New-Object System.Drawing.Point(134, 17)

$btnClear     = New-FlatButton -Text "Clear"      -Font $fontUI
$btnClear.Size = New-Object System.Drawing.Size(80, 36)
$btnClear.Location = New-Object System.Drawing.Point(252, 17)

$script:ThemedButtons += @{ Button = $btnAddFiles;  Base = 'Panel'; Hover = 'PanelAlt'; Text = 'Text' }
$script:ThemedButtons += @{ Button = $btnAddFolder; Base = 'Panel'; Hover = 'PanelAlt'; Text = 'Text' }
$script:ThemedButtons += @{ Button = $btnClear;     Base = 'Panel'; Hover = 'PanelAlt'; Text = 'SubText' }

$statusLabel          = New-Object System.Windows.Forms.Label
$statusLabel.Text     = "0 file(s) in list"
$statusLabel.Font     = $fontUI
$statusLabel.AutoSize = $true
$statusLabel.Location = New-Object System.Drawing.Point(346, 27)

$btnProcess   = New-FlatButton -Text "Unblock Unsigned" -Font $fontUIBold
$btnProcess.Size = New-Object System.Drawing.Size(170, 40)
$btnProcess.Anchor = 'Right'
$script:ThemedButtons += @{ Button = $btnProcess; Base = 'Accent'; Hover = 'AccentHover'; Text = 'AccentText' }

$bottomBar.Controls.AddRange(@($btnAddFiles, $btnAddFolder, $btnClear, $statusLabel, $btnProcess))
$bottomBar.Add_Resize({
    $xProcess = $bottomBar.Width - $btnProcess.Width - 16
    $btnProcess.Location = New-Object System.Drawing.Point($xProcess, 15)
})

# Assemble (order matters for Dock)
$form.Controls.Add($gridWrapper)
$form.Controls.Add($bottomBar)
$form.Controls.Add($titleBar)

Set-RoundedRegion -Control $btnProcess -Radius 6
Set-RoundedRegion -Control $btnAddFiles -Radius 6
Set-RoundedRegion -Control $btnAddFolder -Radius 6
Set-RoundedRegion -Control $btnClear -Radius 6

# ---------------------------------------------------------------------------
# Theme application
# ---------------------------------------------------------------------------
function Set-AppTheme {
    param([string]$Name)
    $t = $script:Themes[$Name]

    $form.BackColor       = $t.Bg
    $titleBar.BackColor    = $t.TitleBar
    $iconLabel.ForeColor  = $t.Accent
    $titleLabel.ForeColor = $t.Text

    foreach ($entry in $script:ThemedButtons) {
        $b = $entry.Button
        $base = $t[$entry.Base]
        $hover = $t[$entry.Hover]
        $b.BackColor = $base
        $b.ForeColor = $t[$entry.Text]
        # Set hover/press colors explicitly so .NET's own automatic
        # hover-darkening never stacks with these values.
        $b.FlatAppearance.MouseOverBackColor = $hover
        $b.FlatAppearance.MouseDownBackColor = $hover
    }
    $btnTheme.Text = $t.ToggleIcon

    $grid.BackgroundColor = $t.Bg
    $grid.GridColor       = $t.Panel
    $grid.ColumnHeadersDefaultCellStyle.BackColor = $t.Panel
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = $t.SubText
    $grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $t.Panel
    $grid.DefaultCellStyle.BackColor = $t.Bg
    $grid.DefaultCellStyle.ForeColor = $t.Text
    $grid.DefaultCellStyle.SelectionBackColor = $t.PanelAlt
    $grid.DefaultCellStyle.SelectionForeColor = $t.Text
    $grid.AlternatingRowsDefaultCellStyle.BackColor = $t.Panel
    $grid.AlternatingRowsDefaultCellStyle.ForeColor = $t.Text

    $gridWrapper.BackColor = $t.Bg
    $bottomBar.BackColor   = $t.Bg

    $statusLabel.ForeColor = $t.SubText

    # Re-color existing rows to match the new theme
    foreach ($row in $grid.Rows) {
        $signedVal  = $row.Cells["Signed"].Value
        $blockedVal = $row.Cells["Blocked"].Value
        $actionVal  = $row.Cells["Action"].Value

        $row.Cells["Signed"].Style.ForeColor  = if ($signedVal -eq "Yes") { $t.Green } else { $t.Text }
        $row.Cells["Blocked"].Style.ForeColor = if ($blockedVal -eq "Yes") { $t.Magenta } else { $t.Text }

        $actionColor = switch -Wildcard ($actionVal) {
            "Skipped*"   { $t.Green }
            "Unblocked"  { $t.Magenta }
            "Already OK" { $t.Gray }
            "Error"      { $t.Red }
            default      { $t.SubText }
        }
        $row.Cells["Action"].Style.ForeColor = $actionColor
    }

    $form.Refresh()
}

Set-AppTheme -Name $script:CurrentTheme

# ---------------------------------------------------------------------------
# Data handling
# ---------------------------------------------------------------------------
$script:knownFiles = New-Object System.Collections.Generic.HashSet[string]

function Add-FileToGrid {
    param([string]$FullPath)
    if ($script:knownFiles.Contains($FullPath)) { return }
    $script:knownFiles.Add($FullPath) | Out-Null

    $sig = $null
    try { $sig = Get-AuthenticodeSignature -LiteralPath $FullPath -ErrorAction Stop } catch {}
    $isSigned  = ($sig -and $sig.Status -eq 'Valid')
    $isBlocked = $null -ne (Get-Item -LiteralPath $FullPath -Stream Zone.Identifier -ErrorAction SilentlyContinue)

    $rowIndex = $grid.Rows.Add($FullPath, $(if ($isSigned) {"Yes"} else {"No"}), $(if ($isBlocked) {"Yes"} else {"No"}), "Pending")
    $row = $grid.Rows[$rowIndex]
    $row.Tag = $FullPath

    $t = $script:Themes[$script:CurrentTheme]
    if ($isSigned) { $row.Cells["Signed"].Style.ForeColor = $t.Green }
    if ($isBlocked) { $row.Cells["Blocked"].Style.ForeColor = $t.Magenta }

    $statusLabel.Text = "$($grid.Rows.Count) file(s) in list"
}

function Add-PathsToGrid {
    param([string[]]$InputPaths)
    foreach ($p in $InputPaths) {
        foreach ($f in (Get-Ps1FilesFromPath -Path $p)) { Add-FileToGrid -FullPath $f }
    }
}

# Buttons
$btnAddFiles.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Select .ps1 file(s)"
    $dlg.Filter = "PowerShell Scripts (*.ps1)|*.ps1|All files (*.*)|*.*"
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Add-PathsToGrid -InputPaths $dlg.FileNames }
})

$btnAddFolder.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select a folder to scan for .ps1 files"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Add-PathsToGrid -InputPaths @($dlg.SelectedPath) }
})

$btnClear.Add_Click({
    $grid.Rows.Clear()
    $script:knownFiles.Clear()
    $statusLabel.Text = "0 file(s) in list"
})

$btnProcess.Add_Click({
    if ($grid.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Add some .ps1 files first.", "Nothing to do") | Out-Null
        return
    }
    $t = $script:Themes[$script:CurrentTheme]
    $unblocked = 0; $skippedSigned = 0; $alreadyOk = 0; $errors = 0
    foreach ($row in $grid.Rows) {
        $path = $row.Tag
        try {
            $sig = Get-AuthenticodeSignature -LiteralPath $path -ErrorAction Stop
            $isSigned  = ($sig.Status -eq 'Valid')
            $isBlocked = $null -ne (Get-Item -LiteralPath $path -Stream Zone.Identifier -ErrorAction SilentlyContinue)

            if ($isSigned) {
                $row.Cells["Action"].Value = "Skipped (signed)"
                $row.Cells["Action"].Style.ForeColor = $t.Green
                $skippedSigned++
            } elseif ($isBlocked) {
                Unblock-File -LiteralPath $path -ErrorAction Stop
                $row.Cells["Blocked"].Value = "No"
                $row.Cells["Action"].Value = "Unblocked"
                $row.Cells["Action"].Style.ForeColor = $t.Magenta
                $unblocked++
            } else {
                $row.Cells["Action"].Value = "Already OK"
                $row.Cells["Action"].Style.ForeColor = $t.Gray
                $alreadyOk++
            }
        } catch {
            $row.Cells["Action"].Value = "Error"
            $row.Cells["Action"].Style.ForeColor = $t.Red
            $errors++
        }
    }
    $summary = "Unblocked: $unblocked`nSkipped (signed): $skippedSigned`nAlready OK: $alreadyOk`nErrors: $errors"
    [System.Windows.Forms.MessageBox]::Show($summary, "Done") | Out-Null
})

# ---------------------------------------------------------------------------
# Show
# ---------------------------------------------------------------------------
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.Add_Shown({ $form.Activate() })
[System.Windows.Forms.Application]::Run($form)
