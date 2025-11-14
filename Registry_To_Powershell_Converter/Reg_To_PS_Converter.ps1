<#
.SYNOPSIS
    Converts Windows Registry (.reg) entries into PowerShell commands using a WPF GUI.

.DESCRIPTION
    This tool provides a graphical interface for converting .reg file content into 
    PowerShell commands (New-Item / New-ItemProperty), including support for:
        • String values
        • DWORD values
        • Binary values
    Additional features include:
        • Loading .reg files
        • Syntax validation using PowerShell parser
        • Registry path existence testing
        • Light/Dark theme toggle
        • Saving converted output as a .ps1 script

.NOTES
    Author: capnhowyoudo
    Filename: Reg_To_PS_Converter.ps1
    Requirements:
        - Windows PowerShell (supports WPF)
        - ExecutionPolicy allowing script execution
    Run Example:
        powershell.exe -ExecutionPolicy Bypass -File .\RegToPS_WPF_Theme.ps1
#>

# Registry → PowerShell Converter (WPF GUI with Dark/Light Theme Toggle)
# Inspired by Reg2PS | Includes Syntax Check + Test-Path + Theme Switch
# Run: powershell -ExecutionPolicy Bypass -File .\RegToPS_WPF_Theme.ps1

Add-Type -AssemblyName PresentationFramework

#----------------------------------------
# Conversion Function
#----------------------------------------
function Convert-RegToPS {
    param($regText)
    $sb = New-Object System.Text.StringBuilder
    $lines = $regText -split "`r?`n"
    $currentKey = ""

    foreach ($line in $lines) {
        $trim = $line.Trim()
        if ($trim -match "^\[(.+)\]$") {
            $currentKey = $matches[1]
            $null = $sb.AppendLine("New-Item -Path 'Registry::$currentKey' -Force | Out-Null")
        }
        elseif ($trim -match '^(\"(.+?)\")=(.+)$') {
            $name = $matches[2]
            $valuePart = $matches[3]
            if ($valuePart -match '^dword:(.+)$') {
                $val = [convert]::ToInt32($matches[1],16)
                $null = $sb.AppendLine("New-ItemProperty -Path 'Registry::$currentKey' -Name '$name' -Value $val -PropertyType DWord -Force")
            }
            elseif ($valuePart -match '^hex:(.+)$') {
                $bytes = $matches[1] -split ','
                $byteArray = ($bytes | ForEach-Object { "0x$_" }) -join ','
                $null = $sb.AppendLine("New-ItemProperty -Path 'Registry::$currentKey' -Name '$name' -Value ([byte[]]@($byteArray)) -PropertyType Binary -Force")
            }
            elseif ($valuePart -match '^\"(.*)\"$') {
                $stringVal = $matches[1].Replace("'", "''")
                $null = $sb.AppendLine("New-ItemProperty -Path 'Registry::$currentKey' -Name '$name' -Value '$stringVal' -PropertyType String -Force")
            }
        }
    }
    return $sb.ToString()
}

#----------------------------------------
# Syntax Check
#----------------------------------------
function Check-PSSyntax {
    param([string]$scriptText)
    try {
        [void][System.Management.Automation.PSParser]::Tokenize($scriptText, [ref]$null)
        [System.Windows.MessageBox]::Show("✅ Syntax check passed!", "Syntax Check", 'OK', 'Information')
    }
    catch {
        [System.Windows.MessageBox]::Show("❌ Syntax error:`n$($_.Exception.Message)", "Syntax Check", 'OK', 'Error')
    }
}

#----------------------------------------
# Test Registry Path
#----------------------------------------
function Test-RegistryPath {
    param([string]$regPath)
    try {
        if (Test-Path "Registry::$regPath") {
            [System.Windows.MessageBox]::Show("✅ Path exists:`n$regPath", "Registry Test", 'OK', 'Information')
        } else {
            [System.Windows.MessageBox]::Show("⚠️ Path does not exist:`n$regPath", "Registry Test", 'OK', 'Warning')
        }
    }
    catch {
        [System.Windows.MessageBox]::Show("❌ Error:`n$($_.Exception.Message)", "Registry Test", 'OK', 'Error')
    }
}

#----------------------------------------
# XAML UI Definition
#----------------------------------------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Registry → PowerShell Converter"
        Height="720" Width="1020"
        WindowStartupLocation="CenterScreen"
        FontFamily="Consolas" >
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Orientation="Horizontal" Grid.Row="0" Margin="0,0,0,5">
            <Button Name="btnLoad" Width="90" Margin="5" Content="Load .reg"/>
            <Button Name="btnConvert" Width="90" Margin="5" Content="Convert"/>
            <Button Name="btnCheck" Width="100" Margin="5" Content="Check Syntax"/>
            <Button Name="btnTest" Width="90" Margin="5" Content="Test Path"/>
            <Button Name="btnSave" Width="90" Margin="5" Content="Save .ps1"/>
            <Button Name="btnTheme" Width="130" Margin="5" Content="&#127769; Toggle Theme"/>
        </StackPanel>

        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <TextBox Name="txtReg" Grid.Column="0" AcceptsReturn="True" TextWrapping="Wrap"
                     VerticalScrollBarVisibility="Auto" Margin="5"/>
            <TextBox Name="txtPS" Grid.Column="1" AcceptsReturn="True" TextWrapping="Wrap"
                     VerticalScrollBarVisibility="Auto" Margin="5" IsReadOnly="True"/>
        </Grid>

        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="5">
            <TextBox Name="txtPath" Width="600" Height="25" Margin="5"
                     Text="HKEY_LOCAL_MACHINE\SOFTWARE\"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

#----------------------------------------
# Controls
#----------------------------------------
$btnLoad  = $window.FindName("btnLoad")
$btnConvert = $window.FindName("btnConvert")
$btnCheck = $window.FindName("btnCheck")
$btnTest  = $window.FindName("btnTest")
$btnSave  = $window.FindName("btnSave")
$btnTheme = $window.FindName("btnTheme")
$txtReg   = $window.FindName("txtReg")
$txtPS    = $window.FindName("txtPS")
$txtPath  = $window.FindName("txtPath")

#----------------------------------------
# Theme Switcher
#----------------------------------------
$global:isDark = $true

function Apply-Theme {
    param([bool]$dark)
    if ($dark) {
        $window.Background = "#1E1E1E"
        $txtReg.Background = "#252526"; $txtReg.Foreground = "White"
        $txtPS.Background = "#252526";  $txtPS.Foreground = "LightGreen"
        $txtPath.Background = "#333333"; $txtPath.Foreground = "White"
        foreach ($btn in @($btnLoad,$btnConvert,$btnCheck,$btnTest,$btnSave,$btnTheme)) {
            $btn.Background = "#3A3D41"; $btn.Foreground = "White"
        }
        $btnTheme.Content = "&#127774; Light Mode"
    } else {
        $window.Background = "#F0F0F0"
        $txtReg.Background = "White"; $txtReg.Foreground = "Black"
        $txtPS.Background = "White";  $txtPS.Foreground = "DarkGreen"
        $txtPath.Background = "White"; $txtPath.Foreground = "Black"
        foreach ($btn in @($btnLoad,$btnConvert,$btnCheck,$btnTest,$btnSave,$btnTheme)) {
            $btn.Background = "#E0E0E0"; $btn.Foreground = "Black"
        }
        $btnTheme.Content = "&#127769; Dark Mode"
    }
}

Apply-Theme $true  # default dark

$btnTheme.Add_Click({
    $global:isDark = -not $global:isDark
    Apply-Theme $global:isDark
})

#----------------------------------------
# Button Actions
#----------------------------------------
$btnLoad.Add_Click({
    $ofd = New-Object Microsoft.Win32.OpenFileDialog
    $ofd.Filter = "Registry Files (*.reg)|*.reg"
    if ($ofd.ShowDialog()) {
        $txtReg.Text = Get-Content $ofd.FileName -Raw
    }
})

$btnConvert.Add_Click({
    $txtPS.Text = Convert-RegToPS $txtReg.Text
})

$btnCheck.Add_Click({
    Check-PSSyntax $txtPS.Text
})

$btnTest.Add_Click({
    Test-RegistryPath $txtPath.Text
})

$btnSave.Add_Click({
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = "PowerShell Script (*.ps1)|*.ps1"
    if ($sfd.ShowDialog()) {
        $txtPS.Text | Out-File -Encoding UTF8 $sfd.FileName
        [System.Windows.MessageBox]::Show("Saved: $($sfd.FileName)", "Saved", 'OK', 'Information')
    }
})

#----------------------------------------
# Run Window
#----------------------------------------
$window.ShowDialog() | Out-Null
