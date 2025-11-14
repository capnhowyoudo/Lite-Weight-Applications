<#
.SYNOPSIS
    GUI tool to export and import Wi-Fi profiles using PowerShell and Netsh.

.DESCRIPTION
    This script launches a WPF-based graphical interface that allows users to:
      • Export all Wi-Fi profiles to a chosen folder (XML format)
      • Import Wi-Fi profiles from a selected folder
      • View a scrollable output log
      • Browse for folders using a Windows Folder Browser dialog

    Features:
      • Auto-elevation to Administrator (required for Wi-Fi profile operations)
      • Resizable GUI
      • Dark-themed interface
      • Safe path validation before actions

.NOTES
    REQUIREMENTS:
      • Must run as Administrator (script auto-elevates)
      • Requires Windows 10 or Windows 11 with WLAN AutoConfig enabled

    USAGE:
      • Run normally: 
            powershell.exe -ExecutionPolicy Bypass -File .\WiFi_Profile_Export_Import_GUI.ps1
      • Export:
            Select a folder → Click Export
      • Import:
            Select a folder containing XML files → Click Import

    FOLDER CONTENT EXPECTATION:
      • Exported profiles are saved as: Wi-Fi-<profilename>.xml
      • Those XML files can be imported on another computer using the Import section.

    AUTHOR:
      capnhowyoudo
#>

#----------------------------------------
# Auto-elevate if not running as admin
#----------------------------------------
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {

    Start-Process powershell `
        "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
        -Verb RunAs
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

#----------------------------------------
# XAML GUI (Resizable & Responsive)
#----------------------------------------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Wi-Fi Profile Export / Import Utility"
        Height="400" Width="650" WindowStartupLocation="CenterScreen"
        FontFamily="Consolas" FontSize="13" Background="#222"
        ResizeMode="CanResize">
    <Grid Margin="8">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Export Section -->
        <GroupBox Header="Export Wi-Fi Profiles" Grid.Row="0" Margin="3" Foreground="White" Background="#333">
            <Grid Margin="3">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="70"/>
                    <ColumnDefinition Width="70"/>
                </Grid.ColumnDefinitions>
                <TextBox Name="txtExportPath" Grid.Column="0" Height="28" Margin="0,0,3,0"/>
                <Button Name="btnExportBrowse" Content="Browse" Grid.Column="1" Width="70" Height="28" Margin="0,0,3,0"/>
                <Button Name="btnExport" Content="Export" Grid.Column="2" Width="70" Height="28" Background="#2E8B57" Foreground="White"/>
            </Grid>
        </GroupBox>

        <!-- Import Section -->
        <GroupBox Header="Import Wi-Fi Profiles" Grid.Row="1" Margin="3" Foreground="White" Background="#333">
            <Grid Margin="3">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="70"/>
                    <ColumnDefinition Width="70"/>
                </Grid.ColumnDefinitions>
                <TextBox Name="txtImportPath" Grid.Column="0" Height="28" Margin="0,0,3,0"/>
                <Button Name="btnImportBrowse" Content="Browse" Grid.Column="1" Width="70" Height="28" Margin="0,0,3,0"/>
                <Button Name="btnImport" Content="Import" Grid.Column="2" Width="70" Height="28" Background="#4682B4" Foreground="White"/>
            </Grid>
        </GroupBox>

        <!-- Output Box -->
        <GroupBox Header="Output Log" Grid.Row="2" Margin="3" Foreground="White" Background="#333">
            <TextBox Name="txtOutput" AcceptsReturn="True" TextWrapping="Wrap"
                     VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                     IsReadOnly="True" Background="Black" Foreground="LightGreen"
                     FontSize="12" FontFamily="Consolas" Margin="3"/>
        </GroupBox>
    </Grid>
</Window>
"@

#----------------------------------------
# Load XAML
#----------------------------------------
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

#----------------------------------------
# Controls
#----------------------------------------
$txtExportPath = $window.FindName("txtExportPath")
$txtImportPath = $window.FindName("txtImportPath")
$btnExportBrowse = $window.FindName("btnExportBrowse")
$btnExport = $window.FindName("btnExport")
$btnImportBrowse = $window.FindName("btnImportBrowse")
$btnImport = $window.FindName("btnImport")
$txtOutput = $window.FindName("txtOutput")

#----------------------------------------
# Folder Browsing
#----------------------------------------
$btnExportBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq "OK") { $txtExportPath.Text = $dlg.SelectedPath }
})

$btnImportBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq "OK") { $txtImportPath.Text = $dlg.SelectedPath }
})

#----------------------------------------
# Export Wi-Fi Profiles
#----------------------------------------
$btnExport.Add_Click({
    $exportPath = $txtExportPath.Text
    if (-not (Test-Path $exportPath)) {
        [System.Windows.MessageBox]::Show("⚠ Invalid export path.", "Error", "OK", "Error")
        return
    }

    $txtOutput.Clear()
    $txtOutput.AppendText("📤 Exporting Wi-Fi profiles...`r`n`r`n")

    $profiles = netsh wlan show profiles |
        Select-String "All User Profile" |
        ForEach-Object { ($_ -split ":")[1].Trim() }

    if ($profiles.Count -eq 0) {
        $txtOutput.AppendText("No Wi-Fi profiles found.`r`n")
        return
    }

    foreach ($profile in $profiles) {
        $txtOutput.AppendText("→ Exporting: $profile`r`n")
        netsh wlan export profile name="$profile" folder="$exportPath" key=clear | Out-Null
    }

    $txtOutput.AppendText("`r`n✅ All profiles exported to:`r`n$exportPath`r`n")
})

#----------------------------------------
# Import Wi-Fi Profiles
#----------------------------------------
$btnImport.Add_Click({
    $importPath = $txtImportPath.Text
    if (-not (Test-Path $importPath)) {
        [System.Windows.MessageBox]::Show("⚠ Invalid import path.", "Error", "OK", "Error")
        return
    }

    $txtOutput.Clear()
    $txtOutput.AppendText("📥 Importing Wi-Fi profiles...`r`n`r`n")

    $profiles = Get-ChildItem -Path $importPath -Filter "*.xml"
    if ($profiles.Count -eq 0) {
        $txtOutput.AppendText("No Wi-Fi profile XML files found in:`r`n$importPath`r`n")
        return
    }

    foreach ($profile in $profiles) {
        $txtOutput.AppendText("→ Importing: $($profile.Name)`r`n")
        netsh wlan add profile filename="$($profile.FullName)" | Out-Null
    }

    $txtOutput.AppendText("`r`n✅ All profiles imported successfully.`r`n")
})

#----------------------------------------
# Show Window
#----------------------------------------
$window.ShowDialog() | Out-Null
