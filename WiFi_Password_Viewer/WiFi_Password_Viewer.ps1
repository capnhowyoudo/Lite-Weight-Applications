<#
.SYNOPSIS
    GUI tool to display saved Wi-Fi profiles and their passwords.

.DESCRIPTION
    This PowerShell script creates a WPF-based GUI that retrieves and displays
    all saved Wi-Fi profiles from the system using the Windows `netsh wlan`
    command. When the user clicks the button, the script:

        • Enumerates all stored Wi-Fi profiles
        • Extracts passwords in clear text (if available)
        • Displays results in a scrollable text output window
        • Automatically elevates to Administrator if required

    Netsh requires elevated permissions to reveal Wi-Fi passwords, so the script
    includes an auto-elevation mechanism that relaunches the script as admin.

.NOTES
    Author: capnhowyoudo
    Version: 1.0
    Requirements:
        • Must run as Administrator (auto-elevation included)
        • Uses WPF (PresentationFramework) for GUI
        • Windows 10/11 compatible

    Additional Notes:
        • Profiles without passwords (e.g., enterprise EAP) will show
          "[Not Set / Hidden]".
        • GUI output uses a dark theme with console-style formatting.
        • Uses only built-in Windows binaries (no external modules required).
#>


#----------------------------------------
# Auto-elevate if not running as admin
#----------------------------------------
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

#----------------------------------------
# XAML GUI
#----------------------------------------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Wi-Fi Password Viewer" Height="400" Width="700"
        WindowStartupLocation="CenterScreen" FontFamily="Consolas" FontSize="14" Background="#222">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Button Name="btnGetPasswords" Content="Get Wi-Fi Passwords" Width="200" Height="35" Margin="0,0,0,10" HorizontalAlignment="Left"/>

        <TextBox Name="txtOutput" Grid.Row="1" AcceptsReturn="True" TextWrapping="Wrap"
                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                 IsReadOnly="True" Background="Black" Foreground="LightGreen" FontSize="14"/>
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
$btnGetPasswords = $window.FindName("btnGetPasswords")
$txtOutput = $window.FindName("txtOutput")

#----------------------------------------
# Button Click - Get Wi-Fi Passwords
#----------------------------------------
$btnGetPasswords.Add_Click({
    $txtOutput.Clear()
    $txtOutput.AppendText("&#128246; Retrieving saved Wi-Fi profiles and passwords...`r`n`r`n")

    # Get all Wi-Fi profiles
    $profiles = netsh wlan show profiles | Select-String "All User Profile" | ForEach-Object { ($_ -split ":")[1].Trim() }

    if ($profiles.Count -eq 0) {
        $txtOutput.AppendText("No Wi-Fi profiles found.`r`n")
        return
    }

    foreach ($profile in $profiles) {
        $txtOutput.AppendText("→ $profile`r`n")
        $passwordInfo = netsh wlan show profile name="$profile" key=clear | Select-String "Key Content"
        if ($passwordInfo) {
            $password = ($passwordInfo -split ":")[1].Trim()
            $txtOutput.AppendText("   Password: $password`r`n`r`n")
        } else {
            $txtOutput.AppendText("   Password: [Not Set / Hidden]`r`n`r`n")
        }
    }

    $txtOutput.AppendText("✅ Done.`r`n")
})

#----------------------------------------
# Show Window
#----------------------------------------
$window.ShowDialog() | Out-Null
