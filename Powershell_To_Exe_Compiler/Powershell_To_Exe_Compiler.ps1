<#
.SYNOPSIS
    A full GUI-based PowerShell-to-EXE compiler using PS2EXE, offering
    metadata configuration, architecture selection, icon embedding,
    logging, presets, and automatic ExecutionPolicy repair.

.DESCRIPTION
    This script provides a Windows Presentation Foundation (WPF) graphical
    interface that allows users to compile a PowerShell (.ps1) script into
    an executable (.exe) using the PS2EXE PowerShell module.

    Key features:
      • GUI file pickers for input PS1, output EXE, and icon files.
      • Metadata configuration (title, description, company, version).
      • Compiler flags: Hide Console, No Error, No Output, Require Admin.
      • Architecture selection between x64 and x86.
      • Persistent user presets via JSON.
      • Log output, progress bar, and “Show Command” generator.
      • Automatic PS2EXE loader with ExecutionPolicy self-repair.
      • Auto-elevation to Administrator when required.

.NOTES
    Author: capnhowyoudo
    Version: 1.0
    Requirements:
        - PowerShell 5+
        - PS2EXE module
        - Windows Presentation Framework (WPF)
    Tested on Windows 10/11.

#>

Add-Type -AssemblyName PresentationFramework

# ===================================================
# ⚙️ PS2EXE Compiler Pro GUI (Auto Execution Policy Fix)
# ===================================================

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="PS2EXE Compiler Pro"
        Height="700" Width="850"
        WindowStartupLocation="CenterScreen"
        Background="#F4F6F9"
        ResizeMode="CanResize">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Background="#2F3640" CornerRadius="10" Padding="10" Margin="0,0,0,15">
            <TextBlock Text="⚙️ PowerShell → EXE Compiler Pro"
                       FontSize="22" FontWeight="Bold" Foreground="White"/>
        </Border>

        <!-- Main Content -->
        <StackPanel Grid.Row="1" Orientation="Vertical" Margin="0,0,0,10">

            <!-- File Section -->
            <TextBlock Text="Files" FontSize="16" FontWeight="Bold" Margin="0,0,0,5"/>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                <TextBlock Text="Input .PS1:" Width="120" VerticalAlignment="Center"/>
                <TextBox Name="InputPath" Width="540" Margin="0,0,6,0"/>
                <Button Name="BrowseInput" Content="Browse" Width="90" Background="#40739e" Foreground="White"/>
            </StackPanel>

            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                <TextBlock Text="Output .EXE:" Width="120" VerticalAlignment="Center"/>
                <TextBox Name="OutputPath" Width="540" Margin="0,0,6,0"/>
                <Button Name="BrowseOutput" Content="Browse" Width="90" Background="#40739e" Foreground="White"/>
            </StackPanel>

            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBlock Text="Icon (.ICO):" Width="120" VerticalAlignment="Center"/>
                <TextBox Name="IconPath" Width="540" Margin="0,0,6,0"/>
                <Button Name="BrowseIcon" Content="Browse" Width="90" Background="#40739e" Foreground="White"/>
            </StackPanel>

            <!-- Metadata -->
            <TextBlock Text="Metadata" FontSize="16" FontWeight="Bold" Margin="0,10,0,5"/>
            <StackPanel Orientation="Vertical" Margin="0,0,0,10">
                <StackPanel Orientation="Horizontal" Margin="0,0,0,5">
                    <TextBlock Text="Product:" Width="120" VerticalAlignment="Center"/>
                    <TextBox Name="ProductName" Width="640"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,5">
                    <TextBlock Text="Description:" Width="120" VerticalAlignment="Center"/>
                    <TextBox Name="Description" Width="640"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,5">
                    <TextBlock Text="Company:" Width="120" VerticalAlignment="Center"/>
                    <TextBox Name="Company" Width="640"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Version:" Width="120" VerticalAlignment="Center"/>
                    <TextBox Name="Version" Width="640" Text="1.0.0.0"/>
                </StackPanel>
            </StackPanel>

            <!-- Options -->
            <TextBlock Text="Options" FontSize="16" FontWeight="Bold" Margin="0,10,0,5"/>
            <WrapPanel>
                <CheckBox Name="HideConsole" Content="Hide Console" IsChecked="True" Margin="0,0,10,5"/>
                <CheckBox Name="NoError" Content="Suppress Errors" Margin="0,0,10,5"/>
                <CheckBox Name="NoOutput" Content="No Output" Margin="0,0,10,5"/>
                <CheckBox Name="RequireAdmin" Content="Require Admin" Margin="0,0,10,5"/>
            </WrapPanel>
            <WrapPanel Margin="0,5,0,15">
                <RadioButton Name="x64" Content="x64" GroupName="arch" IsChecked="True" Margin="0,0,10,0"/>
                <RadioButton Name="x86" Content="x86" GroupName="arch" Margin="0,0,10,0"/>
            </WrapPanel>

            <!-- Progress and Log -->
            <ProgressBar Name="ProgressBar" Height="20" Minimum="0" Maximum="100" Visibility="Collapsed" Margin="0,0,0,5"/>
            <TextBlock Name="StatusText" Text="Ready." FontWeight="Bold" Foreground="#2f3640" Margin="0,0,0,5"/>
            <Border BorderBrush="#dcdde1" BorderThickness="1" CornerRadius="6" Background="White" Height="180">
                <ScrollViewer Name="LogScroll" VerticalScrollBarVisibility="Auto">
                    <TextBox Name="LogOutput" Background="White" BorderThickness="0" FontFamily="Consolas" 
                             FontSize="13" IsReadOnly="True" TextWrapping="Wrap"/>
                </ScrollViewer>
            </Border>

            <!-- Presets -->
            <StackPanel Orientation="Horizontal" Margin="0,15,0,0">
                <Button Name="SavePreset" Content="&#128190; Save Settings" Width="150" Margin="0,0,10,0" Background="#8c7ae6" Foreground="White"/>
                <Button Name="LoadPreset" Content="&#128194; Load Settings" Width="150" Background="#718093" Foreground="White"/>
                <Button Name="ShowCmd" Content="&#129534; Show Command" Width="150" Margin="10,0,0,0" Background="#0097e6" Foreground="White"/>
                <Button Name="SaveLog" Content="&#128221; Save Log" Width="120" Margin="10,0,0,0" Background="#44bd32" Foreground="White"/>
            </StackPanel>
        </StackPanel>

        <!-- Footer -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,15,0,0">
            <Button Name="CompileButton" Content="&#128640; Compile EXE" Width="200" Height="38" Background="#44bd32" Foreground="White" FontWeight="Bold" Margin="0,0,12,0"/>
            <Button Name="ExitButton" Content="❌ Exit" Width="100" Height="38" Background="#e84118" Foreground="White"/>
        </StackPanel>
    </Grid>
</Window>
"@

# ------------------ Load GUI ------------------
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
$xaml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $window.FindName($_.Name) }

# ------------------ Logging ------------------
function Log($msg) {
    $LogOutput.AppendText("$(Get-Date -Format 'HH:mm:ss')  $msg`r`n")
    $LogScroll.ScrollToEnd()
}

# ------------------ Execution Policy Safe Load ------------------
function Check-PS2EXE {
    Log "&#128269; Checking PS2EXE module..."
    try {
        Import-Module ps2exe -ErrorAction Stop
        Log "✅ PS2EXE module loaded successfully."
        return $true
    } catch {
        Log "⚠️ PS2EXE not loaded. Attempting install/load with ExecutionPolicy bypass..."

        try {
            Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
            Install-Module ps2exe -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Import-Module ps2exe -ErrorAction Stop
            Log "✅ PS2EXE installed and loaded successfully under temporary bypass."
            return $true
        } catch {
            Log "⚠️ Temporary bypass failed. Attempting admin elevation..."
            $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            if (-not $isAdmin) {
                Log "&#128272; Relaunching PowerShell as Administrator..."
                $scriptPath = $MyInvocation.MyCommand.Definition
                Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`""
                [System.Windows.MessageBox]::Show("Restarting PowerShell as Administrator to fix Execution Policy.","Restarting")
                exit
            }
            try {
                Set-ExecutionPolicy RemoteSigned -Scope LocalMachine -Force
                Import-Module ps2exe -ErrorAction Stop
                Log "✅ Execution Policy fixed and PS2EXE loaded!"
                return $true
            } catch {
                Log "❌ Failed to fix Execution Policy: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Please run PowerShell as Administrator and execute:`nSet-ExecutionPolicy RemoteSigned -Scope LocalMachine -Force","Execution Policy Error")
                return $false
            }
        }
    }
}

# ------------------ Other Functions ------------------
function Build-Command {
    $cmd = "ps2exe -inputFile `"$($InputPath.Text)`" -outputFile `"$($OutputPath.Text)`""
    if ($IconPath.Text) { $cmd += " -iconFile `"$($IconPath.Text)`"" }
    if ($HideConsole.IsChecked) { $cmd += " -noConsole" }
    if ($NoError.IsChecked) { $cmd += " -noError" }
    if ($NoOutput.IsChecked) { $cmd += " -noOutput" }
    if ($RequireAdmin.IsChecked) { $cmd += " -requireAdmin" }
    if ($x64.IsChecked) { $cmd += " -x64" } else { $cmd += " -x86" }
    if ($ProductName.Text) { $cmd += " -title `"$($ProductName.Text)`"" }
    if ($Description.Text) { $cmd += " -description `"$($Description.Text)`"" }
    if ($Company.Text) { $cmd += " -copyright `"$($Company.Text)`"" }
    if ($Version.Text) { $cmd += " -version `"$($Version.Text)`"" }
    return $cmd
}

function Save-Settings {
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = "Settings (*.json)|*.json"
    if ($dlg.ShowDialog()) {
        $settings = @{
            InputPath = $InputPath.Text
            OutputPath = $OutputPath.Text
            IconPath = $IconPath.Text
            ProductName = $ProductName.Text
            Description = $Description.Text
            Company = $Company.Text
            Version = $Version.Text
            HideConsole = $HideConsole.IsChecked
            NoError = $NoError.IsChecked
            NoOutput = $NoOutput.IsChecked
            RequireAdmin = $RequireAdmin.IsChecked
            Arch = if ($x64.IsChecked) { "x64" } else { "x86" }
        }
        $settings | ConvertTo-Json | Set-Content $dlg.FileName
        Log "&#128190; Settings saved to $($dlg.FileName)"
    }
}

function Load-Settings {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = "Settings (*.json)|*.json"
    if ($dlg.ShowDialog()) {
        $data = Get-Content $dlg.FileName | ConvertFrom-Json
        foreach ($key in $data.PSObject.Properties.Name) {
            if (Get-Variable -Name $key -ErrorAction SilentlyContinue) {
                (Get-Variable -Name $key).Value.Text = $data.$key
            }
        }
        Log "&#128194; Settings loaded from $($dlg.FileName)"
    }
}

function Save-Log {
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = "Text Files (*.txt)|*.txt"
    if ($dlg.ShowDialog()) {
        $LogOutput.Text | Set-Content $dlg.FileName
        Log "&#128221; Log saved to $($dlg.FileName)"
    }
}

# ------------------ Button Events ------------------
$BrowseInput.Add_Click({ $dlg = New-Object Microsoft.Win32.OpenFileDialog; $dlg.Filter = "PowerShell Scripts (*.ps1)|*.ps1"; if ($dlg.ShowDialog()) { $InputPath.Text = $dlg.FileName } })
$BrowseOutput.Add_Click({ $dlg = New-Object Microsoft.Win32.SaveFileDialog; $dlg.Filter = "Executable (*.exe)|*.exe"; if ($dlg.ShowDialog()) { $OutputPath.Text = $dlg.FileName } })
$BrowseIcon.Add_Click({ $dlg = New-Object Microsoft.Win32.OpenFileDialog; $dlg.Filter = "Icon Files (*.ico)|*.ico"; if ($dlg.ShowDialog()) { $IconPath.Text = $dlg.FileName } })
$SavePreset.Add_Click({ Save-Settings })
$LoadPreset.Add_Click({ Load-Settings })
$SaveLog.Add_Click({ Save-Log })
$ShowCmd.Add_Click({ [System.Windows.MessageBox]::Show((Build-Command), "Generated PS2EXE Command") })
$ExitButton.Add_Click({ $window.Close() })

$CompileButton.Add_Click({
    if (-not (Check-PS2EXE)) { return }
    Log "&#128640; Compiling..."
    $ProgressBar.Visibility = "Visible"
    $ProgressBar.Value = 40
    try {
        Invoke-Expression (Build-Command) | Out-Null
        $ProgressBar.Value = 100
        Log "✅ Compilation complete!"
        [System.Windows.MessageBox]::Show("EXE successfully created!","Done")
    } catch {
        Log "❌ Error: $($_.Exception.Message)"
    } finally {
        $ProgressBar.Visibility = "Collapsed"
    }
})

# ------------------ Run ------------------
$window.ShowDialog() | Out-Null
