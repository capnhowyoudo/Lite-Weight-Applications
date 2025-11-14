<#
.SYNOPSIS
    A full-featured graphical Wi-Fi network scanner that displays nearby
    wireless networks, connected network info, vendors, RSSI, band,
    protocol, and exports results to CSV.

.DESCRIPTION
    This script provides a Windows Presentation Foundation (WPF) GUI for
    scanning and analyzing nearby Wi-Fi access points using native
    Windows netsh commands.

    Features:
        • Real-time Wi-Fi network scanning (SSID, BSSID, vendor lookup)
        • Signal strength visualization with colored bar indicators
        • Live vendor lookup from macvendors.com with local caching
        • Connected network details panel
        • Automatic refresh (1–30 seconds adjustable)
        • Export scan results to CSV
        • Responsive GUI built in XAML

    Built-in Logic:
        • MAC normalization helper
        • Vendor lookup with caching and fallback OUI table
        • RSSI estimation and channel/band detection
        • DispatcherTimer for auto-refresh
        • Attractive DataGrid layout with dynamic color triggers

.NOTES
    Author: capnhowyoudo
    Version: 1.0
    Requirements:
        • Windows OS with Wi-Fi adapter
        • PowerShell 5+
        • WPF assemblies (PresentationFramework, PresentationCore, WindowsBase)

    Notes:
        • The script uses netsh commands — these do NOT require admin rights.
        • Vendor lookup uses macvendors.com API and will fallback to local vendor data if offline.
        • Auto-refresh will continually rescan using the chosen interval.
        • CSV export saves exactly what is displayed in the grid.
#>


Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# === Global Cache for Vendors ===
$Global:VendorCache = @{}

# === Helper Functions ===
function Normalize-Mac {
    param([string]$mac)
    if (-not $mac) { return $null }
    $m = $mac.Trim() -replace '[-\.]', ':'
    $m = $m -replace '([0-9A-Fa-f]{2})(?=[0-9A-Fa-f])', '$1:'
    $m = $m.TrimEnd(':').ToLower()
    return $m
}

function Get-Vendor {
    param([string]$mac)
    if (-not $mac) { return "Unknown" }
    $prefix = ($mac -replace "[:-]", "").Substring(0,6).ToUpper()

    if ($Global:VendorCache.ContainsKey($prefix)) {
        return $Global:VendorCache[$prefix]
    }

    try {
        $url = "https://api.macvendors.com/$mac"
        $response = Invoke-RestMethod -Uri $url -TimeoutSec 4 -ErrorAction Stop
        if ($response -and $response -notmatch "not found") {
            $Global:VendorCache[$prefix] = $response
            return $response
        }
    } catch {}

    $vendors = @{
        "00163E" = "Cisco"
        "001E58" = "TP-Link"
        "3C84C6" = "Netgear"
        "F8D111" = "Ubiquiti"
        "E8CC18" = "Huawei"
        "AC9E17" = "ASUS"
        "F4F26D" = "Apple"
        "D8C4E9" = "Samsung"
        "A4CF12" = "Intel"
    }
    if ($vendors.ContainsKey($prefix)) {
        $Global:VendorCache[$prefix] = $vendors[$prefix]
        return $vendors[$prefix]
    }

    $Global:VendorCache[$prefix] = "Unknown"
    return "Unknown"
}

function Parse-NetshNetworks {
    $raw = netsh wlan show networks mode=bssid 2>$null
    if (-not $raw) { return @() }

    $networks = @()
    $currentSSID = $null
    for ($i = 0; $i -lt $raw.Length; $i++) {
        $line = $raw[$i].Trim()
        if ($line -match '^SSID\s+\d+\s+:\s+(.*)$') {
            $currentSSID = $matches[1].Trim()
            continue
        }
        if ($line -match '^BSSID\s+\d+\s+:\s+(.*)$') {
            $bssid = $matches[1].Trim()
            $signal = ""
            $channel = ""
            $band = ""
            $rssi = ""
            $protocol = ""
            for ($j=1; $j -le 6; $j++) {
                if ($i+$j -ge $raw.Length) { break }
                $l2 = $raw[$i+$j].Trim()
                if ($l2 -match '^Signal\s+:\s+(.*)$') { $signal = $matches[1].Trim() }
                if ($l2 -match '^Channel\s+:\s+(.*)$') { $channel = $matches[1].Trim() }
                if ($l2 -match '^Radio type\s+:\s+(.*)$') { $protocol = $matches[1].Trim() }
            }

            $ch = [int]($channel -replace '[^0-9]', '')
            if ($ch -ge 1 -and $ch -le 14) { $band = "2.4 GHz" }
            elseif ($ch -ge 36 -and $ch -le 165) { $band = "5 GHz" }
            elseif ($ch -ge 180) { $band = "6 GHz" }
            else { $band = "Unknown" }

            $signalNum = [int]($signal -replace '[^0-9]', '')
            $rssi = [int](-100 + ($signalNum * 0.6))

            if ($signalNum -ge 80) { $bars = "▮▮▮▮"; $color = "LimeGreen" }
            elseif ($signalNum -ge 60) { $bars = "▮▮▮"; $color = "Gold" }
            elseif ($signalNum -ge 40) { $bars = "▮▮"; $color = "DarkOrange" }
            else { $bars = "▮"; $color = "Tomato" }

            $vendor = Get-Vendor (Normalize-Mac $bssid)

            $networks += [pscustomobject]@{
                SSID = $currentSSID
                BSSID = $bssid
                Vendor = $vendor
                RSSI = "$rssi dBm"
                Signal = "$signal"
                Bars = $bars
                Color = $color
                Channel = $channel
                Band = $band
                Protocol = $protocol
            }
        }
    }
    return ,$networks
}

function Get-ConnectedWifiInfo {
    $raw = netsh wlan show interfaces 2>$null
    if (-not $raw) { return $null }
    $info = @{}
    foreach ($line in $raw) {
        $l = $line.Trim()
        if ($l -match '^Name\s+:\s+(.*)$') { $info.Interface = $matches[1].Trim() }
        if ($l -match '^State\s+:\s+(.*)$') { $info.State = $matches[1].Trim() }
        if ($l -match '^SSID\s+:\s+(.*)$') { $info.SSID = $matches[1].Trim() }
        if ($l -match '^BSSID\s+:\s+(.*)$') { $info.BSSID = $matches[1].Trim() }
    }
    return (New-Object psobject -Property $info)
}

# === GUI ===
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Wi-Fi Scanner" Height="650" Width="1000" Background="#f5f6fa"
        WindowStartupLocation="CenterScreen">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#2f3640" Padding="12" CornerRadius="8">
            <StackPanel Orientation="Horizontal">
                <TextBlock Text="&#128225;" FontSize="22" Margin="0,0,10,0"/>
                <TextBlock Text="Wi-Fi Network Scanner" Foreground="White" FontSize="18" FontWeight="Bold"/>
            </StackPanel>
        </Border>

        <StackPanel Orientation="Horizontal" Grid.Row="1" Margin="0,10,0,10" HorizontalAlignment="Left">
            <Button x:Name="ScanButton" Content="&#128269; Scan Now" Width="130" Height="32"
                    Background="#40739e" Foreground="White" BorderThickness="0" Margin="0,0,10,0" Cursor="Hand"/>
            <CheckBox x:Name="AutoRefreshCheck" Content="Auto Refresh" FontWeight="Bold" VerticalAlignment="Center" Margin="5,0,5,0"/>
            <Slider x:Name="IntervalSlider" Minimum="1" Maximum="30" Value="3" Width="120" TickFrequency="1" IsSnapToTickEnabled="True" Margin="5,0,0,0"/>
            <TextBlock x:Name="IntervalLabel" Text="3s" Margin="5,0,10,0" VerticalAlignment="Center"/>
            <Button x:Name="SaveButton" Content="&#128190; Export CSV" Width="130" Height="32"
                    Background="#44bd32" Foreground="White" BorderThickness="0" Cursor="Hand"/>
            <ProgressBar x:Name="ProgressBar" Width="150" Height="12" Margin="20,0,0,0" Visibility="Collapsed"/>
            <TextBlock x:Name="StatusText" VerticalAlignment="Center" FontWeight="Bold" Margin="10,0,0,0"/>
        </StackPanel>

        <Border Grid.Row="2" Background="White" Padding="10" CornerRadius="8" BorderBrush="#dcdde1" BorderThickness="1" Margin="0,0,0,10">
            <DockPanel LastChildFill="True">
                <TextBlock x:Name="ConnInfo" DockPanel.Dock="Top" FontFamily="Consolas" TextWrapping="Wrap" Foreground="#2f3640"/>
                <DataGrid x:Name="NetworkGrid" AutoGenerateColumns="False" CanUserAddRows="False" IsReadOnly="True"
                          AlternatingRowBackground="#f1f2f6" BorderThickness="0" HeadersVisibility="Column"
                          GridLinesVisibility="None" RowHeight="28" FontSize="13" Margin="0,10,0,0">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="SSID" Binding="{Binding SSID}" Width="2*" />
                        <DataGridTextColumn Header="BSSID" Binding="{Binding BSSID}" Width="2*" />
                        <DataGridTextColumn Header="Vendor" Binding="{Binding Vendor}" Width="2*" />
                        <DataGridTemplateColumn Header="Signal" Width="2*">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                        <TextBlock Text="{Binding Signal}" Foreground="#2f3640" Margin="0,0,5,0"/>
                                        <TextBlock Text="{Binding Bars}" FontWeight="Bold">
                                            <TextBlock.Style>
                                                <Style TargetType="TextBlock">
                                                    <Setter Property="Foreground" Value="Gray"/>
                                                    <Style.Triggers>
                                                        <DataTrigger Binding="{Binding Color}" Value="LimeGreen"><Setter Property="Foreground" Value="LimeGreen"/></DataTrigger>
                                                        <DataTrigger Binding="{Binding Color}" Value="Gold"><Setter Property="Foreground" Value="Gold"/></DataTrigger>
                                                        <DataTrigger Binding="{Binding Color}" Value="DarkOrange"><Setter Property="Foreground" Value="DarkOrange"/></DataTrigger>
                                                        <DataTrigger Binding="{Binding Color}" Value="Tomato"><Setter Property="Foreground" Value="Tomato"/></DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </TextBlock.Style>
                                        </TextBlock>
                                    </StackPanel>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>
                        <DataGridTextColumn Header="RSSI" Binding="{Binding RSSI}" Width="*" />
                        <DataGridTextColumn Header="Channel" Binding="{Binding Channel}" Width="*" />
                        <DataGridTextColumn Header="Band" Binding="{Binding Band}" Width="*" />
                        <DataGridTextColumn Header="802.11" Binding="{Binding Protocol}" Width="*" />
                    </DataGrid.Columns>
                </DataGrid>
            </DockPanel>
        </Border>

        <TextBlock Grid.Row="3" Text="⚙️ Auto-refresh every adjustable seconds (1–30). Vendor data is fetched live from macvendors.com."
                   FontSize="11" Foreground="Gray" HorizontalAlignment="Center"/>
    </Grid>
</Window>
"@

# === Load GUI ===
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
$ScanButton = $window.FindName("ScanButton")
$SaveButton = $window.FindName("SaveButton")
$StatusText = $window.FindName("StatusText")
$ConnInfo = $window.FindName("ConnInfo")
$NetworkGrid = $window.FindName("NetworkGrid")
$ProgressBar = $window.FindName("ProgressBar")
$AutoRefreshCheck = $window.FindName("AutoRefreshCheck")
$IntervalSlider = $window.FindName("IntervalSlider")
$IntervalLabel = $window.FindName("IntervalLabel")

# === Scan Logic ===
function Run-Scan {
    $ProgressBar.Visibility = "Visible"
    $StatusText.Text = "Scanning..."
    $window.Cursor = 'Wait'

    $networks = Parse-NetshNetworks
    $NetworkGrid.ItemsSource = @($networks)

    $conn = Get-ConnectedWifiInfo
    if ($conn -and $conn.State -ieq 'connected') {
        $ConnInfo.Text = "Connected: $($conn.SSID)`nInterface: $($conn.Interface)`nBSSID: $($conn.BSSID)"
    } else {
        $ConnInfo.Text = "Not connected to Wi-Fi."
    }

    $StatusText.Text = "✅ Scan complete — $($networks.Count) networks found."
    $ProgressBar.Visibility = "Collapsed"
    $window.Cursor = 'Arrow'
}

# === Auto Refresh Timer ===
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(3)
$timer.Add_Tick({ if ($AutoRefreshCheck.IsChecked) { Run-Scan } })

# === Event Bindings ===
$ScanButton.Add_Click({ Run-Scan })
$AutoRefreshCheck.Add_Checked({ $timer.Start(); $StatusText.Text = "&#128257; Auto-refresh enabled" })
$AutoRefreshCheck.Add_Unchecked({ $timer.Stop(); $StatusText.Text = "⏹️ Auto-refresh stopped" })
$IntervalSlider.Add_ValueChanged({
    param($s,$e)
    $sec = [math]::Round($s.Value)
    $IntervalLabel.Text = "$sec`s"
    $timer.Interval = [TimeSpan]::FromSeconds($sec)
})
$SaveButton.Add_Click({
    if (-not $NetworkGrid.ItemsSource) {
        [System.Windows.MessageBox]::Show("No results to save. Please scan first.","Export","OK","Information")
        return
    }
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter = "CSV Files|*.csv"
    $dialog.FileName = "WiFiScanResults.csv"
    if ($dialog.ShowDialog()) {
        $NetworkGrid.ItemsSource | Export-Csv -Path $dialog.FileName -NoTypeInformation
        [System.Windows.MessageBox]::Show("Saved to $($dialog.FileName)","Export Complete","OK","Information")
    }
})

# === Launch ===
$window.ShowDialog() | Out-Null
