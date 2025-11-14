<#
.SYNOPSIS
    A full MailKit-powered GUI application for sending emails using SMTP.

.DESCRIPTION
    This PowerShell script provides a modern WPF graphical interface that allows
    users to send emails using MailKit via the Send-MailKitMessage module.
    
    Features:
        • GUI-based SMTP email sender  
        • Supports TLS/SSL  
        • Field validation with status logging  
        • Progress bar for sending status  
        • Dynamic log viewer with color-coded output  
        • Automatically installs Send-MailKitMessage module if missing  
        • Clean, modern UI for Office365, Gmail, Outlook, and custom SMTP servers  

    The script loads a full WPF XAML interface, handles module checks, logs
    events in real-time, validates SMTP credentials, and uses MailKit for
    reliable message delivery.

.NOTES
    Author: capnhowyoudo
    Version: 1.0
    Requirements:
        • Windows PowerShell 5+
        • MailKit module: Send-MailKitMessage
        • .NET WPF components (PresentationFramework)

    Additional Notes:
        • Password is securely converted to SecureString before transmission.
        • SMTP server dropdown includes pre-filled common providers.
        • Status log displays time-stamped events in varying colors.
        • GUI automatically scrolls log output to the latest entry.
        • Works with MFA-enabled accounts using app passwords.
#>

Add-Type -AssemblyName PresentationFramework

# ---------------- GUI Layout ----------------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="MailKit Email Sender"
        Height="740" Width="800"
        WindowStartupLocation="CenterScreen"
        Background="#f4f6f9"
        FontFamily="Segoe UI">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Background="#2f3640" CornerRadius="10" Padding="12" Margin="0,0,0,20">
            <TextBlock Text="&#128231; MailKit Email Sender GUI" FontSize="24" FontWeight="Bold" Foreground="White"/>
        </Border>

        <!-- Main Form -->
        <StackPanel Grid.Row="1" Orientation="Vertical" VerticalAlignment="Top" Margin="5">

            <!-- SMTP Server -->
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                <TextBlock Text="SMTP Server:" Width="170" VerticalAlignment="Center"/>
                <ComboBox Name="SMTPServer" Width="520" IsEditable="True">
                    <ComboBoxItem Content="smtp.office365.com" IsSelected="True"/>
                    <ComboBoxItem Content="smtp-hve.office365.com"/>
                    <ComboBoxItem Content="smtp.gmail.com"/>
                    <ComboBoxItem Content="smtp-mail.outlook.com"/>
                </ComboBox>
            </StackPanel>

            <!-- Port -->
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                <TextBlock Text="Port:" Width="170" VerticalAlignment="Center"/>
                <TextBox Name="Port" Width="520" Text="587"/>
            </StackPanel>

            <!-- From -->
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                <TextBlock Text="From (Email):" Width="170" VerticalAlignment="Center"/>
                <TextBox Name="From" Width="520" Text="sender@example.com"/>
            </StackPanel>

            <!-- To -->
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                <TextBlock Text="To (Recipient):" Width="170" VerticalAlignment="Center"/>
                <TextBox Name="To" Width="520" Text="recipient@example.com"/>
            </StackPanel>

            <!-- Subject -->
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                <TextBlock Text="Subject:" Width="170" VerticalAlignment="Center"/>
                <TextBox Name="Subject" Width="520" Text="This is the subject"/>
            </StackPanel>

            <!-- Body -->
            <TextBlock Text="Message Body:" FontWeight="Bold" Margin="0,5,0,5"/>
            <TextBox Name="Body" Height="120" TextWrapping="Wrap" AcceptsReturn="True"
                     VerticalScrollBarVisibility="Auto" Text="This is the text body."/>

            <!-- Username -->
            <StackPanel Orientation="Horizontal" Margin="0,10,0,8">
                <TextBlock Text="Username:" Width="170" VerticalAlignment="Center"/>
                <TextBox Name="Username" Width="520" Text="sender@example.com"/>
            </StackPanel>

            <!-- Password -->
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                <TextBlock Text="Password:" Width="170" VerticalAlignment="Center"/>
                <PasswordBox Name="Password" Width="520"/>
            </StackPanel>

            <!-- SSL -->
            <CheckBox Name="UseSSL" Content="Use Secure Connection (SSL/TLS)" IsChecked="True" Margin="0,10,0,10"/>

            <!-- Progress -->
            <ProgressBar Name="Progress" Height="18" Minimum="0" Maximum="100" Visibility="Collapsed" Margin="0,8,0,0"/>

            <!-- Status Log -->
            <TextBlock Text="Status Log:" FontWeight="Bold" Margin="0,10,0,4"/>
            <Border Name="StatusBorder" Background="White" CornerRadius="5" Padding="6" Height="120">
                <RichTextBox Name="StatusLog" IsReadOnly="True" VerticalScrollBarVisibility="Auto" Background="White" BorderThickness="0"/>
            </Border>
        </StackPanel>

        <!-- Footer -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,20,0,0">
            <Button Name="SendButton" Content="&#128232; Send Email" Width="180" Height="40" Background="#44bd32"
                    Foreground="White" FontWeight="Bold" Margin="0,0,10,0"/>
            <Button Name="ClearLogButton" Content="&#129529; Clear Log" Width="130" Height="40" Background="#40739e"
                    Foreground="White" FontWeight="Bold" Margin="0,0,10,0"/>
            <Button Name="ExitButton" Content="❌ Exit" Width="110" Height="40" Background="#e84118" Foreground="White"/>
        </StackPanel>
    </Grid>
</Window>
"@

# ---------------- Load GUI ----------------
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    Set-Variable -Name $_.Name -Value $window.FindName($_.Name)
}

# ---------------- Helper Functions ----------------
function LogStatus {
    param(
        [string]$Message,
        [string]$Type = "info"
    )

    $color = switch ($Type.ToLower()) {
        "success" { "Green" }
        "error"   { "Red" }
        "warn"    { "DarkOrange" }
        default   { "Black" }
    }

    $para = New-Object System.Windows.Documents.Paragraph
    $run = New-Object System.Windows.Documents.Run "$(Get-Date -Format 'HH:mm:ss') - $Message"
    $run.Foreground = [System.Windows.Media.Brushes]::$color
    $para.Inlines.Add($run)
    $StatusLog.Document.Blocks.Add($para)
    $StatusLog.ScrollToEnd()
}

function Ensure-MailKitModule {
    try {
        if (-not (Get-Module -ListAvailable -Name Send-MailKitMessage)) {
            LogStatus "&#128230; Installing Send-MailKitMessage module..." "warn"
            Install-Module -Name Send-MailKitMessage -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            LogStatus "✅ Send-MailKitMessage installed successfully." "success"
        } else {
            LogStatus "✔️ Send-MailKitMessage module found." "success"
        }
        Import-Module Send-MailKitMessage -ErrorAction Stop
        return $true
    } catch {
        LogStatus "❌ Failed to install or import Send-MailKitMessage. $($_.Exception.Message)" "error"
        [System.Windows.MessageBox]::Show("Failed to install or import Send-MailKitMessage.`nError: $($_.Exception.Message)", "Module Error", "OK", "Error")
        return $false
    }
}

# ---------------- Event Handlers ----------------
$SendButton.Add_Click({
    if (-not (Ensure-MailKitModule)) { return }

    $SMTP = if ($SMTPServer.SelectedItem) {
        $SMTPServer.SelectedItem.Content
    } else {
        $SMTPServer.Text
    }

    $PortNum = [int]$Port.Text
    $User = $Username.Text
    $Pass = $Password.Password
    $FromEmail = $From.Text
    $ToEmail = $To.Text
    $Subj = $Subject.Text
    $BodyText = $Body.Text
    $UseTLS = $UseSSL.IsChecked

    if (-not $User -or -not $Pass -or -not $ToEmail -or -not $FromEmail) {
        LogStatus "⚠️ Please fill in all required fields before sending." "warn"
        [System.Windows.MessageBox]::Show("Please fill in all required fields before sending.", "Missing Information", "OK", "Warning")
        return
    }

    LogStatus "&#128640; Preparing to send email..." "info"
    $Progress.Visibility = "Visible"
    $Progress.Value = 50

    try {
        $secure = ConvertTo-SecureString $Pass -AsPlainText -Force
        $cred = [PSCredential]::new($User, $secure)

        $recipientList = [MimeKit.InternetAddressList]::new()
        $recipientList.Add([MimeKit.InternetAddress]$ToEmail)

        $params = @{
            UseSecureConnectionIfAvailable = $UseTLS
            SMTPServer                     = $SMTP
            Credential                     = $cred
            Port                           = $PortNum
            From                           = [MimeKit.MailboxAddress]$FromEmail
            RecipientList                  = $recipientList
            Subject                        = $Subj
            TextBody                       = $BodyText
        }

        LogStatus "&#128228; Sending email via $SMTP ..." "info"
        Send-MailKitMessage @params

        $Progress.Value = 100
        LogStatus "✅ Email sent successfully!" "success"
        [System.Windows.MessageBox]::Show("Email sent successfully!","Success","OK","Information")
    } catch {
        LogStatus "❌ Error: $($_.Exception.Message)" "error"
        [System.Windows.MessageBox]::Show("Failed to send email.`n$($_.Exception.Message)", "Error", "OK", "Error")
    } finally {
        $Progress.Visibility = "Collapsed"
    }
})

$ClearLogButton.Add_Click({
    $StatusLog.Document.Blocks.Clear()
    LogStatus "&#129529; Log cleared." "info"
})

$ExitButton.Add_Click({ $window.Close() })

# ---------------- Run GUI ----------------
$window.ShowDialog() | Out-Null
