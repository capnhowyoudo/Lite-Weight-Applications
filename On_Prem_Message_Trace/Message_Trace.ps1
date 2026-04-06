#requires -version 3.0
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================
# Exchange On-Prem Message Trace GUI
# Clean Gray Version
# ============================================================

# -----------------------------
# Global variables
# -----------------------------
$script:ExchangeSession = $null
$script:TraceResults    = New-Object System.Collections.ArrayList

# -----------------------------
# Theme
# -----------------------------
$FormBackColor      = [System.Drawing.Color]::FromArgb(242, 242, 242)
$PanelBackColor     = [System.Drawing.Color]::FromArgb(230, 230, 230)
$HeaderBackColor    = [System.Drawing.Color]::FromArgb(210, 210, 210)
$GridHeaderColor    = [System.Drawing.Color]::FromArgb(220, 220, 220)
$BorderColor        = [System.Drawing.Color]::FromArgb(180, 180, 180)
$TextColor          = [System.Drawing.Color]::FromArgb(45, 45, 45)
$ButtonGray         = [System.Drawing.Color]::FromArgb(215, 215, 215)
$ButtonBlue         = [System.Drawing.Color]::FromArgb(0, 120, 215)
$ButtonBlueText     = [System.Drawing.Color]::White

$FontRegular        = New-Object System.Drawing.Font("Segoe UI", 9)
$FontBold           = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$FontHeader         = New-Object System.Drawing.Font("Segoe UI Semibold", 12)

# -----------------------------
# Helper styling functions
# -----------------------------
function Set-ButtonStyle {
    param(
        [System.Windows.Forms.Button]$Button,
        [System.Drawing.Color]$BackColor = $ButtonGray,
        [System.Drawing.Color]$ForeColor = $TextColor
    )

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderColor = $BorderColor
    $Button.FlatAppearance.BorderSize = 1
    $Button.BackColor = $BackColor
    $Button.ForeColor = $ForeColor
    $Button.Font = $FontRegular
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function Set-TextBoxStyle {
    param([System.Windows.Forms.TextBox]$TextBox)

    $TextBox.Font = $FontRegular
    $TextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $TextBox.BackColor = [System.Drawing.Color]::White
    $TextBox.ForeColor = $TextColor
}

function Set-LabelStyle {
    param(
        [System.Windows.Forms.Label]$Label,
        [System.Drawing.Font]$Font = $FontRegular
    )

    $Label.Font = $Font
    $Label.ForeColor = $TextColor
    $Label.BackColor = [System.Drawing.Color]::Transparent
}

# -----------------------------
# Functions
# -----------------------------
function Write-Status {
    param(
        [string]$Message,
        [string]$Color = "Black"
    )
    $txtStatus.SelectionColor = [System.Drawing.Color]::$Color
    $txtStatus.AppendText("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message`r`n")
    $txtStatus.ScrollToCaret()
}

function Disconnect-Exchange {
    try {
        if ($script:ExchangeSession) {
            Remove-PSSession -Session $script:ExchangeSession -ErrorAction SilentlyContinue
            $script:ExchangeSession = $null
            Write-Status "Disconnected from Exchange session." "DarkOrange"
        }

        if (Get-Command Get-MessageTrackingLog -ErrorAction SilentlyContinue) {
            Write-Status "Exchange remote session removed." "DarkOrange"
        }

        $lblConnectionStatus.Text = "Not Connected"
        $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Red
    }
    catch {
        Write-Status "Disconnect error: $($_.Exception.Message)" "Red"
    }
}

function Connect-Exchange {
    param(
        [string]$ExchangeServer
    )

    try {
        if ([string]::IsNullOrWhiteSpace($ExchangeServer)) {
            throw "Please enter an Exchange server name or FQDN."
        }

        Disconnect-Exchange

        Write-Status "Connecting to Exchange server: $ExchangeServer" "Blue"

        $connectionUri = "http://$ExchangeServer/PowerShell/"

        $script:ExchangeSession = New-PSSession `
            -ConfigurationName Microsoft.Exchange `
            -ConnectionUri $connectionUri `
            -Authentication Kerberos `
            -ErrorAction Stop

        Import-PSSession $script:ExchangeSession -DisableNameChecking -AllowClobber -ErrorAction Stop | Out-Null

        $lblConnectionStatus.Text = "Connected: $ExchangeServer"
        $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Green

        Write-Status "Connected and Exchange session imported successfully." "Green"
    }
    catch {
        $lblConnectionStatus.Text = "Not Connected"
        $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Red
        Write-Status "Failed to connect: $($_.Exception.Message)" "Red"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Exchange server.`r`n`r`n$($_.Exception.Message)",
            "Connection Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

function Ensure-Connected {
    if (-not $script:ExchangeSession) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please connect to an Exchange server first.",
            "Not Connected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return $false
    }
    return $true
}

function Clear-Grid {
    $gridResults.DataSource = $null
    $gridResults.Rows.Clear()
    $gridResults.Columns.Clear()
}

function Show-ResultsInGrid {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$Data
    )

    try {
        Clear-Grid

        $table = New-Object System.Data.DataTable

        $first = $Data | Select-Object -First 1
        if (-not $first) {
            Write-Status "No results found." "DarkOrange"
            return
        }

        foreach ($prop in $first.PSObject.Properties.Name) {
            [void]$table.Columns.Add($prop)
        }

        foreach ($item in $Data) {
            $row = $table.NewRow()
            foreach ($prop in $item.PSObject.Properties.Name) {
                if ($item.$prop -is [System.Array]) {
                    $row[$prop] = ($item.$prop -join "; ")
                }
                else {
                    $row[$prop] = [string]$item.$prop
                }
            }
            [void]$table.Rows.Add($row)
        }

        $gridResults.DataSource = $table
        Write-Status "Loaded $($table.Rows.Count) result(s) into grid." "Green"
    }
    catch {
        Write-Status "Failed to display results: $($_.Exception.Message)" "Red"
    }
}

function Save-LastResultsToCsv {
    try {
        if (-not $script:TraceResults -or $script:TraceResults.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "There are no results to export.",
                "Export CSV",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv"
        $saveDialog.Title = "Save Message Trace Results"
        $saveDialog.FileName = "MessageTrace_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")

        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:TraceResults | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
            Write-Status "Results exported to $($saveDialog.FileName)" "Green"
        }
    }
    catch {
        Write-Status "Export failed: $($_.Exception.Message)" "Red"
    }
}

function Get-DateRange {
    $startDate = $dtpStart.Value
    $endDate   = $dtpEnd.Value

    if ($endDate -lt $startDate) {
        throw "End date cannot be earlier than start date."
    }

    $endDate = $endDate.Date.AddDays(1).AddSeconds(-1)

    return @{
        Start = $startDate
        End   = $endDate
    }
}

function Run-TraceQuery {
    param(
        [ValidateSet("Sender", "Recipient", "All", "SenderFullHistory", "SenderLast7", "RecipientLast1", "RecipientLast7", "RecipientLast30")]
        [string]$Mode
    )

    if (-not (Ensure-Connected)) { return }

    try {
        $range = Get-DateRange
        $sender     = $txtSender.Text.Trim()
        $recipient  = $txtRecipient.Text.Trim()
        $subject    = $txtSubject.Text.Trim()
        $resultSize = if ([string]::IsNullOrWhiteSpace($txtResultSize.Text)) { "Unlimited" } else { $txtResultSize.Text.Trim() }

        Write-Status "Running query mode: $Mode" "Blue"

        switch ($Mode) {
            "Sender" {
                if ([string]::IsNullOrWhiteSpace($sender)) { throw "Enter a sender email address." }

                $results = Get-MessageTrackingLog `
                    -Sender $sender `
                    -Start $range.Start `
                    -End $range.End `
                    -ResultSize $resultSize |
                    Select-Object Timestamp, EventId, Source, MessageSubject,
                        @{Name="All_Recipients";Expression={($_.Recipients -join '; ')}},
                        @{Name="RecipientStatus";Expression={($_.RecipientStatus -join '; ')}} |
                    Sort-Object Timestamp -Descending
            }

            "Recipient" {
                if ([string]::IsNullOrWhiteSpace($recipient)) { throw "Enter a recipient email address." }

                $results = Get-MessageTrackingLog `
                    -Recipient $recipient `
                    -Start $range.Start `
                    -End $range.End `
                    -ResultSize $resultSize |
                    Select-Object Timestamp, Sender, Recipients, EventId, Source, MessageSubject |
                    Sort-Object Timestamp -Descending
            }

            "All" {
                $results = Get-MessageTrackingLog `
                    -Start $range.Start `
                    -End $range.End `
                    -ResultSize $resultSize |
                    Select-Object Timestamp, Sender, Recipients, EventId, Source, MessageSubject |
                    Sort-Object Timestamp -Descending
            }

            "SenderFullHistory" {
                if ([string]::IsNullOrWhiteSpace($sender)) { throw "Enter a sender email address." }

                $results = Get-MessageTrackingLog `
                    -Sender $sender `
                    -Start (Get-Date).AddDays(-30) `
                    -ResultSize $resultSize |
                    Select-Object Timestamp, EventId, Source, MessageSubject, Recipients, RecipientStatus,
                        @{Name="Database";Expression={$_.SourceContext}} |
                    Sort-Object Timestamp -Descending
            }

            "SenderLast7" {
                if ([string]::IsNullOrWhiteSpace($sender)) { throw "Enter a sender email address." }

                $results = Get-MessageTrackingLog `
                    -Sender $sender `
                    -Start (Get-Date).AddDays(-7) `
                    -ResultSize $resultSize |
                    Select-Object Timestamp, EventId, Source, MessageSubject,
                        @{Name="All_Recipients";Expression={$_.Recipients}},
                        RecipientStatus |
                    Sort-Object Timestamp -Descending
            }

            "RecipientLast1" {
                if ([string]::IsNullOrWhiteSpace($recipient)) { throw "Enter a recipient email address." }

                $results = Get-MessageTrackingLog `
                    -Recipient $recipient `
                    -Start (Get-Date).AddDays(-1) `
                    -ResultSize $resultSize |
                    Select-Object Timestamp, Sender, EventId, Source, MessageSubject |
                    Sort-Object Timestamp -Descending
            }

            "RecipientLast7" {
                if ([string]::IsNullOrWhiteSpace($recipient)) { throw "Enter a recipient email address." }

                $results = Get-MessageTrackingLog `
                    -Recipient $recipient `
                    -Start (Get-Date).AddDays(-7) `
                    -ResultSize $resultSize |
                    Select-Object Timestamp, Sender, EventId, Source, MessageSubject |
                    Sort-Object Timestamp -Descending
            }

            "RecipientLast30" {
                if ([string]::IsNullOrWhiteSpace($recipient)) { throw "Enter a recipient email address." }

                $results = Get-MessageTrackingLog `
                    -Recipient $recipient `
                    -Start (Get-Date).AddDays(-30) `
                    -ResultSize $resultSize |
                    Select-Object Timestamp, Sender, EventId, Source, MessageSubject |
                    Sort-Object Timestamp -Descending
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($subject)) {
            $results = $results | Where-Object { $_.MessageSubject -like "*$subject*" }
        }

        $script:TraceResults = New-Object System.Collections.ArrayList
        foreach ($item in $results) {
            [void]$script:TraceResults.Add($item)
        }

        Show-ResultsInGrid -Data $script:TraceResults
    }
    catch {
        Write-Status "Query failed: $($_.Exception.Message)" "Red"
        [System.Windows.Forms.MessageBox]::Show(
            $($_.Exception.Message),
            "Query Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

# -----------------------------
# Form
# -----------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Exchange On-Prem Message Trace GUI"
$form.Size = New-Object System.Drawing.Size(1240, 780)
$form.StartPosition = "CenterScreen"
$form.TopMost = $false
$form.BackColor = $FormBackColor
$form.Font = $FontRegular

# -----------------------------
# Header Panel
# -----------------------------
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Location = New-Object System.Drawing.Point(0, 0)
$pnlHeader.Size = New-Object System.Drawing.Size(1240, 50)
$pnlHeader.BackColor = $HeaderBackColor
$pnlHeader.Anchor = "Top,Left,Right"
$form.Controls.Add($pnlHeader)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Exchange Message Trace"
$lblTitle.Location = New-Object System.Drawing.Point(15, 13)
$lblTitle.AutoSize = $true
Set-LabelStyle -Label $lblTitle -Font $FontHeader
$pnlHeader.Controls.Add($lblTitle)

# -----------------------------
# Connection Panel
# -----------------------------
$grpConnection = New-Object System.Windows.Forms.GroupBox
$grpConnection.Text = "Exchange Connection"
$grpConnection.Location = New-Object System.Drawing.Point(10, 60)
$grpConnection.Size = New-Object System.Drawing.Size(1200, 80)
$grpConnection.BackColor = $PanelBackColor
$grpConnection.ForeColor = $TextColor
$grpConnection.Font = $FontBold
$form.Controls.Add($grpConnection)

$lblServer = New-Object System.Windows.Forms.Label
$lblServer.Text = "Exchange Server / FQDN:"
$lblServer.Location = New-Object System.Drawing.Point(15, 35)
$lblServer.AutoSize = $true
Set-LabelStyle -Label $lblServer
$grpConnection.Controls.Add($lblServer)

$txtServer = New-Object System.Windows.Forms.TextBox
$txtServer.Location = New-Object System.Drawing.Point(170, 31)
$txtServer.Size = New-Object System.Drawing.Size(280, 24)
Set-TextBoxStyle -TextBox $txtServer
$grpConnection.Controls.Add($txtServer)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect"
$btnConnect.Location = New-Object System.Drawing.Point(470, 28)
$btnConnect.Size = New-Object System.Drawing.Size(100, 28)
Set-ButtonStyle -Button $btnConnect -BackColor $ButtonBlue -ForeColor $ButtonBlueText
$grpConnection.Controls.Add($btnConnect)

$btnDisconnect = New-Object System.Windows.Forms.Button
$btnDisconnect.Text = "Disconnect"
$btnDisconnect.Location = New-Object System.Drawing.Point(580, 28)
$btnDisconnect.Size = New-Object System.Drawing.Size(100, 28)
Set-ButtonStyle -Button $btnDisconnect
$grpConnection.Controls.Add($btnDisconnect)

$lblConnectionStatusTitle = New-Object System.Windows.Forms.Label
$lblConnectionStatusTitle.Text = "Status:"
$lblConnectionStatusTitle.Location = New-Object System.Drawing.Point(705, 35)
$lblConnectionStatusTitle.AutoSize = $true
Set-LabelStyle -Label $lblConnectionStatusTitle
$grpConnection.Controls.Add($lblConnectionStatusTitle)

$lblConnectionStatus = New-Object System.Windows.Forms.Label
$lblConnectionStatus.Text = "Not Connected"
$lblConnectionStatus.Location = New-Object System.Drawing.Point(755, 35)
$lblConnectionStatus.AutoSize = $true
$lblConnectionStatus.ForeColor = [System.Drawing.Color]::Red
$lblConnectionStatus.Font = $FontBold
$grpConnection.Controls.Add($lblConnectionStatus)

# -----------------------------
# Search Group
# -----------------------------
$grpSearch = New-Object System.Windows.Forms.GroupBox
$grpSearch.Text = "Search Criteria"
$grpSearch.Location = New-Object System.Drawing.Point(10, 150)
$grpSearch.Size = New-Object System.Drawing.Size(1200, 190)
$grpSearch.BackColor = $PanelBackColor
$grpSearch.ForeColor = $TextColor
$grpSearch.Font = $FontBold
$form.Controls.Add($grpSearch)

$lblSender = New-Object System.Windows.Forms.Label
$lblSender.Text = "Sender:"
$lblSender.Location = New-Object System.Drawing.Point(15, 32)
$lblSender.AutoSize = $true
Set-LabelStyle -Label $lblSender
$grpSearch.Controls.Add($lblSender)

$txtSender = New-Object System.Windows.Forms.TextBox
$txtSender.Location = New-Object System.Drawing.Point(80, 28)
$txtSender.Size = New-Object System.Drawing.Size(260, 24)
Set-TextBoxStyle -TextBox $txtSender
$grpSearch.Controls.Add($txtSender)

$lblRecipient = New-Object System.Windows.Forms.Label
$lblRecipient.Text = "Recipient:"
$lblRecipient.Location = New-Object System.Drawing.Point(360, 32)
$lblRecipient.AutoSize = $true
Set-LabelStyle -Label $lblRecipient
$grpSearch.Controls.Add($lblRecipient)

$txtRecipient = New-Object System.Windows.Forms.TextBox
$txtRecipient.Location = New-Object System.Drawing.Point(430, 28)
$txtRecipient.Size = New-Object System.Drawing.Size(260, 24)
Set-TextBoxStyle -TextBox $txtRecipient
$grpSearch.Controls.Add($txtRecipient)

$lblSubject = New-Object System.Windows.Forms.Label
$lblSubject.Text = "Subject Contains:"
$lblSubject.Location = New-Object System.Drawing.Point(710, 32)
$lblSubject.AutoSize = $true
Set-LabelStyle -Label $lblSubject
$grpSearch.Controls.Add($lblSubject)

$txtSubject = New-Object System.Windows.Forms.TextBox
$txtSubject.Location = New-Object System.Drawing.Point(815, 28)
$txtSubject.Size = New-Object System.Drawing.Size(220, 24)
Set-TextBoxStyle -TextBox $txtSubject
$grpSearch.Controls.Add($txtSubject)

$lblStart = New-Object System.Windows.Forms.Label
$lblStart.Text = "Start Date:"
$lblStart.Location = New-Object System.Drawing.Point(15, 72)
$lblStart.AutoSize = $true
Set-LabelStyle -Label $lblStart
$grpSearch.Controls.Add($lblStart)

$dtpStart = New-Object System.Windows.Forms.DateTimePicker
$dtpStart.Location = New-Object System.Drawing.Point(80, 68)
$dtpStart.Size = New-Object System.Drawing.Size(200, 24)
$dtpStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtpStart.CustomFormat = "yyyy-MM-dd HH:mm"
$dtpStart.Value = (Get-Date).AddDays(-1)
$grpSearch.Controls.Add($dtpStart)

$lblEnd = New-Object System.Windows.Forms.Label
$lblEnd.Text = "End Date:"
$lblEnd.Location = New-Object System.Drawing.Point(300, 72)
$lblEnd.AutoSize = $true
Set-LabelStyle -Label $lblEnd
$grpSearch.Controls.Add($lblEnd)

$dtpEnd = New-Object System.Windows.Forms.DateTimePicker
$dtpEnd.Location = New-Object System.Drawing.Point(360, 68)
$dtpEnd.Size = New-Object System.Drawing.Size(200, 24)
$dtpEnd.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtpEnd.CustomFormat = "yyyy-MM-dd HH:mm"
$dtpEnd.Value = Get-Date
$grpSearch.Controls.Add($dtpEnd)

$lblResultSize = New-Object System.Windows.Forms.Label
$lblResultSize.Text = "ResultSize:"
$lblResultSize.Location = New-Object System.Drawing.Point(585, 72)
$lblResultSize.AutoSize = $true
Set-LabelStyle -Label $lblResultSize
$grpSearch.Controls.Add($lblResultSize)

$txtResultSize = New-Object System.Windows.Forms.TextBox
$txtResultSize.Location = New-Object System.Drawing.Point(660, 68)
$txtResultSize.Size = New-Object System.Drawing.Size(110, 24)
$txtResultSize.Text = "Unlimited"
Set-TextBoxStyle -TextBox $txtResultSize
$grpSearch.Controls.Add($txtResultSize)

# Row 1 buttons
$btnSender = New-Object System.Windows.Forms.Button
$btnSender.Text = "Search Sender"
$btnSender.Location = New-Object System.Drawing.Point(15, 110)
$btnSender.Size = New-Object System.Drawing.Size(120, 30)
Set-ButtonStyle -Button $btnSender -BackColor $ButtonBlue -ForeColor $ButtonBlueText
$grpSearch.Controls.Add($btnSender)

$btnRecipient = New-Object System.Windows.Forms.Button
$btnRecipient.Text = "Search Recipient"
$btnRecipient.Location = New-Object System.Drawing.Point(145, 110)
$btnRecipient.Size = New-Object System.Drawing.Size(130, 30)
Set-ButtonStyle -Button $btnRecipient -BackColor $ButtonBlue -ForeColor $ButtonBlueText
$grpSearch.Controls.Add($btnRecipient)

$btnAll = New-Object System.Windows.Forms.Button
$btnAll.Text = "Search All"
$btnAll.Location = New-Object System.Drawing.Point(285, 110)
$btnAll.Size = New-Object System.Drawing.Size(100, 30)
Set-ButtonStyle -Button $btnAll -BackColor $ButtonBlue -ForeColor $ButtonBlueText
$grpSearch.Controls.Add($btnAll)

$btnSender30 = New-Object System.Windows.Forms.Button
$btnSender30.Text = "Sender Last 30 Days"
$btnSender30.Location = New-Object System.Drawing.Point(395, 110)
$btnSender30.Size = New-Object System.Drawing.Size(150, 30)
Set-ButtonStyle -Button $btnSender30
$grpSearch.Controls.Add($btnSender30)

$btnSender7 = New-Object System.Windows.Forms.Button
$btnSender7.Text = "Sender Last 7 Days"
$btnSender7.Location = New-Object System.Drawing.Point(555, 110)
$btnSender7.Size = New-Object System.Drawing.Size(140, 30)
Set-ButtonStyle -Button $btnSender7
$grpSearch.Controls.Add($btnSender7)

$btnRecipient1 = New-Object System.Windows.Forms.Button
$btnRecipient1.Text = "Recipient Last 1 Day"
$btnRecipient1.Location = New-Object System.Drawing.Point(705, 110)
$btnRecipient1.Size = New-Object System.Drawing.Size(145, 30)
Set-ButtonStyle -Button $btnRecipient1
$grpSearch.Controls.Add($btnRecipient1)

# Row 2 buttons
$btnRecipient7 = New-Object System.Windows.Forms.Button
$btnRecipient7.Text = "Recipient Last 7 Days"
$btnRecipient7.Location = New-Object System.Drawing.Point(15, 146)
$btnRecipient7.Size = New-Object System.Drawing.Size(150, 30)
Set-ButtonStyle -Button $btnRecipient7
$grpSearch.Controls.Add($btnRecipient7)

$btnRecipient30 = New-Object System.Windows.Forms.Button
$btnRecipient30.Text = "Recipient Last 30 Days"
$btnRecipient30.Location = New-Object System.Drawing.Point(175, 146)
$btnRecipient30.Size = New-Object System.Drawing.Size(155, 30)
Set-ButtonStyle -Button $btnRecipient30
$grpSearch.Controls.Add($btnRecipient30)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export CSV"
$btnExport.Location = New-Object System.Drawing.Point(350, 146)
$btnExport.Size = New-Object System.Drawing.Size(110, 30)
Set-ButtonStyle -Button $btnExport
$grpSearch.Controls.Add($btnExport)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear Results"
$btnClear.Location = New-Object System.Drawing.Point(470, 146)
$btnClear.Size = New-Object System.Drawing.Size(120, 30)
Set-ButtonStyle -Button $btnClear
$grpSearch.Controls.Add($btnClear)

# -----------------------------
# Results Grid
# -----------------------------
$gridResults = New-Object System.Windows.Forms.DataGridView
$gridResults.Location = New-Object System.Drawing.Point(10, 350)
$gridResults.Size = New-Object System.Drawing.Size(1200, 290)
$gridResults.ReadOnly = $true
$gridResults.AllowUserToAddRows = $false
$gridResults.AllowUserToDeleteRows = $false
$gridResults.AllowUserToResizeRows = $false
$gridResults.RowHeadersVisible = $false
$gridResults.AutoSizeColumnsMode = "Fill"
$gridResults.Anchor = "Top,Bottom,Left,Right"
$gridResults.BackgroundColor = [System.Drawing.Color]::White
$gridResults.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$gridResults.GridColor = $BorderColor
$gridResults.EnableHeadersVisualStyles = $false
$gridResults.ColumnHeadersDefaultCellStyle.BackColor = $GridHeaderColor
$gridResults.ColumnHeadersDefaultCellStyle.ForeColor = $TextColor
$gridResults.ColumnHeadersDefaultCellStyle.Font = $FontBold
$gridResults.DefaultCellStyle.Font = $FontRegular
$gridResults.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
$gridResults.DefaultCellStyle.ForeColor = $TextColor
$gridResults.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 248, 248)
$form.Controls.Add($gridResults)

# -----------------------------
# Status Box
# -----------------------------
$grpStatus = New-Object System.Windows.Forms.GroupBox
$grpStatus.Text = "Status / Log"
$grpStatus.Location = New-Object System.Drawing.Point(10, 650)
$grpStatus.Size = New-Object System.Drawing.Size(1200, 80)
$grpStatus.Anchor = "Bottom,Left,Right"
$grpStatus.BackColor = $PanelBackColor
$grpStatus.ForeColor = $TextColor
$grpStatus.Font = $FontBold
$form.Controls.Add($grpStatus)

$txtStatus = New-Object System.Windows.Forms.RichTextBox
$txtStatus.Location = New-Object System.Drawing.Point(10, 20)
$txtStatus.Size = New-Object System.Drawing.Size(1180, 50)
$txtStatus.ReadOnly = $true
$txtStatus.Anchor = "Top,Bottom,Left,Right"
$txtStatus.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtStatus.BackColor = [System.Drawing.Color]::White
$txtStatus.Font = $FontRegular
$grpStatus.Controls.Add($txtStatus)

# -----------------------------
# Events
# -----------------------------
$btnConnect.Add_Click({
    Connect-Exchange -ExchangeServer $txtServer.Text.Trim()
})

$btnDisconnect.Add_Click({
    Disconnect-Exchange
})

$btnSender.Add_Click({
    Run-TraceQuery -Mode "Sender"
})

$btnRecipient.Add_Click({
    Run-TraceQuery -Mode "Recipient"
})

$btnAll.Add_Click({
    Run-TraceQuery -Mode "All"
})

$btnSender30.Add_Click({
    Run-TraceQuery -Mode "SenderFullHistory"
})

$btnSender7.Add_Click({
    Run-TraceQuery -Mode "SenderLast7"
})

$btnRecipient1.Add_Click({
    Run-TraceQuery -Mode "RecipientLast1"
})

$btnRecipient7.Add_Click({
    Run-TraceQuery -Mode "RecipientLast7"
})

$btnRecipient30.Add_Click({
    Run-TraceQuery -Mode "RecipientLast30"
})

$btnExport.Add_Click({
    Save-LastResultsToCsv
})

$btnClear.Add_Click({
    $script:TraceResults = New-Object System.Collections.ArrayList
    Clear-Grid
    Write-Status "Results cleared." "DarkOrange"
})

$form.Add_FormClosing({
    Disconnect-Exchange
})

# -----------------------------
# Startup
# -----------------------------
Write-Status "Application started. Enter Exchange server/FQDN and click Connect." "Blue"
[void]$form.ShowDialog()
