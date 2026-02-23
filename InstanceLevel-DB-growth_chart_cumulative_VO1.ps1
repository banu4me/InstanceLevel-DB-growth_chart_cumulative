# ---------------- CONFIG ----------------
$ServerListFile = "G:\Healthcheck-report\cserverlist.txt"
$Database = "master"

# SMTP
$SMTPServer = "smtp.office365.com"
$SMTPPort   = 587
$From       = "announcements@gmail.com"
$To         = "bhanumurthy.msch@gmail.com"
$SMTPUser   = "announcements@gmail.com"
$SMTPPass   = "xxxXXXxxxXXX"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---------------- LAST 12 MONTH AXIS ----------------
$Last12Months = for ($i = 11; $i -ge 0; $i--) {
    (Get-Date).AddMonths(-$i).ToString("yyyy-MM")
}

# ---------------- CORRECT SQL (MONTH-END DB SIZE) ----------------
$SqlQuery = @"
;WITH LastFullPerDbMonth AS
(
    SELECT
        bs.database_name,
        DATEFROMPARTS(YEAR(bs.backup_start_date), MONTH(bs.backup_start_date), 1) AS MonthStart,
        bs.backup_set_id,
        ROW_NUMBER() OVER
        (
            PARTITION BY
                bs.database_name,
                DATEFROMPARTS(YEAR(bs.backup_start_date), MONTH(bs.backup_start_date), 1)
            ORDER BY bs.backup_start_date DESC, bs.backup_set_id DESC
        ) AS rn
    FROM msdb.dbo.backupset bs
    WHERE bs.type = 'D'
      AND bs.database_name NOT IN ('master','model','msdb','tempdb')
      AND bs.backup_start_date >= DATEADD(
            MONTH,-11,DATEFROMPARTS(YEAR(GETDATE()),MONTH(GETDATE()),1))
),
InstanceMonthEndSize AS
(
    SELECT
        FORMAT(MonthStart,'yyyy-MM') AS YearMonth,
        SUM(bf.file_size) / 1073741824.0 AS TotalGB
    FROM LastFullPerDbMonth l
    JOIN msdb.dbo.backupfile bf
        ON bf.backup_set_id = l.backup_set_id
    WHERE rn = 1
    GROUP BY FORMAT(MonthStart,'yyyy-MM')
)
SELECT
    YearMonth,
    CAST(TotalGB AS DECIMAL(18,2)) AS TotalGB
FROM InstanceMonthEndSize
ORDER BY YearMonth;
"@

# ---------------- DATA COLLECTION ----------------
$MonthlyTotals    = @{}
$PerServerMonthly = @{}
$Servers = Get-Content $ServerListFile | Where-Object { $_.Trim() -ne "" }

foreach ($Server in $Servers) {
    Write-Host "Processing server: $Server"
    try {
        $Data = Invoke-Sqlcmd -ServerInstance $Server -Database $Database -Query $SqlQuery -ErrorAction Stop

        $PerServerMonthly[$Server] = @{}

        foreach ($row in $Data) {

            if (-not $MonthlyTotals.ContainsKey($row.YearMonth)) {
                $MonthlyTotals[$row.YearMonth] = 0
            }
            $MonthlyTotals[$row.YearMonth] += [double]$row.TotalGB

            $PerServerMonthly[$Server][$row.YearMonth] = [double]$row.TotalGB
        }
    }
    catch {
        Write-Host "Skipping $Server (connection failed)" -ForegroundColor Yellow
    }
}

# ---------------- CHART FUNCTION (UNCHANGED STYLE) ----------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

function New-GrowthChart {
    param (
        [double[]]$Values,
        [string[]]$MonthAxis
    )

    $Chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $Chart.Width = 1200
    $Chart.Height = 550
    $Chart.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)

    $Area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $Area.AxisX.Interval = 1
    $Area.AxisX.Title = "Month"
    $Area.AxisY.Title = "Total Size (GB)"
    $Area.AxisX.MajorGrid.Enabled = $false
    $Area.AxisY.MajorGrid.Enabled = $false
    $Chart.ChartAreas.Add($Area)

    $Series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    $Series.ChartType = "Column"
    $Series.IsValueShownAsLabel = $true
    $Series["PointWidth"] = "0.6"

    $Colors = @(
        "#4E79A7","#F28E2B","#E15759","#76B7B2",
        "#59A14F","#EDC948","#B07AA1","#FF9DA7",
        "#9C755F","#BAB0AC","#5DA5DA","#B276B2"
    )

    $PeakValue = ($Values | Measure-Object -Maximum).Maximum

    for ($i=0; $i -lt 12; $i++) {
        $dp = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint
        $dp.YValues   = @($Values[$i])
        $dp.AxisLabel = (Get-Date "$($MonthAxis[$i])-01").ToString("MMM yyyy")
        $dp.Color = [System.Drawing.ColorTranslator]::FromHtml($Colors[$i])
        $dp.BackGradientStyle = "TopBottom"
        $dp.BackSecondaryColor = [System.Drawing.Color]::White

        if ($Values[$i] -eq $PeakValue -and $PeakValue -gt 0) {
            $dp.Label = "$($Values[$i]) ★"
        }

        $Series.Points.Add($dp) | Out-Null
    }

    $Chart.Series.Add($Series)

    $Stream = New-Object System.IO.MemoryStream
    $Chart.SaveImage($Stream,"Png")
    $Chart.Dispose()
    $Stream.Position = 0
    return $Stream
}

# ---------------- CUMULATIVE CHART ----------------
$CumValues = foreach ($m in $Last12Months) {
    if ($MonthlyTotals.ContainsKey($m)) { [math]::Round($MonthlyTotals[$m],2) } else { 0 }
}

$CumStream = New-GrowthChart -Values $CumValues -MonthAxis $Last12Months

# ---------------- EMAIL BODY ----------------
$Body = @"
<html>
<body style="font-family:Segoe UI;background:#f4f6f8;padding:20px">

<h2>📊 Database Growth Dashboard (Last 12 Months)</h2>
<img src="cid:CUMULATIVE" style="width:100%;max-width:1000px;border-radius:12px"/>

<h2 style="margin-top:30px">📈 Per Server Database Growth</h2>
"@

# ---------------- PER SERVER CHARTS ----------------
$Attachments = @()

foreach ($Server in $PerServerMonthly.Keys) {

    $Vals = foreach ($m in $Last12Months) {
        if ($PerServerMonthly[$Server].ContainsKey($m)) {
            [math]::Round($PerServerMonthly[$Server][$m],2)
        } else { 0 }
    }

    $Stream = New-GrowthChart -Values $Vals -MonthAxis $Last12Months
    $CID = "SRV_$Server"

    $Body += "<h3>$Server</h3>
              <img src='cid:$CID' style='width:100%;max-width:900px;border-radius:12px'/>"

    $att = New-Object System.Net.Mail.Attachment($Stream,"$Server.png","image/png")
    $att.ContentId = $CID
    $att.ContentDisposition.Inline = $true
    $Attachments += $att
}

# ---------------- SEND EMAIL ----------------
$Mail = New-Object System.Net.Mail.MailMessage
$Mail.From = $From
$Mail.To.Add($To)
$Mail.Subject = "📊 Database Growth Dashboard – Last 12 Months"
$Mail.Body = $Body
$Mail.IsBodyHtml = $true

$CumAtt = New-Object System.Net.Mail.Attachment($CumStream,"cumulative.png","image/png")
$CumAtt.ContentId = "CUMULATIVE"
$CumAtt.ContentDisposition.Inline = $true
$Mail.Attachments.Add($CumAtt)

foreach ($a in $Attachments) { $Mail.Attachments.Add($a) }

$SMTP = New-Object System.Net.Mail.SmtpClient($SMTPServer,$SMTPPort)
$SMTP.EnableSsl = $true
$SMTP.Credentials = New-Object System.Net.NetworkCredential($SMTPUser,$SMTPPass)
$SMTP.Send($Mail)

Write-Host "SUCCESS: DB month-end growth chart matches SQL report"
