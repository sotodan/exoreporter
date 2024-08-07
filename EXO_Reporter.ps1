<#
.SYNOPSIS
  Generates a detailed report of mailboxes in an Exchange Online environment.
.DESCRIPTION
  This script connects to Exchange Online, retrieves mailboxes with a specified domain, and gathers detailed information about each mailbox, including size, item count, and hold statuses. It then generates an HTML report and logs the processing user.
.PARAMETER domain
    The domain to filter mailboxes (e.g., "@M365DS219944.onmicrosoft.com").
.INPUTS
  None
.OUTPUTS
  HTML report stored in C:\temp\ExchangeOnlineMailboxReport.html
  Log file stored in C:\temp\ProcessingUserLog.txt
.NOTES
  Version:        1.0
  Author:         Daniel Soto
  GitHub:         Sotodan
  Creation Date:  07/14/2024
  Purpose/Change: Initial script development
#>

# Define the domain to filter mailboxes
$domain = "@M365DS219944.onmicrosoft.com"

# Get the current user for tracking
$processingUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

# Import Exchange Online module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName admin@M365DS219944.onmicrosoft.com -ShowProgress $true

# Get mailboxes with the specified domain
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.PrimarySmtpAddress -like "*$domain" }

# Initialize an array to hold mailbox information
$mailboxInfo = @()

# Initialize progress bar
$totalMailboxes = $mailboxes.Count
$currentMailbox = 0
$totalMailboxSize = 0
$totalMailboxesOnHold = 0

foreach ($mailbox in $mailboxes) {
    $currentMailbox++
    Write-Progress -Activity "Processing Mailboxes" -Status "Processing $currentMailbox of $totalMailboxes" -PercentComplete (($currentMailbox / $totalMailboxes) * 100)

    $mailboxStats = Get-MailboxStatistics -Identity $mailbox.PrimarySmtpAddress
    $mailboxSizeString = $mailboxStats.TotalItemSize.ToString()
    $mailboxSizeValue = [double]::Parse($mailboxSizeString -replace '[^\d.]')
    $mailboxSizeGB = if ($mailboxSizeString -like "*MB*") { [math]::Round($mailboxSizeValue / 1024, 2) } elseif ($mailboxSizeString -like "*KB*") { [math]::Round($mailboxSizeValue / (1024 * 1024), 2) } else { [math]::Round($mailboxSizeValue, 2) }
    $totalMailboxSize += $mailboxSizeGB

    if ($mailbox.LitigationHoldEnabled -eq $true -or $mailbox.InPlaceHolds.Count -gt 0) {
        $totalMailboxesOnHold++
    }

    $mailboxDetails = [PSCustomObject]@{
        PrimarySMTPAddress     = $mailbox.PrimarySmtpAddress
        SecondaryEmailAddress  = ($mailbox.EmailAddresses -ne $mailbox.PrimarySmtpAddress -and $_ -like "SMTP:*") -join ", "
        MailboxSize            = "$mailboxSizeGB GB"
        ItemCount              = $mailboxStats.ItemCount
        LitigationHoldEnabled  = $mailbox.LitigationHoldEnabled
        LitigationHoldDate     = $mailbox.LitigationHoldDate
        LitigationHoldOwner    = $mailbox.LitigationHoldOwner
        LitigationHoldDuration = $mailbox.LitigationHoldDuration
        InPlaceHolds           = ($mailbox.InPlaceHolds -join ", ")
        ComplianceTagHoldApplied = $mailbox.ComplianceTagHoldApplied
        exchangeGUID           = $mailbox.ExchangeGuid
    }
    $mailboxInfo += $mailboxDetails
}

# Generate HTML report
$htmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Mailbox Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f9f9f9;
        }
        .container {
            width: 90%;
            margin: auto;
            overflow: hidden;
        }
        header {
            background: #0078d4;
            color: #fff;
            padding: 20px 0;
            text-align: center;
            border-radius: 8px 8px 0 0;
        }
        header h1 {
            margin: 0;
            font-size: 2em;
        }
        table {
            width: 100%;
            margin: 20px 0;
            border-collapse: collapse;
            box-shadow: 0 2px 3px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
        }
        table, th, td {
            border: 1px solid #ddd;
        }
        th, td {
            padding: 12px;
            text-align: left;
        }
        th {
            background-color: #0078d4;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        tr:hover {
            background-color: #e1f5fe;
        }
        .summary {
            margin: 20px 0;
            padding: 10px;
            background-color: #e2e2e2;
            border: 1px solid #ccc;
            display: flex;
            justify-content: space-between;
            border-radius: 8px;
        }
        .summary div {
            flex: 1;
            padding: 10px;
            background-color: #fff;
            border: 1px solid #ccc;
            margin: 5px;
            text-align: center;
            border-radius: 8px;
        }
    </style>
</head>
<body>
    <header>
        <h1>Mailbox Report</h1>
    </header>
    <div class="container">
        <div class="summary">
            <div>
                <p><strong>Total Mailboxes</strong></p>
                <p>$totalMailboxes</p>
            </div>
            <div>
                <p><strong>Total Mailbox Size</strong></p>
                <p>$([math]::Round($totalMailboxSize, 2)) GB</p>
            </div>
            <div>
                <p><strong>Total Mailboxes on Hold</strong></p>
                <p>$totalMailboxesOnHold</p>
            </div>
            <div>
                <p><strong>Report Generated On</strong></p>
                <p>$(Get-Date)</p>
            </div>
        </div>
        <table>
            <thead>
                <tr>
                    <th>Primary SMTP Address</th>
                    <th>Secondary Email Address</th>
                    <th>Mailbox Size</th>
                    <th>Item Count</th>
                    <th>Litigation Hold Enabled</th>
                    <th>Litigation Hold Date</th>
                    <th>Litigation Hold Owner</th>
                    <th>Litigation Hold Duration</th>
                    <th>In-Place Holds</th>
                    <th>Compliance Tag Hold Applied</th>
                    <th>Exchange GUID</th>
                </tr>
            </thead>
            <tbody>
"@

foreach ($mailbox in $mailboxInfo) {
    $htmlReport += "<tr>"
    $htmlReport += "<td>$($mailbox.PrimarySMTPAddress)</td>"
    $htmlReport += "<td>$($mailbox.SecondaryEmailAddress)</td>"
    $htmlReport += "<td>$($mailbox.MailboxSize)</td>"
    $htmlReport += "<td>$($mailbox.ItemCount)</td>"
    $htmlReport += "<td>$($mailbox.LitigationHoldEnabled)</td>"
    $htmlReport += "<td>$($mailbox.LitigationHoldDate)</td>"
    $htmlReport += "<td>$($mailbox.LitigationHoldOwner)</td>"
    $htmlReport += "<td>$($mailbox.LitigationHoldDuration)</td>"
    $htmlReport += "<td>$($mailbox.InPlaceHolds)</td>"
    $htmlReport += "<td>$($mailbox.ComplianceTagHoldApplied)</td>"
    $htmlReport += "<td>$($mailbox.exchangeGUID)</td>"
    $htmlReport += "</tr>"
}

$htmlReport += @"
            </tbody>
        </table>
    </div>
</body>
</html>
"@

# Save the HTML report to a file
$reportPath = "C:\temp\ExchangeOnlineMailboxReport.html"
$htmlReport | Out-File -FilePath $reportPath

# Log the processing user
$logPath = "C:\temp\ProcessingUserLog.txt"
Add-Content -Path $logPath -Value "Report generated by: $processingUser on $(Get-Date)"

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

Write-Output "Report generated and saved to $reportPath"
Write-Output "Processing user logged in $logPath"
