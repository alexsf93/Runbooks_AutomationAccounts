<#
.SYNOPSIS
    Generates an HTML report of App Registration certificates expiration.

.DESCRIPTION
    Retrieves application records and certificates from Microsoft Graph API.
    Generates a color-coded HTML report based on days remaining until expiration.
    Sends the report via email using Microsoft Graph.

.PARAMETER GraphClientId
    Application (client) ID registered in Azure AD.

.PARAMETER GraphTenantId
    Directory (tenant) ID in Azure AD.

.PARAMETER GraphSecret
    Application secret.

.PARAMETER Correo_No-Reply
    SMTP credentials (PSCredential).

.EXAMPLE
    # Run as Azure Automation Runbook
    Start-AutomationRunbook -Name "Runbook_Azure-Report_AppRegistrationsExpirationCertificates"

.NOTES
    Author: Automated System
    Version: 1.0
    Prerequisites: Application.Read.All, Mail.Send
#>

# Configurable variables
$clientId = Get-AutomationVariable -Name "GraphClientId"
$tenantId = Get-AutomationVariable -Name "GraphTenantId"
$clientSecret = Get-AutomationVariable -Name "GraphSecret"

$smtpCredential = Get-AutomationPSCredential -Name "Correo_No-Reply"
$smtpUser = $smtpCredential.UserName

# Recipient list
$recipients = @("example1@domain.com", "example2@domain.com")

$useAttachment = 0  # 1 = attachment + body, 0 = body only

# Logos
$Client2Logo = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"
$Client1Logo = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"

# Alert thresholds (in days)
$thresholdExpired = 0
$thresholdAlert = 15
$thresholdCritical = 90
$thresholdWarning = 180

# Get access token
$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$accessToken = $tokenResponse.access_token
$headers = @{ Authorization = "Bearer $accessToken" }

# Get application records
$appRegs = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/applications?`$top=999" -Headers $headers

$certList = @()

foreach ($app in $appRegs.value) {
    if ($app.keyCredentials) {
        foreach ($cert in $app.keyCredentials) {
            $expiresOn = [DateTime]$cert.endDateTime
            $daysRemaining = [math]::Floor(($expiresOn - (Get-Date)).TotalDays)

            $certList += [PSCustomObject]@{
                AppName       = $app.displayName
                AppId         = $app.appId
                CertificateId = $cert.keyId
                ExpiresOn     = $expiresOn
                DaysRemaining = $daysRemaining
            }
        }
    }
}

# Sort by application name
$certList = $certList | Sort-Object AppName

# Get current time in Spain
$reportDate = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), "Romance Standard Time").ToString("dd/MM/yyyy HH:mm")

# Build HTML
if ($certList.Count -eq 0) {
    $htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr><td colspan='5' style='border: 2px solid #000000; text-align:center;'>
    <img src='$Client2Logo' style='width:200px; height:50px;'/>
</td></tr>
<tr><td colspan='5' style='text-align:center; border: 2px solid #000000; font-weight:bold; font-size:18px;'>
    Applications Report - Certificates
    <br/><span style='font-size:12px;'>Generated: $reportDate</span>
</td></tr>
<tr><td colspan='5' style='text-align:center; border: 2px solid #000000;'>No certificates found.</td></tr>
<tr><td colspan='5' style='border: 2px solid #000000; text-align:center;'>
    <img src='$Client1Logo' style='width:200px; height:50px;'/>
</td></tr>
</table>
"@
}
else {
    $htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr><td colspan='5' style='text-align:center; border: 2px solid #000000;'>
    <img src='$Client2Logo' style='width:200px; height:50px;'/>
</td></tr>
<tr><td colspan='5' style='text-align:center; border: 2px solid #000000; font-weight:bold; font-size:18px;'>
    Applications Report - Certificates
    <br/><span style='font-size:12px;'>Generated: $reportDate</span>
</td></tr>
<tr style='text-align:center; font-weight:bold; background-color:#f0f0f0;'>
    <th style='border: 1px solid #000000;'>App Name</th>
    <th style='border: 1px solid #000000;'>App ID</th>
    <th style='border: 1px solid #000000;'>Certificate ID</th>
    <th style='border: 1px solid #000000;'>Expiration Date</th>
    <th style='border: 1px solid #000000;'>Days Remaining</th>
</tr>
"@
    foreach ($cert in $certList) {
        $color = ""
        if ($cert.DaysRemaining -le $thresholdExpired) {
            $color = "background-color:#B20000;"  # red
        }
        elseif ($cert.DaysRemaining -le $thresholdAlert) {
            $color = "background-color:#FFA500;"  # orange
        }
        elseif ($cert.DaysRemaining -le $thresholdCritical) {
            $color = "background-color:#FFFF00;"  # yellow
        }
        elseif ($cert.DaysRemaining -le $thresholdWarning) {
            $color = "background-color:#a6ff00;"  # light green
        }

        $htmlTable += "<tr style='text-align:center;$color'>" +
        "<td style='border: 1px solid #000000;'>$($cert.AppName)</td>" +
        "<td style='border: 1px solid #000000;'>$($cert.AppId)</td>" +
        "<td style='border: 1px solid #000000;'>$($cert.CertificateId)</td>" +
        "<td style='border: 1px solid #000000;'>$($cert.ExpiresOn.ToString('dd/MM/yyyy HH:mm'))</td>" +
        "<td style='border: 1px solid #000000;'>$($cert.DaysRemaining)</td>" +
        "</tr>"
    }
    $htmlTable += "<tr><td colspan='5' style='border: 2px solid #000000; text-align:center;'>
        <img src='$Client1Logo' style='width:200px; height:50px;'/>
    </td></tr></table>"
}

# Prepare recipients
$toRecipientsArray = @()
foreach ($mail in $recipients) {
    $toRecipientsArray += @{ emailAddress = @{ address = $mail } }
}

# Email subject
$subject = "Azure Report - Application Certificates"

# Send email
if ($useAttachment -eq 1) {
    $htmlBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($htmlTable))

    $emailPayload = @{
        message         = @{
            subject      = $subject
            body         = @{
                contentType = "HTML"
                content     = $htmlTable
            }
            toRecipients = $toRecipientsArray
            attachments  = @(@{
                    '@odata.type' = "#microsoft.graph.fileAttachment"
                    name          = "Certificate_Report.html"
                    contentType   = "text/html"
                    contentBytes  = $htmlBase64
                })
        }
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 5 -Compress
}
else {
    $emailPayload = @{
        message         = @{
            subject      = $subject
            body         = @{
                contentType = "HTML"
                content     = $htmlTable
            }
            toRecipients = $toRecipientsArray
        }
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 4 -Compress
}

Invoke-RestMethod -Method POST `
    -Uri "https://graph.microsoft.com/v1.0/users/$smtpUser/sendMail" `
    -Headers @{ Authorization = "Bearer $accessToken" } `
    -Body $emailPayload `
    -ContentType "application/json; charset=utf-8"
