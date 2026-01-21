<#
.SYNOPSIS
    Generates an HTML report of App Registration secrets expiration.

.DESCRIPTION
    Retrieves application records and client secrets from Microsoft Graph API.
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
    Start-AutomationRunbook -Name "Runbook_Azure-Report_AppRegistrationsExpirationSecrets"

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

# Recipients array
$recipients = @("example@domain.com", "example1@domain.com")

$useAttachment = 0  # 1 = send HTML attachment + body, 0 = only HTML in body

# Logo URLs
$Client2Logo = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"
$Client1Logo = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"

# Threshold values for row coloring
$expiredThreshold = 0
$warningThreshold = 15
$criticalThreshold = 90
$noticeThreshold = 180

# Get Microsoft Graph token
$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$accessToken = $tokenResponse.access_token
$headers = @{ Authorization = "Bearer $accessToken" }

# Get App Registrations
$appRegs = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/applications?`$top=999" -Headers $headers

$secretsList = @()

foreach ($app in $appRegs.value) {
    if ($app.passwordCredentials) {
        foreach ($secret in $app.passwordCredentials) {
            $expiresOn = [DateTime]$secret.endDateTime
            $daysRemaining = [math]::Floor(($expiresOn - (Get-Date)).TotalDays)

            $secretsList += [PSCustomObject]@{
                AppName       = $app.displayName
                AppId         = $app.appId
                SecretId      = $secret.keyId
                ExpiresOn     = $expiresOn
                DaysRemaining = $daysRemaining
            }
        }
    }
}

# Sort by AppName
$secretsList = $secretsList | Sort-Object AppName

# Get current time in Spain timezone
$reportDate = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), "Romance Standard Time").ToString("dd/MM/yyyy HH:mm")

# Build HTML report
if ($secretsList.Count -eq 0) {
    $htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr><td colspan='5' style='border: 2px solid #000000; text-align:center;'>
    <img src='$Client2Logo' alt='Client2 Logo' style='width:200px; height:50px; margin: 10px auto; display:block;'/>
</td></tr>
<tr><td colspan='5' style='text-align:center; border: 2px solid #000000; font-weight:bold; font-size:18px;'>
    App Registrations Report - Secret Expirations
    <br/><span style='font-size:12px;'>Generated on: $reportDate</span>
</td></tr>
<tr><td colspan='5' style='text-align:center; border: 2px solid #000000;'>No App Registrations with secrets found.</td></tr>
<tr><td colspan='5' style='border: 2px solid #000000; text-align:center;'>
    <img src='$Client1Logo' alt='Client1 Logo' style='width:200px; height:50px; margin: 10px auto; display:block;'/>
</td></tr>
</table>
"@
}
else {
    $htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr><td colspan='5' style='border: 2px solid #000000; text-align:center;'>
    <img src='$Client2Logo' alt='Client2 Logo' style='width:200px; height:50px; margin: 10px auto; display:block;'/>
</td></tr>
<tr><td colspan='5' style='text-align:center; border: 2px solid #000000; font-weight:bold; font-size:18px;'>
    App Registrations Report - Secret Expirations
    <br/><span style='font-size:12px;'>Generated on: $reportDate</span>
</td></tr>
<tr style='text-align:center; font-weight:bold; background-color:#f0f0f0;'>
    <th style='border: 1px solid #000000;'>App Name</th>
    <th style='border: 1px solid #000000;'>App ID</th>
    <th style='border: 1px solid #000000;'>Secret ID</th>
    <th style='border: 1px solid #000000;'>Expiration Date</th>
    <th style='border: 1px solid #000000;'>Days Remaining</th>
</tr>
"@
    foreach ($secret in $secretsList) {
        $color = ""
        if ($secret.DaysRemaining -le $expiredThreshold) {
            $color = "background-color:#B20000;"  # red
        }
        elseif ($secret.DaysRemaining -le $warningThreshold) {
            $color = "background-color:#FFA500;"  # orange
        }
        elseif ($secret.DaysRemaining -le $criticalThreshold) {
            $color = "background-color:#FFFF00;"  # yellow
        }
        elseif ($secret.DaysRemaining -le $noticeThreshold) {
            $color = "background-color:#a6ff00;"  # bright green
        }

        $htmlTable += "<tr style='text-align:center;$color'>" +
        "<td style='border: 1px solid #000000;'>$($secret.AppName)</td>" +
        "<td style='border: 1px solid #000000;'>$($secret.AppId)</td>" +
        "<td style='border: 1px solid #000000;'>$($secret.SecretId)</td>" +
        "<td style='border: 1px solid #000000;'>$($secret.ExpiresOn.ToString('dd/MM/yyyy HH:mm'))</td>" +
        "<td style='border: 1px solid #000000;'>$($secret.DaysRemaining)</td>" +
        "</tr>"
    }
    $htmlTable += "<tr><td colspan='5' style='border: 2px solid #000000; text-align:center;'>
        <img src='$Client1Logo' alt='Client1 Logo' style='width:200px; height:50px; margin: 10px auto; display:block;'/>
    </td></tr>"
    $htmlTable += "</table>"
}

# Prepare recipients for email payload
$toRecipientsArray = @()
foreach ($mail in $recipients) {
    $toRecipientsArray += @{ emailAddress = @{ address = $mail } }
}

# Email subject
$subject = "Azure Report - App Registrations Secrets"

# Prepare and send email via Microsoft Graph API
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
                    name          = "AppRegistrations_Secrets_Report.html"
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
