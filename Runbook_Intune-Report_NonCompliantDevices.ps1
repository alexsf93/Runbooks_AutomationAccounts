<#
.SYNOPSIS
Azure Automation script to report non-compliant devices from Microsoft Intune and optionally notify affected users, using Microsoft Graph API.

.DESCRIPTION
This script performs the following actions:
- Retrieves all non-compliant devices from Intune using Microsoft Graph API.
- Optionally sends notification emails to affected users with guidance.
- Sends an HTML report to administrators, either embedded or as an attachment.
- Includes a configurable corporate logo in the report footer.

.PARAMETER Configurable Variables (Azure Automation Variables):
- GraphClientId           : Application (client) ID
- GraphTenantId           : Tenant (directory) ID
- GraphSecret             : Client secret

.PARAMETER Credential (Azure Automation Credential Asset):
- Correo_No-Reply         : Credential for the email account used to send the report

.PARAMETER Other variables:
- adminRecipients         : Array of email addresses to receive the admin report
- ClientLogo              : URL of the logo image included at the bottom of the report
- useAttachment           : 1 to send report as an attachment, 0 to include in email body only
- notifyUsers             : 1 to notify users about their non-compliant devices, 0 to skip

.API Permissions Required:
- Microsoft Graph API:
  - DeviceManagementManagedDevices.Read.All
  - Mail.Send

.NOTES
- Timezone is set to "Romance Standard Time" (Spain).
- Can be run on a schedule as an Azure Automation Runbook.
#>

# --- Configuration Variables ---
$useAttachment = 0  # 1 to attach HTML to admin email, 0 for inline only
$notifyUsers   = 1  # 1 to notify users, 0 to skip user emails
$ClientLogo    = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"

# --- Authentication and Environment Variables ---
$clientId     = Get-AutomationVariable -Name "GraphClientId"
$tenantId     = Get-AutomationVariable -Name "GraphTenantId"
$clientSecret = Get-AutomationVariable -Name "GraphSecret"
$smtpCredential = Get-AutomationPSCredential -Name "Correo_No-Reply"
$smtpUser = $smtpCredential.UserName

# --- Admin Recipients ---
$adminRecipients = @("admin1@example.com", "admin2@example.com")

# --- Get Microsoft Graph Access Token ---
$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$accessToken = $tokenResponse.access_token
$headers = @{ Authorization = "Bearer $accessToken" }

# --- Get Non-Compliant Devices ---
$devicesResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=complianceState eq 'noncompliant'" -Headers $headers -Method Get
$deviceReport = @()

foreach ($device in $devicesResponse.value) {
    $userName = ""
    $userEmail = ""
    $notifiedStatus = "No"

    try {
        $userResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($device.id)/users" -Headers $headers -Method Get
        $userInfo = $userResponse.value | Select-Object -First 1
        $userName = $userInfo.displayName
        $userEmail = $userInfo.userPrincipalName

        if ($notifyUsers -eq 1 -and $userEmail) {
            $textMessage = @"
Good morning $userName,
We have identified that your corporate device $($device.deviceName) (Serial: $($device.serialNumber)) is currently non-compliant in Intune.
Last sync: $($device.lastSyncDateTime.ToString("dd/MM/yyyy HH:mm"))
Please turn on the device and sync with the company portal.
"@
            $htmlMessage = @"
<p>$($textMessage -replace "`n", "<br>")</p>
<p>
  <a href='https://inkoovadigital-my.sharepoint.com/:v:/g/personal/asuarez_inkoova_com/EToEbrk38IxMuvcYpK6sdjkB4FuNyVkMmL4l-gvbPw795Q?e=goMBpV' target='_blank'>
    <img src='https://i.ibb.co/NnYzp7js/Miniature-portal-de-empresa.png' alt='Video explicativo' style='width:640px; height:auto; display:block; margin-bottom:10px;'>
  </a>
</p>
<p>For assistance, contact <a href='mailto:support@company.com'>Support</a>.</p>
<p><img src='$ClientLogo' alt='Company Logo' style='height:50px; float:left; margin-top:10px;'></p>
"@

            $userEmailBody = @{
                message = @{
                    subject = "ACTION REQUIRED: Your corporate device is non-compliant"
                    body = @{ contentType = "HTML"; content = $htmlMessage }
                    toRecipients = @(@{ emailAddress = @{ address = $userEmail } })
                }
                saveToSentItems = $false
            } | ConvertTo-Json -Depth 4

            Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$smtpUser/sendMail" -Headers @{ Authorization = "Bearer $accessToken" } -Body $userEmailBody -ContentType "application/json; charset=utf-8"
            $notifiedStatus = "Yes"
        } elseif (-not $userEmail) {
            $notifiedStatus = "No email"
        }
    } catch {
        $notifiedStatus = "Error: $_"
    }

    $deviceReport += [PSCustomObject]@{
        Device       = $device.deviceName
        User         = $userName
        Email        = $userEmail
        OS           = $device.operatingSystem
        Serial       = $device.serialNumber
        LastSync     = $device.lastSyncDateTime.ToString("dd/MM/yyyy HH:mm")
        Notified     = $notifiedStatus
    }
}

# --- Build HTML Report ---
$reportDate = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), "Romance Standard Time").ToString("dd/MM/yyyy HH:mm")
$htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr><td colspan='7' style='text-align:center; font-weight:bold; font-size:18px;'>Non-Compliant Device Report<br><span style='font-size:12px;'>Generated: $reportDate</span></td></tr>
<tr style='background-color:#f0f0f0; text-align:center; font-weight:bold;'><th>Device</th><th>User</th><th>Email</th><th>OS</th><th>Serial</th><th>Last Sync</th><th>Notified</th></tr>
"@
foreach ($entry in $deviceReport) {
    $htmlTable += "<tr style='text-align:center;'>" +
        "<td>$($entry.Device)</td><td>$($entry.User)</td><td>$($entry.Email)</td>" +
        "<td>$($entry.OS)</td><td>$($entry.Serial)</td><td>$($entry.LastSync)</td><td>$($entry.Notified)</td></tr>"
}
$htmlTable += "<tr><td colspan='7' style='text-align:center;'><img src='$ClientLogo' style='height:50px;'></td></tr></table>"

# --- Send Admin Report Email ---
$adminMail = @{
    message = @{
        subject = "Intune Report - Non-Compliant Devices"
        body = @{ contentType = "HTML"; content = $htmlTable }
        toRecipients = @($adminRecipients | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
    }
    saveToSentItems = $false
}

if ($useAttachment -eq 1) {
    $filePath = Join-Path -Path $env:TEMP -ChildPath "NonCompliantDevices.html"
    $htmlTable | Out-File -FilePath $filePath -Encoding UTF8
    $fileBytes = [System.IO.File]::ReadAllBytes($filePath)
    $base64 = [Convert]::ToBase64String($fileBytes)
    $adminMail.message.Attachments = @(
        @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            name = "NonCompliantDevices.html"
            contentType = "text/html"
            contentBytes = $base64
        }
    )
}

$adminMailJson = $adminMail | ConvertTo-Json -Depth 4 -Compress
Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$smtpUser/sendMail" -Headers @{ Authorization = "Bearer $accessToken" } -Body $adminMailJson -ContentType "application/json; charset=utf-8"
