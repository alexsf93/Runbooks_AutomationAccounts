<#
.SYNOPSIS
    Generates a report of users' MFA status (Enabled/Disabled) in Azure AD.

.DESCRIPTION
    Retrieves users and checks their multi-factor authentication methods.
    Generates a color-coded HTML report (Green=Enabled, Orange=Not Enabled, Red=Error) and sends it via email.

.PARAMETER GraphClientId
    Application (client) ID registered in Azure AD.

.PARAMETER GraphTenantId
    Directory (tenant) ID in Azure AD.

.PARAMETER GraphSecret
    Application secret.

.PARAMETER No-Reply Email
    SMTP credentials (PSCredential).

.EXAMPLE
    # Run as Azure Automation Runbook
    Start-AutomationRunbook -Name "Runbook_EntraID-Report_MFA_Enabled_Disabled"

.NOTES
    Author: Automated System
    Version: 1.0
    Prerequisites: AuthenticationMethod.Read.All, Mail.Send
#>

# Configurable variables (same as your script)
$clientId = Get-AutomationVariable -Name "GraphClientId"
$tenantId = Get-AutomationVariable -Name "GraphTenantId"
$clientSecret = Get-AutomationVariable -Name "GraphSecret"

$smtpCredential = Get-AutomationPSCredential -Name "No-Reply Email"
$smtpUser = $smtpCredential.UserName

# Recipients list
$recipients = @("example1@domain.com", "example2@domain.com")

$useAttachment = 0  # 1 = attachment + body, 0 = only body

# Logos (you can change if you want)
$Client2Logo = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"
$Client1Logo = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"

# Get access token
$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$accessToken = $tokenResponse.access_token
$headers = @{ Authorization = "Bearer $accessToken" }

# Get users (up to 999)
$users = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users?`$top=999" -Headers $headers

# Get current date/time in Spain timezone
$reportDate = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), "Romance Standard Time").ToString("dd/MM/yyyy HH:mm")

# Check MFA status for each user
$mfaStatusList = @()

foreach ($user in $users.value) {
    $userId = $user.id
    $userDisplayName = $user.displayName
    $userPrincipalName = $user.userPrincipalName

    $methodsUri = "https://graph.microsoft.com/v1.0/users/$userId/authentication/methods"

    try {
        $methods = Invoke-RestMethod -Uri $methodsUri -Headers $headers

        # Look for common MFA methods
        $mfaMethods = $methods.value | Where-Object {
            $_.'@odata.type' -in @(
                "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod",
                "#microsoft.graph.phoneAuthenticationMethod"
            )
        }

        $mfaStatus = if ($mfaMethods.Count -gt 0) { "Enabled" } else { "Not Enabled" }
    }
    catch {
        $mfaStatus = "Error querying status"
    }

    $mfaStatusList += [PSCustomObject]@{
        DisplayName       = $userDisplayName
        UserPrincipalName = $userPrincipalName
        MFAStatus         = $mfaStatus
    }
}

# Build the HTML report
if ($mfaStatusList.Count -eq 0) {
    $htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr><td colspan='3' style='border: 2px solid #000000; text-align:center;'>
    <img src='$Client2Logo' style='width:200px; height:50px;'/>
</td></tr>
<tr><td colspan='3' style='text-align:center; border: 2px solid #000000; font-weight:bold; font-size:18px;'>
    MFA Status Report for Users
    <br/><span style='font-size:12px;'>Generated: $reportDate</span>
</td></tr>
<tr><td colspan='3' style='text-align:center; border: 2px solid #000000;'>No users found.</td></tr>
<tr><td colspan='3' style='border: 2px solid #000000; text-align:center;'>
    <img src='$Client1Logo' style='width:200px; height:50px;'/>
</td></tr>
</table>
"@
}
else {
    $htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr><td colspan='3' style='text-align:center; border: 2px solid #000000;'>
    <img src='$Client2Logo' style='width:200px; height:50px;'/>
</td></tr>
<tr><td colspan='3' style='text-align:center; border: 2px solid #000000; font-weight:bold; font-size:18px;'>
    MFA Status Report for Users
    <br/><span style='font-size:12px;'>Generated: $reportDate</span>
</td></tr>
<tr style='text-align:center; font-weight:bold; background-color:#f0f0f0;'>
    <th style='border: 1px solid #000000;'>User Name</th>
    <th style='border: 1px solid #000000;'>User Principal Name</th>
    <th style='border: 1px solid #000000;'>MFA Status</th>
</tr>
"@

    foreach ($entry in $mfaStatusList) {
        $color = ""
        switch ($entry.MFAStatus) {
            "Enabled" { $color = "background-color:#07a500;" }  # green
            "Not Enabled" { $color = "background-color:#ff8000;" }  # orange
            default { $color = "background-color:#B20000;" }  # red
        }

        $htmlTable += "<tr style='text-align:center;$color'>" +
        "<td style='border: 1px solid #000000;'>$($entry.DisplayName)</td>" +
        "<td style='border: 1px solid #000000;'>$($entry.UserPrincipalName)</td>" +
        "<td style='border: 1px solid #000000;'>$($entry.MFAStatus)</td>" +
        "</tr>"
    }

    $htmlTable += "<tr><td colspan='3' style='border: 2px solid #000000; text-align:center;'>
        <img src='$Client1Logo' style='width:200px; height:50px;'/>
    </td></tr></table>"
}

# Prepare recipients
$toRecipientsArray = @()
foreach ($mail in $recipients) {
    $toRecipientsArray += @{ emailAddress = @{ address = $mail } }
}

# Email subject
$subject = "Azure Report - Users MFA Status"

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
                    name          = "MFA_Status_Report_Users.html"
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
