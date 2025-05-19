<#
.SYNOPSIS
Azure Automation script to identify and block inactive member users in Azure AD, and send a detailed HTML report via email using Microsoft Graph API.

.DESCRIPTION
This script performs the following actions:
- Retrieves all active "Member" type users from Azure AD using Microsoft Graph API.
- Checks sign-in activity and detects users who have not signed in for a configurable number of days.
- Blocks inactive users by setting `accountEnabled` to false.
- Builds an HTML report listing affected users and their last sign-in date.
- Sends the report to specified recipients via email, either embedded or as an attachment.
- Optionally includes corporate logos in the report header and footer.

.PARAMETER Configurable Variables (Azure Automation Variables):
- GraphClientId           : Application (client) ID
- GraphTenantId           : Tenant (directory) ID
- GraphSecret             : Client secret

.PARAMETER Credential (Azure Automation Credential Asset):
- Correo_No-Reply         : Credential for the email account used to send the report

.PARAMETER Other variables:
- recipient               : List of email addresses to receive the report
- ClientLogo1             : URL of the logo to display in the report header
- ClientLogo2             : URL of the logo to display in the report footer
- useAttachment           : 1 to send the report as an attachment, 0 to include it in the body only
- inactivityDays          : Number of days without activity to consider a user inactive

.API Permissions Required:
- Microsoft Graph API:
  - User.Read.All
  - AuditLog.Read.All
  - User.EnableDisableAccount.All
  - Mail.Send

.NOTES
- Time zone used: "Romance Standard Time" (Spain).
- Can be scheduled as an Azure Automation Runbook.
- Users whose UPN contains '#' are excluded from the analysis.
#>

# Configurable variables
$clientId     = Get-AutomationVariable -Name "GraphClientId"
$tenantId     = Get-AutomationVariable -Name "GraphTenantId"
$clientSecret = Get-AutomationVariable -Name "GraphSecret"

$smtpCredential = Get-AutomationPSCredential -Name "Correo_No-Reply"
$smtpUser       = $smtpCredential.UserName

# Recipients
$recipient = @("alexsf93@gmail.com", "example@domain.com")

# Send report as attachment (1) or embed in email body (0)
$useAttachment = 0

# Logo URLs
$ClientLogo1 = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"
$ClientLogo2 = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"

# Inactivity threshold (in days)
$inactivityDays = 2
$cutoffDate = (Get-Date).AddDays(-$inactivityDays)

# Get Microsoft Graph token
$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$accessToken = $tokenResponse.access_token
$headers = @{ Authorization = "Bearer $accessToken"; "ConsistencyLevel" = "eventual" }

# Get all active Member users
$uriUsers = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Member' and accountEnabled eq true&`$select=id,displayName,userPrincipalName,accountEnabled&`$top=999"
$users = @()
do {
    $response = Invoke-RestMethod -Method Get -Uri $uriUsers -Headers $headers
    $users += $response.value
    $uriUsers = $response.'@odata.nextLink'
} while ($uriUsers)

# Filter users
$inactiveUsers = @()

foreach ($user in $users) {
    if ($user.userPrincipalName -like "*#*") { continue }

    $uriSignIns = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=userId eq '$($user.id)'&`$orderby=createdDateTime desc&`$top=50"
    try {
        $signIns = Invoke-RestMethod -Method Get -Uri $uriSignIns -Headers $headers -ErrorAction Stop
        $successfulSignIns = $signIns.value | Where-Object { $_.status.errorCode -eq 0 }

        if ($successfulSignIns.Count -gt 0) {
            $lastSignIn = ($successfulSignIns | Select-Object -First 1).createdDateTime
            if ([datetime]$lastSignIn -lt $cutoffDate) {
                $inactiveUsers += [PSCustomObject]@{
                    Id               = $user.id
                    DisplayName      = $user.displayName
                    UserPrincipalName= $user.userPrincipalName
                    LastSignIn       = [datetime]$lastSignIn
                    AccountEnabled   = $user.accountEnabled
                }
            }
        }
    } catch {
        # Ignore errors on sign-in query
    }
}

# Block users and update state
foreach ($u in $inactiveUsers) {
    if ($u.AccountEnabled) {
        $patchBody = @{ accountEnabled = $false } | ConvertTo-Json
        try {
            Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($u.Id)" -Headers $headers -Method PATCH -Body $patchBody -ContentType "application/json"
            $u.AccountEnabled = $true  # Was active and has now been blocked
        } catch {
            $u.AccountEnabled = $true
        }
    } else {
        $u.AccountEnabled = $false  # Already blocked
    }
}

# Get current time in Spain
$timestamp = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), "Romance Standard Time").ToString("dd/MM/yyyy HH:mm")

# Build HTML table
if ($inactiveUsers.Count -eq 0) {
    $htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr><td colspan='4' style='border: 2px solid #000000; text-align:center;'>
    <img src='$ClientLogo1' alt='Client Logo' style='width:200px; height:50px; margin: 10px auto; display:block;'/>
</td></tr>
<tr><td colspan='4' style='text-align:center; border: 2px solid #000000; font-weight:bold; font-size:18px;'>
    Inactive and Blocked Member Users Report
    <br/><span style='font-size:12px;'>Report generated: $timestamp</span>
</td></tr>
<tr><td colspan='4' style='text-align:center; border: 2px solid #000000;'>No inactive users found (more than $inactivityDays days without login).</td></tr>
<tr><td colspan='4' style='border: 2px solid #000000; text-align:center;'>
    <img src='$ClientLogo2' alt='Client Logo' style='width:200px; height:50px; margin: 10px auto; display:block;'/>
</td></tr>
</table>
"@
} else {
    $htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr><td colspan='4' style='border: 2px solid #000000; text-align:center;'>
    <img src='$ClientLogo1' alt='Client Logo' style='width:200px; height:50px; margin: 10px auto; display:block;'/>
</td></tr>
<tr><td colspan='4' style='text-align:center; border: 2px solid #000000; font-weight:bold; font-size:18px;'>
    Inactive and Blocked Member Users Report
    <br/><span style='font-size:12px;'>Report generated: $timestamp</span>
</td></tr>
<tr style='text-align:center; font-weight:bold; background-color:#f0f0f0;'>
    <th style='border: 1px solid #000000;'>Display Name</th>
    <th style='border: 1px solid #000000;'>User Principal Name</th>
    <th style='border: 1px solid #000000;'>Last Sign-In</th>
    <th style='border: 1px solid #000000;'>New Status</th>
</tr>
"@
    foreach ($u in $inactiveUsers) {
        $dateStr = $u.LastSignIn.ToString("dd/MM/yyyy HH:mm")
        $newStatus = if ($u.AccountEnabled) { "Blocked now" } else { "Already blocked" }
        $htmlTable += "<tr style='text-align:center;'>" +
            "<td style='border: 1px solid #000000;'>$($u.DisplayName)</td>" +
            "<td style='border: 1px solid #000000;'>$($u.UserPrincipalName)</td>" +
            "<td style='border: 1px solid #000000;'>$dateStr</td>" +
            "<td style='border: 1px solid #000000;'>$newStatus</td>" +
            "</tr>"
    }
    $htmlTable += "<tr><td colspan='4' style='border: 2px solid #000000; text-align:center;'>
        <img src='$ClientLogo2' alt='Client Logo' style='width:200px; height:50px; margin: 10px auto; display:block;'/>
    </td></tr>"
    $htmlTable += "</table>"
}

# Prepare recipients
$toRecipientsArray = @()
foreach ($mail in $recipient) {
    $toRecipientsArray += @{ emailAddress = @{ address = $mail } }
}

# Build email payload
$subject = "Azure Report - Inactive and Blocked Member Users"

if ($useAttachment -eq 1) {
    $htmlBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($htmlTable))

    $emailPayload = @{
        message = @{
            subject = $subject
            body = @{
                contentType = "HTML"
                content     = $htmlTable
            }
            toRecipients = $toRecipientsArray
            attachments = @(@{
                '@odata.type' = "#microsoft.graph.fileAttachment"
                name          = "InactiveUsersReport.html"
                contentType   = "text/html"
                contentBytes  = $htmlBase64
            })
        }
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 5 -Compress
} else {
    $emailPayload = @{
        message = @{
            subject = $subject
            body = @{
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
