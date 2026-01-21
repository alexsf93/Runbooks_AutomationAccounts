<#
.SYNOPSIS
    Blocks Azure AD users who haven't logged in recently.

.DESCRIPTION
    Retrieves active Member users and checks their last sign-in date.
    If the last sign-in is older than the threshold, the user is blocked (or simulated).
    Generates an HTML report and sends it via email.
    Excludes high-privilege users and specific excluded accounts.

.PARAMETER GraphClientId
    Application (client) ID registered in Azure AD.

.PARAMETER GraphTenantId
    Directory (tenant) ID in Azure AD.

.PARAMETER GraphSecret
    Application secret.

.PARAMETER Email
    SMTP credentials (PSCredential).

.EXAMPLE
    # Run as Azure Automation Runbook
    Start-AutomationRunbook -Name "Runbook_EntraID-Report_DisableUsersInactivty"

.NOTES
    Author: Automated System
    Version: 1.0
    Prerequisites: User.Read.All, Directory.Read.All, AuditLog.Read.All, Mail.Send, User.EnableDisableAccount.All
#>

# Configurable variables
$clientId = Get-AutomationVariable -Name "GraphClientId"
$tenantId = Get-AutomationVariable -Name "GraphTenantId"
$clientSecret = Get-AutomationVariable -Name "GraphSecret"
$smtpCredential = Get-AutomationPSCredential -Name "Email"
$smtpUser = $smtpCredential.UserName
$recipient = @("example@domain.com", "example1@domain.com")
$useAttachment = 0
$ClientLogo1 = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"
$ClientLogo2 = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"
$inactivityDays = 30
$blockUsers = 1  # 1 = block users, 0 = simulate only
$reportHighPrivilegeUsers = 0 # 1 = include high-privilege users in report, 0 = exclude them
$cutoffDate = (Get-Date).AddDays(-$inactivityDays)

# Static list of users to exclude from blocking and report (userPrincipalName)
$excludedUsers = @(
    "no-reply@karanai102.es",
    "role@karanai102.es"
)

# Obtain Microsoft Graph token
$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$accessToken = $tokenResponse.access_token
$headers = @{ Authorization = "Bearer $accessToken"; "ConsistencyLevel" = "eventual" }

# Get active directory roles
$uriDirectoryRoles = "https://graph.microsoft.com/v1.0/directoryRoles"
$directoryRoles = Invoke-RestMethod -Uri $uriDirectoryRoles -Headers $headers -Method Get

# High privilege roles to detect
$highPrivilegeRoleNames = @(
    "Global Administrator",
    "Privileged Role Administrator",
    "Security Administrator",
    "Exchange Administrator",
    "SharePoint Administrator",
    "User Administrator"
)

# Filter active roles that are considered privileged
$activeHighRoles = $directoryRoles.value | Where-Object { $highPrivilegeRoleNames -contains $_.displayName }

# Get users with high privilege roles
$privilegedUsers = @{}
foreach ($role in $activeHighRoles) {
    $roleMembersUri = "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)/members?`$select=id,userPrincipalName"
    $membersResponse = Invoke-RestMethod -Uri $roleMembersUri -Headers $headers -Method Get
    foreach ($member in $membersResponse.value) {
        $privilegedUsers[$member.id] = $member.userPrincipalName
    }
}

# Get active Member-type users
$uriUsers = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Member' and accountEnabled eq true&`$select=id,displayName,userPrincipalName,accountEnabled&`$top=999"
$users = @()
do {
    $response = Invoke-RestMethod -Method Get -Uri $uriUsers -Headers $headers
    $users += $response.value
    $uriUsers = $response.'@odata.nextLink'
} while ($uriUsers)

# Get service principals to exclude them
$servicePrincipals = @()
$uriSP = "https://graph.microsoft.com/v1.0/servicePrincipals?`$select=id"
do {
    $spResponse = Invoke-RestMethod -Method Get -Uri $uriSP -Headers $headers
    $servicePrincipals += $spResponse.value
    $uriSP = $spResponse.'@odata.nextLink'
} while ($uriSP)

# Process users
$processedUsers = @()
foreach ($user in $users) {
    # Exclude special users
    if ($user.userPrincipalName -like "*#*") { continue }
    if ($servicePrincipals | Where-Object { $_.id -eq $user.id }) { continue }
    if ($excludedUsers -contains $user.userPrincipalName) { continue }

    # Get last sign-in
    $uriSignIns = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=userId eq '$($user.id)'&`$orderby=createdDateTime desc&`$top=1"
    try {
        $signIns = Invoke-RestMethod -Method Get -Uri $uriSignIns -Headers $headers
        if ($signIns.value.Count -eq 0) { continue }
        $lastSignIn = [datetime]$signIns.value[0].createdDateTime

        if ($lastSignIn -lt $cutoffDate) {
            # High-privilege user
            if ($privilegedUsers.ContainsKey($user.id)) {
                if ($reportHighPrivilegeUsers -eq 1) {
                    $status = "User not blocked, high privileges"
                    $processedUsers += [PSCustomObject]@{
                        Id                = $user.id
                        DisplayName       = $user.displayName
                        UserPrincipalName = $user.userPrincipalName
                        LastSignIn        = $lastSignIn
                        Status            = $status
                    }
                }
                # If reportHighPrivilegeUsers = 0, skip entirely
            }
            else {
                # User without high privileges
                if ($blockUsers -eq 1) {
                    # Block user
                    $patchBody = @{ accountEnabled = $false } | ConvertTo-Json
                    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($user.id)" -Headers $headers -Method PATCH -Body $patchBody -ContentType "application/json"
                    $blockStatus = "User has been blocked"
                }
                else {
                    $blockStatus = "Simulation: user not blocked"
                }
                $processedUsers += [PSCustomObject]@{
                    Id                = $user.id
                    DisplayName       = $user.displayName
                    UserPrincipalName = $user.userPrincipalName
                    LastSignIn        = $lastSignIn
                    Status            = $blockStatus
                }
            }
        }
    }
    catch {
        Write-Output "Error processing user $($user.userPrincipalName): $_"
    }
}

# Report generation timestamp
$timestamp = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), "Romance Standard Time").ToString("dd/MM/yyyy HH:mm")

# Build HTML report
$htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr><td colspan='4' style='text-align:center; border: 2px solid #000000;'>
    <img src='$ClientLogo2' style='width:200px; height:50px;'/>
</td></tr>
<tr><td colspan='4' style='text-align:center; border: 2px solid #000000; font-weight:bold; font-size:18px;'>
    Inactive and Blocked Users Report<br/>
    <span style='font-size:12px;'>Generated: $timestamp</span><br/>
"@

if ($blockUsers -eq 0) {
    $htmlTable += "<span style='color:blue; font-weight:bold;'>This is a simulation report. No users have been blocked.</span><br/>"
}

$htmlTable += "</td></tr>"

if ($processedUsers.Count -eq 0) {
    $htmlTable += "<tr><td colspan='4' style='text-align:center;'>No inactive users found for blocking.</td></tr>"
}
else {
    $htmlTable += "<tr style='text-align:center; font-weight:bold; background-color:#f0f0f0;'>
        <th>DisplayName</th><th>UserPrincipalName</th><th>Last Sign-In</th><th>Status</th></tr>"
    foreach ($u in $processedUsers) {
        $dateStr = if ($u.LastSignIn -is [datetime]) { $u.LastSignIn.ToString("dd/MM/yyyy HH:mm") } else { $u.LastSignIn }
        $htmlTable += "<tr style='text-align:center;'>
            <td>$($u.DisplayName)</td>
            <td>$($u.UserPrincipalName)</td>
            <td>$dateStr</td>
            <td>$($u.Status)</td>
        </tr>"
    }
}

$htmlTable += "<tr><td colspan='4' style='text-align:center; border: 2px solid #000000;'>
    <img src='$ClientLogo1' style='width:200px; height:50px;'/>
</td></tr></table>"

# Prepare email
$toRecipientsArray = @()
foreach ($mail in $recipient) {
    $toRecipientsArray += @{ emailAddress = @{ address = $mail } }
}

$subject = if ($blockUsers -eq 1) { "Azure Report - Inactive Users Blocked" } else { "Azure Report - User Blocking Simulation" }

if ($useAttachment -eq 1) {
    $htmlBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($htmlTable))
    $emailPayload = @{
        message         = @{
            subject      = $subject
            body         = @{ contentType = "HTML"; content = $htmlTable }
            toRecipients = $toRecipientsArray
            attachments  = @(@{
                    '@odata.type' = "#microsoft.graph.fileAttachment"
                    name          = "InactiveUsersReport.html"
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
            body         = @{ contentType = "HTML"; content = $htmlTable }
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
