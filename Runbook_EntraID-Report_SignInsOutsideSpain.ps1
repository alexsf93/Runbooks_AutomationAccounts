<#
.SYNOPSIS
    Generates an HTML report of Sign-Ins performed outside of Spain.

.DESCRIPTION
    This script queries Microsoft Entra ID (Azure AD) sign-in audit logs
    using the Microsoft Graph API, looking for recent events (last 7 days by default).
    It filters events to find those originating from locations outside of Spain ("ES" or "Spain").
    It allows defining custom exclusions (e.g. key users or specific countries).
    Finally, it generates an HTML report grouped by User and Country, including dates and locations.
    The report is sent via email using Microsoft Graph.

.PARAMETER GraphClientId
    Application (client) ID registered in Azure AD.

.PARAMETER GraphTenantId
    Directory (tenant) ID in Azure AD.

.PARAMETER GraphSecret
    Application secret.

.PARAMETER Correo_No-Reply
    SMTP credentials (PSCredential) used as the sender of the report.

.EXAMPLE
    # Run as Azure Automation Runbook
    Start-AutomationRunbook -Name "Runbook_EntraID-Report_SignInsOutsideSpain"

.NOTES
    Author: Alejandro Suárez @alexsf93
    Version: 1.0
    Prerequisites: Required Graph API permission -> AuditLog.Read.All, Mail.Send 
#>

# =====================================================================
# CONFIGURATION AND EXCLUSIONS
# =====================================================================

# Email Settings
$recipients = @("admin1@yourdomain.com", "security@yourdomain.com")
$useAttachment = 1  # 1 = attachment + body, 0 = body only

# Logos (Standard Reporting)
$Client1Logo = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"
$Client2Logo = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"

# List of users to exclude (UserPrincipalName)
$ExcludedUsers = @(
    # "admin@yourdomain.com",
    # "service@yourdomain.com"
)

# List of countries to exclude (besides Spain). Use the ISO code or name returned by Graph.
$ExcludedCountries = @(
    "ES",
    "Spain"
)

# Time range to search sign-ins (last 7 days by default)
$DaysToSearch = 7
$StartDate = (Get-Date).ToUniversalTime().AddDays(-$DaysToSearch).ToString("yyyy-MM-ddTHH:mm:ssZ")

# Generate dynamic name with Year and Week of the Year
$Date = Get-Date
$Year = $Date.Year
$WeekNumber = (Get-Culture).Calendar.GetWeekOfYear($Date, [System.Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday)
$ReportName = "SignIns_OutsideSpain_Report_${Year}-W${WeekNumber}.html"
$ReportPath = Join-Path -Path $env:TEMP -ChildPath $ReportName

# =====================================================================
# GRAPH API AUTHENTICATION
# =====================================================================

# Retrieve variables from Automation Account
$clientId = Get-AutomationVariable -Name "GraphClientId"
$tenantId = Get-AutomationVariable -Name "GraphTenantId"
$clientSecret = Get-AutomationVariable -Name "GraphSecret"

$smtpCredential = Get-AutomationPSCredential -Name "Correo_No-Reply"
$smtpUser = $smtpCredential.UserName

# Scopes and authentication endpoint
$scopes = "https://graph.microsoft.com/.default"
$tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# Obtain access token
$body = @{
    client_id     = $clientId
    scope         = $scopes
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUrl -ContentType "application/x-www-form-urlencoded" -Body $body
$accessToken = $tokenResponse.access_token

# =====================================================================
# HELPER FUNCTIONS
# =====================================================================

function Invoke-GraphGet {
    param (
        [string]$uri
    )
    return Invoke-RestMethod -Headers @{ Authorization = "Bearer $accessToken" } -Uri $uri -Method Get
}

# =====================================================================
# DATA COLLECTION AND FILTERING
# =====================================================================

Write-Output "Retrieving Sign-in logs since $StartDate..."

# Base request to fetch recent events
$filter = "createdDateTime ge $StartDate"
$baseUri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=$filter&`$top=1000"
$nextLink = $baseUri

$allSignIns = @()

# Loop with pagination to get all results according to the date filter
do {
    $response = Invoke-GraphGet -uri $nextLink

    if ($null -eq $response.value -or $response.value.Count -eq 0) {
        break
    }

    $allSignIns += $response.value

    $nextLink = $null
    if ($response.'@odata.nextLink') {
        $nextLink = $response.'@odata.nextLink'
    }
} while ($nextLink)

Write-Output "$($allSignIns.Count) total sign-in records found. Filtering those located outside of Spain..."

# =====================================================================
# APPLY EXCLUSIONS
# =====================================================================

$suspiciousSignIns = $allSignIns | Where-Object {
    $country = $_.location.countryOrRegion
    $user = $_.userPrincipalName
    
    # Check if the sign-in has a country and is NOT in the excluded list (like ES or Spain)
    $isOutsideSpainBox = (-not [string]::IsNullOrEmpty($country)) -and ($ExcludedCountries -notcontains $country)
    
    # Check if the user is NOT in the exclusion list
    $isNotExcludedUser = ($ExcludedUsers -notcontains $user)
    
    return ($isOutsideSpainBox -and $isNotExcludedUser)
}

# =====================================================================
# HTML REPORT GENERATION & EMAIL SENDING
# =====================================================================

$reportDate = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), "Romance Standard Time").ToString("dd/MM/yyyy HH:mm")

if ($suspiciousSignIns.Count -eq 0) {
    $HTML = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr><td colspan='4' style='border: 2px solid #000000; text-align:center;'>
    <img src='$Client2Logo' style='width:200px; height:50px;'/>
</td></tr>
<tr><td colspan='4' style='text-align:center; border: 2px solid #000000; font-weight:bold; font-size:18px;'>
    Sign-Ins Report Outside of Spain
    <br/><span style='font-size:12px;'>Generated: $reportDate</span>
</td></tr>
<tr><td colspan='4' style='text-align:center; border: 2px solid #000000;'>No suspicious sign-ins found outside of Spain in the last $DaysToSearch days.</td></tr>
<tr><td colspan='4' style='border: 2px solid #000000; text-align:center;'>
    <img src='$Client1Logo' style='width:200px; height:50px;'/>
</td></tr>
</table>
"@
}
else {
    $HTML = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr><td colspan='4' style='text-align:center; border: 2px solid #000000;'>
    <img src='$Client2Logo' style='width:200px; height:50px;'/>
</td></tr>
<tr><td colspan='4' style='text-align:center; border: 2px solid #000000; font-weight:bold; font-size:18px;'>
    Sign-Ins Report Outside of Spain
    <br/><span style='font-size:12px;'>Generated: $reportDate</span>
</td></tr>
<tr style='text-align:center; font-weight:bold; background-color:#f0f0f0;'>
    <th style='border: 1px solid #000000;'>User & Country</th>
    <th style='border: 1px solid #000000;'>Connection Date (UTC)</th>
    <th style='border: 1px solid #000000;'>IP Address</th>
    <th style='border: 1px solid #000000;'>App & Client Status</th>
</tr>
"@
    
    # Group results by UPN (User) and Country
    $groupedData = $suspiciousSignIns | Group-Object -Property userPrincipalName, @{Expression = { $_.location.countryOrRegion } }

    foreach ($group in $groupedData) {
        $user = $group.Values[0]
        $country = $group.Values[1]

        $HTML += "<tr><td colspan='4' style='background-color:#d9e1f2; border: 1px solid #000000; padding:8px; font-weight:bold; text-align:left;'>User: $user | Connection Country: $country</td></tr>"

        foreach ($signInEvent in $group.Group | Sort-Object createdDateTime -Descending) {
            $time = $signInEvent.createdDateTime
            $ip = $signInEvent.ipAddress
            $app = $signInEvent.appDisplayName
            $clientApp = $signInEvent.clientAppUsed
            
            $HTML += "<tr style='text-align:center;'>" +
            "<td style='border: 1px solid #000000;'></td>" +
            "<td style='border: 1px solid #000000;'>$time</td>" +
            "<td style='border: 1px solid #000000;'>$ip</td>" +
            "<td style='border: 1px solid #000000;'>$app / $clientApp</td>" +
            "</tr>"
        }
    }
    
    $HTML += "<tr><td colspan='4' style='border: 2px solid #000000; text-align:center;'>
        <img src='$Client1Logo' style='width:200px; height:50px;'/>
    </td></tr></table>"
}

# Save the file locally
$HTML | Out-File -FilePath $ReportPath -Encoding UTF8
    
Write-Output "Report generation finished. HTML structure mapped."
    
# =====================================================================
# SEND EMAIL
# =====================================================================
Write-Output "Preparing outbox email to send report..."

# Prepare recipients
$toRecipientsArray = @()
foreach ($mail in $recipients) {
    $toRecipientsArray += @{ emailAddress = @{ address = $mail } }
}

# Email subject
$subject = "Azure Report - Sign-Ins Outside Spain [$Year-W$WeekNumber]"

# Send email
if ($useAttachment -eq 1) {
    $htmlBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($HTML))

    $emailPayload = @{
        message         = @{
            subject      = $subject
            body         = @{
                contentType = "HTML"
                content     = $HTML
            }
            toRecipients = $toRecipientsArray
            attachments  = @(@{
                    '@odata.type' = "#microsoft.graph.fileAttachment"
                    name          = $ReportName
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
                content     = $HTML
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

Write-Output "Report sent successfully to designated recipients."
