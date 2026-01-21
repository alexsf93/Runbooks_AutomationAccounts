<#
.SYNOPSIS
    Generates a Microsoft 365 license availability report.

.DESCRIPTION
    Retrieves license consumption and availability details from Microsoft 365.
    Generates an HTML report highlighting negative availability and sends it via email.

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
    Start-AutomationRunbook -Name "Runbook_EntraID-Report_AvailableLicenses"

.NOTES
    Author: Automated System
    Version: 1.0
    Prerequisites: Directory.Read.All, Mail.Send
#>

# Configuration variables
$clientId = Get-AutomationVariable -Name "GraphClientId"
$tenantId = Get-AutomationVariable -Name "GraphTenantId"
$clientSecret = Get-AutomationVariable -Name "GraphSecret"
$useAttachment = 0

$smtpCredential = Get-AutomationPSCredential -Name "Correo_No-Reply"
$smtpUser = $smtpCredential.UserName

$recipients = @("example@domain.com", "example1@domain.com")

# Logos URLs as variables
$Client1Logo = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"
$Client2Logo = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"

# Get Spain timezone and current date adjusted
$spainTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Romance Standard Time")
$currentUTC = [DateTime]::UtcNow
$currentDate = [System.TimeZoneInfo]::ConvertTimeFromUtc($currentUTC, $spainTimeZone)

# Acquire access token
$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$accessToken = $tokenResponse.access_token

# Query subscribed SKUs/licenses
$headers = @{ Authorization = "Bearer $accessToken" }
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -Headers $headers -Method GET

# License SKU code to friendly name mapping
$skuNamesMap = @{
    "ENTERPRISEPREMIUM"              = "Microsoft 365 E5"
    "BUSINESS_PREMIUM"               = "Microsoft 365 Business Premium"
    "DEVELOPERPACK_E5"               = "Developer Pack E5"
    "FLOW_FREE"                      = "Power Automate Free"
    "ENTERPRISEPACK"                 = "Microsoft 365 E3"
    "STANDARDPACK"                   = "Microsoft 365 E1"
    "DESKLESSPACK"                   = "Microsoft 365 F3"
    "O365_BUSINESS_ESSENTIALS"       = "Microsoft 365 Business Basic"
    "O365_BUSINESS_PREMIUM"          = "Microsoft 365 Business Standard"
    "SPB"                            = "Microsoft 365 Business Premium"
    "SPE_E3"                         = "Microsoft 365 E3"
    "SPE_E5"                         = "Microsoft 365 E5"
    "AAD_BASIC"                      = "Azure Active Directory Basic"
    "AAD_PREMIUM"                    = "Azure AD Premium P1"
    "AAD_PREMIUM_P2"                 = "Azure AD Premium P2"
    "EMS"                            = "EMS E3"
    "EMSPREMIUM"                     = "EMS E5"
    "EXCHANGESTANDARD"               = "Exchange Online Plan 1"
    "EXCHANGEENTERPRISE"             = "Exchange Online Plan 2"
    "WACONEDRIVESTANDARD"            = "OneDrive for Business Plan 1"
    "WACONEDRIVEENTERPRISE"          = "OneDrive for Business Plan 2"
    "SHAREPOINTSTANDARD"             = "SharePoint Online Plan 1"
    "SHAREPOINTENTERPRISE"           = "SharePoint Online Plan 2"
    "POWER_BI_STANDARD"              = "Power BI Free"
    "POWER_BI_PRO"                   = "Power BI Pro"
    "PROJECTCLIENT"                  = "Project for Office 365"
    "VISIOCLIENT"                    = "Visio Online Plan 2"
    "MCOEV"                          = "Phone System"
    "MCOMEETADV"                     = "Audio Conferencing"
    "MCOSTANDARD"                    = "Skype for Business Online Plan 2"
    "RIGHTSMANAGEMENT"               = "Azure Information Protection P1"
    "RIGHTSMANAGEMENT_ADHOC"         = "Rights Management Ad Hoc"
    "CRMSTANDARD"                    = "Dynamics CRM Online"
    "DYN365_ENTERPRISE_SALES"        = "Dynamics 365 for Sales Enterprise"
    "DYN365_ENTERPRISE_TEAM_MEMBERS" = "Dynamics 365 Team Members"
    "INTUNE_A"                       = "Microsoft Intune"
    "DEVELOPERPACK"                  = "Office 365 E3 Developer"
    "ENTERPRISEWITHSCAL"             = "Office 365 Enterprise E4"
    "LITEPACK"                       = "Office 365 Small Business"
    "LITEPACK_P2"                    = "Office 365 Small Business Premium"
    "MIDSIZEPACK"                    = "Office 365 Midsize Business"
    "STANDARDWOFFPACK"               = "Office 365 Enterprise E2"
    "OFFICESUBSCRIPTION"             = "Office 365 ProPlus"
    "Microsoft_365_Copilot"          = "Microsoft 365 Copilot"
    "Microsoft_Teams_Rooms_Basic"    = "Microsoft Teams Rooms Basic"
    "Power_Pages_vTrial_for_Makers"  = "Power Pages Trial for Makers"
    "POWERAPPS_VIRAL"                = "Power Apps Viral"
    "WINDOWS_STORE"                  = "Windows Store"
}

# Process licenses data and calculate available licenses
$filteredLicenses = $response.value | ForEach-Object {
    $skuCode = $_.skuPartNumber
    $name = if ($skuNamesMap.ContainsKey($skuCode)) { $skuNamesMap[$skuCode] } else { $skuCode }
    [PSCustomObject]@{
        "License"   = $name
        "Available" = ($_.prepaidUnits.enabled + $_.prepaidUnits.warning + $_.prepaidUnits.suspended - $_.consumedUnits)
    }
}

# Exclude unwanted license SKUs
$filteredLicenses = $filteredLicenses | Where-Object {
    ($_.'License' -notmatch 'Windows Store') -and
    ($_.'License' -notmatch 'Stream') -and
    ($_.'License' -notmatch 'Ad Hoc') -and
    ($_.'License' -notmatch 'vTrial for Makers') -and
    ($_.'License' -notmatch 'Free') -and
    ($_.'License' -notmatch 'Trial')
}

function Get-LicenseHtmlTable {
    param (
        [Parameter(Mandatory = $true)][array]$licenseData,
        [Parameter(Mandatory = $true)][string]$reportDate,
        [Parameter(Mandatory = $false)][bool]$isAttachment = $false
    )

    $tableHeader = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr>
    <td colspan='2' style='text-align:center; padding:10px;'>
        <img src='$Client1Logo' alt='Client1 Logo' style='height:45px; display:inline-block; pointer-events: none;' />
    </td>
</tr>
<tr>
    <td colspan='2' style='text-align:center; font-weight:bold; font-size:18px; padding:10px;'>
        Available Licenses Report<br>
        <span style='font-size:12px; font-weight:normal;'>Report generated: $reportDate</span>
    </td>
</tr>
<tr style='text-align:center; font-weight:bold; background-color:#f0f0f0;'>
    <th style='border: 1px solid #000000;'>License</th>
    <th style='border: 1px solid #000000;'>Available</th>
</tr>
"@

    $tableRows = ""
    foreach ($item in $licenseData | Sort-Object License) {
        $rowStyle = if ($isAttachment -and $item.Available -lt 0) { "background-color: #ff6666;" } else { "" }
        $tableRows += "<tr style='text-align:center; $rowStyle'>" +
        "<td style='border: 1px solid #000000;'>$($item.License)</td>" +
        "<td style='border: 1px solid #000000;'>$($item.Available)</td>" +
        "</tr>"
    }

    $tableFooter = @"
<tr>
    <td colspan='2' style='text-align:center; padding:10px;'>
        <img src='$Client2Logo' alt='Client2 Logo' style='height:45px; display:inline-block; pointer-events: none;' />
    </td>
</tr>
</table>
"@

    return $tableHeader + $tableRows + $tableFooter
}

$reportDateText = $currentDate.ToString("dd/MM/yyyy HH:mm")
$htmlBody = Get-LicenseHtmlTable -licenseData $filteredLicenses -reportDate $reportDateText -isAttachment:$false
$htmlAttachment = Get-LicenseHtmlTable -licenseData $filteredLicenses -reportDate $reportDateText -isAttachment:$true

$weekNumber = [int][math]::Ceiling($currentDate.Day / 7)
$months = @("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
$monthName = $months[$currentDate.Month - 1]
$subject = "EntraID Report - Available Licenses (Week $weekNumber $monthName)"

$toList = @()
foreach ($recipient in $recipients) {
    $toList += @{ emailAddress = @{ address = $recipient } }
}

$message = @{
    subject      = $subject
    body         = @{ contentType = "HTML"; content = $htmlBody }
    toRecipients = $toList
}

if ($useAttachment -eq 1) {
    $tempFilePath = Join-Path -Path $env:TEMP -ChildPath "license_report.html"
    $htmlAttachment | Out-File -FilePath $tempFilePath -Encoding UTF8
    $fileBytes = [System.IO.File]::ReadAllBytes($tempFilePath)
    $base64File = [System.Convert]::ToBase64String($fileBytes)

    $attachment = @{
        "@odata.type" = "#microsoft.graph.fileAttachment"
        name          = "License_Report.html"
        contentType   = "text/html"
        contentBytes  = $base64File
    }

    $message.attachments = @($attachment)
}

$graphBody = @{ message = $message; saveToSentItems = $false } | ConvertTo-Json -Depth 4 -Compress

Invoke-RestMethod -Method POST `
    -Uri "https://graph.microsoft.com/v1.0/users/$smtpUser/sendMail" `
    -Headers @{ Authorization = "Bearer $accessToken" } `
    -Body $graphBody `
    -ContentType "application/json; charset=utf-8"
