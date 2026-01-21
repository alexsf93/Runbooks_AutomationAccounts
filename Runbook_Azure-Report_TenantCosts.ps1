<#
.SYNOPSIS
Azure Automation script to generate a monthly cost report with VAT for all subscriptions in a single Azure tenant, sent via email using Microsoft Graph API.

.DESCRIPTION
This script runs inside Azure Automation and performs the following:
- Authenticates against an Azure tenant using a service principal.
- Retrieves all Azure subscriptions within the tenant.
- Queries the current month's cost (pre-tax) for each subscription using the Azure Cost Management API.
- Calculates the total cost including VAT (21%).
- Builds an HTML report summarizing subscription names, IDs, and total cost with VAT.
- Sends the report by email through Microsoft Graph API using a service account.
- Includes a configurable client logo and provider logo in the HTML email.

.PARAMETER Configurable Variables (Azure Automation Variables):
- GraphClientId       : Application (client) ID for the Azure AD app registration
- GraphTenantId       : Directory (tenant) ID for the Azure tenant
- GraphSecret         : Client secret used to authenticate the app

.PARAMETER Credential (Azure Automation Credential Asset):
- Correo_No-Reply     : Credential used to authenticate the email sender (must have Mail.Send permissions)

.PARAMETER Other Variables:
- recipients          : Array of email addresses to receive the report
- ClientLogo          : URL of the main client logo shown in the report
- Client2Logo         : URL of the secondary/provider logo displayed at the bottom

.API Permissions Required (Azure AD App):
- Microsoft Graph API:
  - Mail.Send (Application permission)
- Azure Service Management API:
  - user_impersonation

.ROLES Required:
- Reader              : Assigned to the app registration for each subscription
- Cost Management Reader : To allow access to cost data

.NOTES
- VAT is hardcoded at 21% but can be changed via the `$iva` variable.
- The report uses the "Romance Standard Time" zone for timestamping (Spain).
- Email is sent through Microsoft Graph using a service account specified in Azure Automation Credentials.

.EXAMPLE
This script can be scheduled in Azure Automation to run monthly and send a detailed HTML cost report with VAT for all subscriptions in a tenant.
#>

# Variables and configuration - Defined in Azure Automation
$clientId = Get-AutomationVariable -Name "GraphClientId"
$tenantId = Get-AutomationVariable -Name "GraphTenantId"
$clientSecret = Get-AutomationVariable -Name "GraphSecret"

$smtpCredential = Get-AutomationPSCredential -Name "Correo_No-Reply"
$smtpUser = $smtpCredential.UserName
$recipients = @("example@domain.com", "example2@domain.com")

$ClientLogo = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"
$Client2Logo = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"

# Get Azure Management API token
$tokenResponseAzure = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body @{
    client_id     = $clientId
    scope         = "https://management.azure.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$accessTokenAzure = $tokenResponseAzure.access_token

# Get list of subscriptions
$subsUri = "https://management.azure.com/subscriptions?api-version=2020-01-01"
$subsResponse = Invoke-RestMethod -Method GET -Uri $subsUri -Headers @{ Authorization = "Bearer $accessTokenAzure" }

if ($subsResponse.value.Count -eq 0) {
    exit 0
}

$iva = 0.21
$filasTabla = ""

foreach ($sub in $subsResponse.value) {
    $subscriptionId = $sub.subscriptionId
    $subscriptionName = $sub.displayName

    $costUri = "https://management.azure.com/subscriptions/$subscriptionId/providers/Microsoft.CostManagement/query?api-version=2023-03-01"
    $costBody = @{
        type = "ActualCost"
        timeframe = "MonthToDate"
        dataset = @{
            aggregation = @{
                totalCost = @{
                    name = "PreTaxCost"
                    function = "Sum"
                }
            }
            granularity = "None"
        }
    } | ConvertTo-Json -Depth 10 -Compress

    try {
        $responseCost = Invoke-RestMethod -Method POST `
            -Uri $costUri `
            -Headers @{ Authorization = "Bearer $accessTokenAzure" } `
            -Body $costBody `
            -ContentType "application/json"

        $costPreTax = $responseCost.properties.rows[0][0]
        $costPreTax = if ($null -eq $costPreTax) { 0 } else { $costPreTax }

        $costWithIVA = [math]::Round($costPreTax * (1 + $iva), 2)
    }
    catch {
        $costWithIVA = "Error"
    }

    $filasTabla += @"
<tr style='text-align:center;'>
    <td style='border: 1px solid #000000;'>$subscriptionName</td>
    <td style='border: 1px solid #000000;'>$subscriptionId</td>
    <td style='border: 1px solid #000000;'>$costWithIVA</td>
</tr>
"@
}

# Current date in Spain timezone
$spainTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Romance Standard Time")
$fechaActualUTC = [DateTime]::UtcNow
$fechaActual = [System.TimeZoneInfo]::ConvertTimeFromUtc($fechaActualUTC, $spainTimeZone)
$fechaTexto = $fechaActual.ToString("dd/MM/yyyy HH:mm")

$cuerpoHtml = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<title>Azure Cost Report</title>
</head>
<body>
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial, sans-serif; border-color:#000000; width: 100%; max-width:600px; margin:auto;'>
<tr>
    <td colspan='3' style='text-align:center; padding:10px;'>
        <img src='$ClientLogo' alt='Client Logo' style='height:45px; display:inline-block; pointer-events:none;' />
    </td>
</tr>
<tr>
    <td colspan='3' style='text-align:center; font-weight:bold; font-size:18px; padding:10px;'>
        Monthly Azure Cost Report with VAT<br>
        <span style='font-size:12px; font-weight:normal;'>Report generated: $fechaTexto</span>
    </td>
</tr>
<tr style='text-align:center; font-weight:bold; background-color:#f0f0f0;'>
    <th style='border: 1px solid #000000;'>Subscription Name</th>
    <th style='border: 1px solid #000000;'>Subscription ID</th>
    <th style='border: 1px solid #000000;'>Amount (€)</th>
</tr>
$filasTabla
<tr>
    <td colspan='3' style='text-align:center; padding:10px;'>
        <img src='$Client2Logo' alt='Provider Logo' style='height:45px; display:inline-block; pointer-events:none;' />
    </td>
</tr>
</table>
</body>
</html>
"@

# Get MS Graph token to send email
$tokenResponseGraph = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$accessTokenGraph = $tokenResponseGraph.access_token

$headersGraph = @{ Authorization = "Bearer $accessTokenGraph" }

$mailBody = @{
    message = @{
        subject = "Monthly Azure Cost Report with VAT - $fechaTexto"
        body = @{
            contentType = "HTML"
            content = $cuerpoHtml
        }
        toRecipients = @()
    }
}

foreach ($r in $recipients) {
    $mailBody.message.toRecipients += @{ emailAddress = @{ address = $r } }
}

$mailBodyJson = $mailBody | ConvertTo-Json -Depth 10

$sendMailUri = "https://graph.microsoft.com/v1.0/users/$smtpUser/sendMail"

Invoke-RestMethod -Method POST -Uri $sendMailUri -Headers $headersGraph -Body $mailBodyJson -ContentType "application/json; charset=utf-8"
