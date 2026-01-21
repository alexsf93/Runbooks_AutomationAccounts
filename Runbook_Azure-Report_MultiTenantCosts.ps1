<#
.SYNOPSIS
    Generates a monthly cost report with VAT for two Azure tenants.

.DESCRIPTION
    Authenticates against two Azure tenants, retrieves all subscriptions, and calculates the current month's cost with VAT (21%).
    Generates an HTML report and sends it via email using Microsoft Graph API.

.PARAMETER GraphClientId
    Application (client) ID for Tenant 1.

.PARAMETER GraphTenantId
    Directory (tenant) ID for Tenant 1.

.PARAMETER GraphSecret
    Client secret for Tenant 1.

.PARAMETER GraphClientId2
    Application (client) ID for Tenant 2.

.PARAMETER GraphTenantId2
    Directory (tenant) ID for Tenant 2.

.PARAMETER GraphSecret2
    Client secret for Tenant 2.

.PARAMETER Correo_No-Reply
    SMTP credentials (PSCredential).

.EXAMPLE
    # Run as Azure Automation Runbook
    Start-AutomationRunbook -Name "Runbook_Azure-Report_MultiTenantCosts"

.NOTES
    Author: Automated System
    Version: 1.0
    Prerequisites: Mail.Send, Cost Management Reader
#>

$graphClientId1 = Get-AutomationVariable -Name "GraphClientId"
$graphTenantId1 = Get-AutomationVariable -Name "GraphTenantId"
$graphClientSecret1 = Get-AutomationVariable -Name "GraphSecret"

$graphClientId2 = Get-AutomationVariable -Name "GraphClientId2"
$graphTenantId2 = Get-AutomationVariable -Name "GraphTenantId2"
$graphClientSecret2 = Get-AutomationVariable -Name "GraphSecret2"

$emailCredential = Get-AutomationPSCredential -Name "Correo_No-Reply"
$emailUser = $emailCredential.UserName
$recipients = @("example@domain.com", "example1@domain.com")

$clientLogo = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"

function Get-AccessToken {
    param(
        [string]$tenant,
        [string]$clientId,
        [string]$clientSecret,
        [string]$scope
    )
    $body = @{
        client_id     = $clientId
        scope         = $scope
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }
    $tokenResponse = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token" -Body $body
    return $tokenResponse.access_token
}

function Get-Subscriptions {
    param([string]$accessToken)
    $subsUri = "https://management.azure.com/subscriptions?api-version=2020-01-01"
    $response = Invoke-RestMethod -Method GET -Uri $subsUri -Headers @{ Authorization = "Bearer $accessToken" }
    return $response.value
}

function Get-SubscriptionName {
    param ([string]$accessToken, [string]$subscriptionId)
    $subUri = "https://management.azure.com/subscriptions/${subscriptionId}?api-version=2020-01-01"
    try {
        $responseSub = Invoke-RestMethod -Method GET -Uri $subUri -Headers @{ Authorization = "Bearer $accessToken" }
        return $responseSub.displayName
    }
    catch {
        return "Name not available"
    }
}

function Get-CostWithVAT {
    param([string]$accessToken, [string]$subscriptionId)
    
    $today = (Get-Date).ToUniversalTime()
    $startDate = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0
    $endDate = $today

    $costUri = "https://management.azure.com/subscriptions/$subscriptionId/providers/Microsoft.CostManagement/query?api-version=2023-03-01"
    $costBody = @{
        type       = "ActualCost"
        timeframe  = "Custom"
        timePeriod = @{
            from = $startDate.ToString("yyyy-MM-ddT00:00:00Z")
            to   = $endDate.ToString("yyyy-MM-ddT23:59:59Z")
        }
        dataset    = @{
            aggregation = @{
                totalCost = @{
                    name     = "PreTaxCost"
                    function = "Sum"
                }
            }
            granularity = "None"
        }
    } | ConvertTo-Json -Depth 10 -Compress

    try {
        $responseCost = Invoke-RestMethod -Method POST -Uri $costUri -Headers @{ Authorization = "Bearer $accessToken" } -Body $costBody -ContentType "application/json"
        $costPreTax = $responseCost.properties.rows[0][0]
        $vat = 0.21
        $costWithVAT = [math]::Round($costPreTax * (1 + $vat), 2)
        return $costWithVAT
    }
    catch {
        return 0
    }
}

$accessToken1 = Get-AccessToken -tenant $graphTenantId1 -clientId $graphClientId1 -clientSecret $graphClientSecret1 -scope "https://management.azure.com/.default"
$accessToken2 = Get-AccessToken -tenant $graphTenantId2 -clientId $graphClientId2 -clientSecret $graphClientSecret2 -scope "https://management.azure.com/.default"

$subscriptions1 = Get-Subscriptions -accessToken $accessToken1
$subscriptions2 = Get-Subscriptions -accessToken $accessToken2

$allSubscriptions = @()
if ($subscriptions1) { $allSubscriptions += $subscriptions1 }
if ($subscriptions2) { $allSubscriptions += $subscriptions2 }

$timeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Romance Standard Time")
$nowUTC = [DateTime]::UtcNow
$currentDate = [System.TimeZoneInfo]::ConvertTimeFromUtc($nowUTC, $timeZone)
$currentDateText = $currentDate.ToString("dd/MM/yyyy HH:mm", [System.Globalization.CultureInfo]::GetCultureInfo("en-US"))

$startPeriod = (Get-Date -Day 1).ToString("dd/MM/yyyy", [System.Globalization.CultureInfo]::GetCultureInfo("en-US"))
$endPeriod = (Get-Date).ToString("dd/MM/yyyy", [System.Globalization.CultureInfo]::GetCultureInfo("en-US"))
$reportPeriod = "$startPeriod - $endPeriod"

$totalCost = 0
$rowsHtml = ""
foreach ($sub in $allSubscriptions) {
    $subId = $sub.subscriptionId
    $token = if ($subscriptions1 -and ($subscriptions1 | Where-Object { $_.subscriptionId -eq $subId })) { $accessToken1 } else { $accessToken2 }

    $subName = Get-SubscriptionName -accessToken $token -subscriptionId $subId
    $costVAT = Get-CostWithVAT -accessToken $token -subscriptionId $subId
    $totalCost += $costVAT

    $rowsHtml += @"
<tr style='text-align:center;'>
    <td style='border: 1px solid #000;'>$subName</td>
    <td style='border: 1px solid #000;'>$subId</td>
    <td style='border: 1px solid #000;'>$costVAT €</td>
</tr>
"@
}

$bodyHtml = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<title>Azure Cost Report</title>
</head>
<body>
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000; width: 100%; max-width:600px; margin:auto;'>
<tr>
    <td colspan='3' style='text-align:center; font-weight:bold; font-size:18px; padding:10px;'>
        Azure Monthly Cost Report with VAT<br>
        <span style='font-size:12px; font-weight:normal;'>Period: $reportPeriod<br>Generated: $currentDateText</span>
    </td>
</tr>
<tr style='text-align:center; font-weight:bold; background-color:#f0f0f0;'>
    <th style='border: 1px solid #000;'>Subscription Name</th>
    <th style='border: 1px solid #000;'>Subscription ID</th>
    <th style='border: 1px solid #000;'>Amount</th>
</tr>
$rowsHtml
<tr style='font-weight:bold; background-color:#dcdcdc;'>
    <td colspan='2' style='border: 1px solid #000; text-align:right;'>Total</td>
    <td style='border: 1px solid #000; text-align:center;'>$totalCost €</td>
</tr>
<tr>
    <td colspan='3' style='text-align:center; padding:10px;'>
        <img src='$clientLogo' alt='Client Logo' style='height:45px; display:inline-block; pointer-events: none;' />
    </td>
</tr>
</table>
</body>
</html>
"@

$graphTokenResponse = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$graphTenantId1/oauth2/v2.0/token" -Body @{
    client_id     = $graphClientId1
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $graphClientSecret1
    grant_type    = "client_credentials"
}
$graphToken = $graphTokenResponse.access_token
$headersGraph = @{ Authorization = "Bearer $graphToken" }

$mailBody = @{
    message = @{
        subject      = "Azure Cost Report with VAT ($reportPeriod)"
        body         = @{
            contentType = "HTML"
            content     = $bodyHtml
        }
        toRecipients = @()
    }
}
foreach ($recipient in $recipients) {
    $mailBody.message.toRecipients += @{ emailAddress = @{ address = $recipient } }
}
$mailJson = $mailBody | ConvertTo-Json -Depth 10
$sendMailUri = "https://graph.microsoft.com/v1.0/users/$emailUser/sendMail"

Invoke-RestMethod -Method POST -Uri $sendMailUri -Headers $headersGraph -Body $mailJson -ContentType "application/json; charset=utf-8"
