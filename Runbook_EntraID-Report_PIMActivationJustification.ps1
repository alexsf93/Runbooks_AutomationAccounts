<#
.SYNOPSIS
Generates a report of Privileged Identity Management (PIM) activations in the last 30 days using Microsoft Graph API.

.DESCRIPTION
This script retrieves audit logs related to PIM role activations, generates an HTML report with validation highlighting,
and sends the report via email using Microsoft Graph. Rows with justification shorter than the minimum defined are highlighted in red.

.PARAMETER Configurable variables
- useAttachment: 1 to send the HTML report as an attachment and in the body, 0 to send only in the body.
- minJustificationLength: Minimum characters required in the justification to avoid red highlight.
- recipients: Array of email recipient addresses.
- client1LogoUrl: URL of the first client logo.
- client2LogoUrl: URL of the second client logo.

.API Permissions Needed as Application
- AuditLog.Read.All
- Mail.Send
#>

# --- Configurable variables ---
$useAttachment = 0
$minJustificationLength = 10
$recipients = @("example@domain.com", "example1@domain.com")

$client1LogoUrl = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"
$client2LogoUrl = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"

# --- Secure Automation environment variables ---
$clientId     = Get-AutomationVariable -Name "GraphClientId"
$tenantId     = Get-AutomationVariable -Name "GraphTenantId"
$clientSecret = Get-AutomationVariable -Name "GraphSecret"
$smtpCredential = Get-AutomationPSCredential -Name "Correo_No-Reply"
$smtpUser = $smtpCredential.UserName

# --- Authenticate with Microsoft Graph ---
$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$accessToken = $tokenResponse.access_token
$headers = @{ Authorization = "Bearer $accessToken" }

# --- Report date ranges ---
$startDate = (Get-Date).AddDays(-30).ToString("dd/MM/yyyy")
$endDate = (Get-Date).ToString("dd/MM/yyyy")
$reportDate = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), "Romance Standard Time").ToString("dd/MM/yyyy HH:mm")

# --- Retrieve PIM logs ---
$startDateISO = (Get-Date).AddDays(-30).ToString("yyyy-MM-ddTHH:mm:ssZ")
$uri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=activityDisplayName eq 'Add member to role completed (PIM activation)' and activityDateTime ge $startDateISO"

$logs = @()

try {
    do {
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET
        $logs += $response.value
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    if (!$logs) {
        $htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr>
    <td colspan='4' style='border: 2px solid #000000; text-align:center; vertical-align: middle; padding:10px;'>
        <img src='$client1LogoUrl' alt='Client1 Logo' style='height:50px;' />
    </td>
</tr>
<tr>
    <td colspan='4' style='border: 2px solid #000000; text-align:center; font-weight:bold; font-size:18px; vertical-align: middle; padding:10px;'>
        PIM Activations Report<br>
        <span style='font-size:12px; font-weight:normal;'>Report generated: $reportDate</span>
    </td>
</tr>
<tr>
    <td colspan='4' style='border: 2px solid #000000; text-align:center; padding:10px;'>No PIM activations found in the last 30 days.</td>
</tr>
<tr>
    <td colspan='4' style='border: 2px solid #000000; text-align:center; vertical-align: middle; padding:10px;'>
        <img src='$client2LogoUrl' alt='Client2 Logo' style='height:50px;' />
    </td>
</tr>
</table>
"@
    } else {
        $htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr>
    <td colspan='4' style='border: 2px solid #000000; text-align:center; vertical-align: middle; padding:10px;'>
        <img src='$client1LogoUrl' alt='Client1 Logo' style='height:50px;' />
    </td>
</tr>
<tr>
    <td colspan='4' style='border: 2px solid #000000; text-align:center; font-weight:bold; font-size:18px; vertical-align: middle; padding:10px;'>
        PIM Activations Report<br>
        <span style='font-size:12px; font-weight:normal;'>Report generated: $reportDate</span>
    </td>
</tr>
<tr style='text-align:center; font-weight:bold; background-color:#f0f0f0;'>
    <th style='border: 1px solid #000000;'>User</th>
    <th style='border: 1px solid #000000;'>Role</th>
    <th style='border: 1px solid #000000;'>Date</th>
    <th style='border: 1px solid #000000;'>Justification</th>
</tr>
"@

        foreach ($log in $logs) {
            $user = $log.initiatedBy.user.displayName
            $roleObject = $log.targetResources | Where-Object { $_.type -eq "Role" }
            $role = $roleObject.displayName
            $date = (Get-Date $log.activityDateTime).ToLocalTime().ToString("g")

            $justification = $null
            if ($log.additionalDetails) {
                $justification = ($log.additionalDetails | Where-Object { $_.key -match "(?i)justification|reason" }) | Select-Object -ExpandProperty value -First 1
            }
            if (-not $justification) {
                $justification = "(No justification)"
            }

            $rowStyle = ($justification.Length -lt $minJustificationLength) ? "background-color:#ffcccc;" : ""

            $htmlTable += "<tr style='text-align:center; $rowStyle'>" +
                "<td style='border: 1px solid #000000;'>$user</td>" +
                "<td style='border: 1px solid #000000;'>$role</td>" +
                "<td style='border: 1px solid #000000;'>$date</td>" +
                "<td style='border: 1px solid #000000;'>$justification</td>" +
                "</tr>"
        }

        $htmlTable += @"
<tr>
    <td colspan='4' style='border: 2px solid #000000; text-align:center; vertical-align: middle; padding:10px;'>
        <img src='$client2LogoUrl' alt='Client2 Logo' style='height:50px;' />
    </td>
</tr>
</table>
"@
    }

} catch {
    $htmlTable = "<p>Error retrieving data: $_</p>"
}

# --- If useAttachment is enabled, create HTML file and encode it ---
$attachments = @()

if ($useAttachment -eq 1) {
    $tempHtmlPath = "$env:TEMP\pim_report.html"
    $htmlTable | Out-File -FilePath $tempHtmlPath -Encoding UTF8
    $fileBytes = [System.IO.File]::ReadAllBytes($tempHtmlPath)
    $fileContent = [System.Convert]::ToBase64String($fileBytes)

    $attachments += @(
        @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            name          = "pim_report.html"
            contentType   = "text/html"
            contentBytes  = $fileContent
        }
    )
}

# --- Construct the email ---
$subject = "PIM Activation Report - Justifications last 30 days from $startDate to $endDate"

$toRecipientsArray = @()
foreach ($email in $recipients) {
    $toRecipientsArray += @{ emailAddress = @{ address = $email } }
}

$graphBody = @{
    message = @{
        subject = $subject
        body = @{
            contentType = "HTML"
            content = $htmlTable
        }
        toRecipients = $toRecipientsArray
        attachments = $attachments
    }
    saveToSentItems = $false
} | ConvertTo-Json -Depth 4 -Compress

# --- Send the email ---
Invoke-RestMethod -Method POST `
    -Uri "https://graph.microsoft.com/v1.0/users/$smtpUser/sendMail" `
    -Headers @{ Authorization = "Bearer $accessToken" } `
    -Body $graphBody `
    -ContentType "application/json; charset=utf-8"
