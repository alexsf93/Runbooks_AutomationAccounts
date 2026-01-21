<#
.SYNOPSIS
    Generates a report of SharePoint files larger than 500MB.

.DESCRIPTION
    Scans internal SharePoint sites for files exceeding 500MB.
    Generates an HTML report with file details (Name, Size, Owner, Link) and sends it via email.

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
    Start-AutomationRunbook -Name "Runbook_Sharepoint-Report_LargerThan500MB"

.NOTES
    Author: Automated System
    Version: 1.0
    Prerequisites: Sites.Read.All, Mail.Send
#>

# Configurable variables
$clientId = Get-AutomationVariable -Name "GraphClientId"
$tenantId = Get-AutomationVariable -Name "GraphTenantId"
$clientSecret = Get-AutomationVariable -Name "GraphSecret"
$smtpCredential = Get-AutomationPSCredential -Name "Correo_No-Reply"
$smtpUser = $smtpCredential.UserName

# Multiple email recipients
$recipients = @("example@domain.com", "example1@domain.com")

# Attachment flag: 1 = send HTML attachment + body, 0 = body only
$useAttachment = 0

# Logo URLs (updated)
$client1LogoUrl = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"
$client2LogoUrl = "https://staintunenaxvan.blob.core.windows.net/wallpapers/LOGO_NAXVAN_Mesa_de_trabajo_1_copia_2.png"

# Get Microsoft Graph token
$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$accessToken = $tokenResponse.access_token
$headers = @{ Authorization = "Bearer $accessToken" }
$minSizeBytes = 500MB
$largeFiles = New-Object System.Collections.Generic.List[PSObject]

function Get-LargeFilesRecursive {
    param (
        [string]$siteName,
        [string]$driveName,
        [string]$driveId,
        [string]$itemId = "root"
    )
    $url = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/children?`$top=999"
    do {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        foreach ($item in $response.value) {
            if ($item.folder) {
                Get-LargeFilesRecursive -siteName $siteName -driveName $driveName -driveId $driveId -itemId $item.id
            }
            elseif ($item.file -and $item.size -gt $minSizeBytes) {
                $owner = $item.createdBy.user.displayName
                if (-not $owner) { $owner = "(Unknown)" }
                $largeFile = [PSCustomObject]@{
                    Name   = $item.name
                    SizeMB = [math]::Round($item.size / 1MB, 2)
                    Site   = $siteName
                    Owner  = $owner
                    Link   = $item.webUrl
                }
                $largeFiles.Add($largeFile)
            }
        }
        $url = $response.'@odata.nextLink'
    } while ($url)
}

# Retrieve SharePoint sites
$sites = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites?search=*" -Headers $headers
foreach ($site in $sites.value) {
    try {
        $siteId = $site.id
        $siteName = $site.name
        $drives = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives" -Headers $headers
        foreach ($drive in $drives.value) {
            Get-LargeFilesRecursive -siteName $siteName -driveName $drive.name -driveId $drive.id
        }
    }
    catch {
        Write-Host "Error accessing site $($site.name). Skipping."
    }
}

# Current date/time in Spain timezone
$dateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), "Romance Standard Time").ToString("dd/MM/yyyy HH:mm")

# Build HTML table
if ($largeFiles.Count -eq 0) {
    $htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr>
    <td colspan='5' style='border: 2px solid #000000; text-align:center; vertical-align: middle; padding:10px;'>
        <img src='$client2LogoUrl' alt='Client Logo' style='height:50px; display:inline-block;'/>
    </td>
</tr>
<tr>
    <td colspan='5' style='border: 2px solid #000000; text-align:center; font-weight:bold; font-size:18px; vertical-align: middle; padding:10px;'>
        SharePoint Files Larger Than 500MB Report<br>
        <span style='font-size:12px; font-weight:normal;'>Report generated: $dateTime</span>
    </td>
</tr>
<tr style='text-align:center; font-weight:bold; background-color:#f0f0f0;'>
    <th>Site</th><th>Name</th><th>Size (MB)</th><th>Owner</th><th>Link</th>
</tr>
<tr><td colspan='5' style='text-align:center;'>No files larger than 500MB were found.</td></tr>
<tr>
    <td colspan='5' style='border: 2px solid #000000; text-align:center; vertical-align: middle; padding:10px;'>
        <img src='$client1LogoUrl' alt='Client Logo' style='height:50px; display:inline-block;' />
    </td>
</tr>
</table>
"@
}
else {
    $htmlTable = @"
<table border='2' cellpadding='6' cellspacing='0' style='border-collapse: collapse; font-family: Arial; border-color:#000000; width: 100%;'>
<tr>
    <td colspan='5' style='border: 2px solid #000000; text-align:center; vertical-align: middle; padding:10px;'>
        <img src='$client2LogoUrl' alt='Client Logo' style='height:50px; display:inline-block;'/>
    </td>
</tr>
<tr>
    <td colspan='5' style='border: 2px solid #000000; text-align:center; font-weight:bold; font-size:18px; vertical-align: middle; padding:10px;'>
        SharePoint Files Larger Than 500MB Report<br>
        <span style='font-size:12px; font-weight:normal;'>Report generated: $dateTime</span>
    </td>
</tr>
<tr style='text-align:center; font-weight:bold; background-color:#f0f0f0;'>
    <th>Site</th><th>Name</th><th>Size (MB)</th><th>Owner</th><th>Link</th>
</tr>
"@
    foreach ($file in $largeFiles) {
        $htmlTable += "<tr style='text-align:center;'>" +
        "<td style='border: 1px solid #000000;'>$($file.Site)</td>" +
        "<td style='border: 1px solid #000000;'>$($file.Name)</td>" +
        "<td style='border: 1px solid #000000;'>$($file.SizeMB)</td>" +
        "<td style='border: 1px solid #000000;'>$($file.Owner)</td>" +
        "<td style='border: 1px solid #000000;'><a href='$($file.Link)' target='_blank'>View file</a></td>" +
        "</tr>"
    }
    $htmlTable += @"
<tr>
    <td colspan='5' style='border: 2px solid #000000; text-align:center; vertical-align: middle; padding:10px;'>
        <img src='$client1LogoUrl' alt='Client Logo' style='height:50px; display:inline-block;' />
    </td>
</tr>
</table>
"@
}

# Prepare recipient list for Graph API payload
$toRecipientsArray = $recipients | ForEach-Object {
    @{ emailAddress = @{ address = $_ } }
}

# Email subject
$subject = "SharePoint Report - Files > 500MB"

# Construct email payload with or without attachment
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
                    name          = "SharePoint_Report_LargeFiles.html"
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

# Send email via Microsoft Graph
Invoke-RestMethod -Method POST `
    -Uri "https://graph.microsoft.com/v1.0/users/$smtpUser/sendMail" `
    -Headers @{ Authorization = "Bearer $accessToken" } `
    -Body $emailPayload `
    -ContentType "application/json; charset=utf-8"
