<#
.SYNOPSIS
    Assigns Fabric Free license to users with Power BI Pro.

.DESCRIPTION
    Queries users with Power BI Pro license and assigns Fabric Free license if missing.
    Uses Microsoft Graph API and handles pagination.

.PARAMETER GraphClientId
    Application (client) ID registered in Azure AD.

.PARAMETER GraphTenantId
    Directory (tenant) ID in Azure AD.

.PARAMETER GraphSecret
    Application secret.

.EXAMPLE
    # Run as Azure Automation Runbook
    Start-AutomationRunbook -Name "Runbook_Azure-Licenses_GroupBasedLicenseTenantFree"

.NOTES
    Author: Automated System
    Version: 1.0
    Prerequisites: User.Read.All, Directory.ReadWrite.All
#>

# Retrieve variables from Automation Account
$clientId = Get-AutomationVariable -Name "GraphClientId"
$tenantId = Get-AutomationVariable -Name "GraphTenantId"
$clientSecret = Get-AutomationVariable -Name "GraphSecret"

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

# Helper functions for Graph calls
function Invoke-GraphGet {
    param (
        [string]$uri
    )
    return Invoke-RestMethod -Headers @{ Authorization = "Bearer $accessToken" } -Uri $uri -Method Get
}

function Invoke-GraphPost {
    param (
        [string]$uri,
        [object]$body
    )
    return Invoke-RestMethod -Headers @{ Authorization = "Bearer $accessToken"; "Content-Type" = "application/json" } -Uri $uri -Method Post -Body ($body | ConvertTo-Json -Depth 5)
}

# 1. Retrieve available SKUs
$skus = Invoke-GraphGet -uri "https://graph.microsoft.com/v1.0/subscribedSkus"
$powerBiProSku = $skus.value | Where-Object { $_.skuPartNumber -eq "POWER_BI_PRO" }
$fabricFreeSku = $skus.value | Where-Object { $_.skuPartNumber -like "*FABRIC_FREE*" }

if (-not $powerBiProSku) {
    Write-Output "POWER_BI_PRO SKU not found. Check the exact name in your tenant."
    Write-Output "Available SKUs:"
    $skus.value | Select-Object skuPartNumber, skuId
    return
}
if (-not $fabricFreeSku) {
    Write-Output "Fabric Free SKU not found. Check the exact name in your tenant."
    Write-Output "Available SKUs:"
    $skus.value | Select-Object skuPartNumber, skuId
    return
}

# 2. Retrieve ALL users with Power BI Pro (pagination included)
$filter = "assignedLicenses/any(x:x/skuId eq $($powerBiProSku.skuId))"
$baseUri = "https://graph.microsoft.com/v1.0/users?`$filter=$filter&`$select=id,userPrincipalName,assignedLicenses&`$top=100"
$nextLink = $baseUri

$atLeastOneAssigned = $false

do {
    $usersResponse = Invoke-GraphGet -uri $nextLink

    if ($null -eq $usersResponse.value -or $usersResponse.value.Count -eq 0) {
        break
    }

    foreach ($user in $usersResponse.value) {
        $alreadyHasFabricFree = $user.assignedLicenses | Where-Object { $_.skuId -eq $fabricFreeSku.skuId }
        if (-not $alreadyHasFabricFree) {
            $assignBody = @{
                "addLicenses"    = @(@{ "skuId" = $fabricFreeSku.skuId })
                "removeLicenses" = @()
            }
            $assignUri = "https://graph.microsoft.com/v1.0/users/$($user.id)/assignLicense"
            Invoke-GraphPost -uri $assignUri -body $assignBody | Out-Null
            Write-Output "Fabric Free license assigned to user: $($user.userPrincipalName)"
            $atLeastOneAssigned = $true
        }
    }

    $nextLink = $null
    if ($usersResponse.'@odata.nextLink') {
        $nextLink = $usersResponse.'@odata.nextLink'
    }
} while ($nextLink)

if (-not $atLeastOneAssigned) {
    Write-Output "No assignments made. All users already had the Fabric Free license."
}
