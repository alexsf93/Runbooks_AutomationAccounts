<#
.SYNOPSIS
Synchronizes an Azure AD group: adds users whose EmployeeId matches, and removes those who no longer meet the condition.

.DESCRIPTION
For the specified EmployeeId, this script adds all users who have it to the defined group.
Afterward, it removes from the group any members who no longer meet the condition, ensuring only valid users remain.

.CONFIGURATION PARAMETERS
- GraphClientId: Application (client) ID registered in Azure AD.
- GraphTenantId: Directory (tenant) ID in Azure AD.
- GraphSecret: Application secret.

.REQUIRED PERMISSIONS (application)
- User.Read.All
- Group.ReadWrite.All

.NOTES
- Uses Write-Output for logging in Azure Automation Runbook.
- Modify EmployeeId and GroupName as needed.

#>

# Retrieve variables from Automation Account
$clientId     = Get-AutomationVariable -Name "GraphClientId"
$tenantId     = Get-AutomationVariable -Name "GraphTenantId"
$clientSecret = Get-AutomationVariable -Name "GraphSecret"

# Editable configuration variables
$EmployeeIdToCheck = "00001"
$GroupName = "SG-TEST"

Write-Output "Script started. EmployeeId: ${EmployeeIdToCheck}, Group: ${GroupName}"

# Get access token
$Scope = "https://graph.microsoft.com/.default"
$Body = @{
    client_id     = $clientId
    client_secret = $clientSecret
    scope         = $Scope
    grant_type    = "client_credentials"
}
try {
    $TokenResponse = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $Body
    $AccessToken = $TokenResponse.access_token
} catch {
    Write-Output "Error obtaining token: $_"
    exit 1
}

function Invoke-Graph {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        $Body = $null
    )
    $Headers = @{Authorization = "Bearer $AccessToken"}
    if ($Body) {
        $BodyJson = $Body | ConvertTo-Json -Depth 5
        return Invoke-RestMethod -Uri $Uri -Headers $Headers -Method $Method -Body $BodyJson -ContentType "application/json"
    } else {
        return Invoke-RestMethod -Uri $Uri -Headers $Headers -Method $Method
    }
}

# Find group by name
try {
    $Group = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$GroupName'"
    if (-not $Group.value) {
        Write-Output "Group not found: ${GroupName}"
        exit 0
    }
    $GroupId = $Group.value[0].id
} catch {
    Write-Output "Error searching for group: $_"
    exit 1
}

# Find users with the required EmployeeId
try {
    $UsersToKeep = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/users?`$filter=employeeId eq '$EmployeeIdToCheck'&`$select=id,userPrincipalName,employeeId"
    $UsersToKeepDict = @{}
    foreach ($u in $UsersToKeep.value) { $UsersToKeepDict[$u.id] = $u }
    $upns = ($UsersToKeep.value | ForEach-Object { $_.userPrincipalName }) -join ", "
    Write-Output "Users with EmployeeId ${EmployeeIdToCheck}: $upns"
} catch {
    Write-Output "Error searching for users: $_"
    exit 1
}

# Get current group members (with paging)
$GroupMembers = @()
$NextLink = "https://graph.microsoft.com/v1.0/groups/$GroupId/members?`$select=id,userPrincipalName"
do {
    try {
        $Response = Invoke-Graph -Uri $NextLink
        $GroupMembers += $Response.value
        $NextLink = $Response.'@odata.nextLink'
    } catch {
        Write-Output "Error getting group members: $_"
        exit 1
    }
} while ($NextLink)

$memberUPNs = ($GroupMembers | Where-Object { $_.userPrincipalName } | ForEach-Object { $_.userPrincipalName }) -join ", "
Write-Output "Current group members: $memberUPNs"
$GroupMembersUserIds = $GroupMembers | Where-Object { $_.userPrincipalName } | ForEach-Object { $_.id }

# ADD: users with required EmployeeId who are not in the group
foreach ($User in $UsersToKeep.value) {
    if (-not ($GroupMembersUserIds -contains $User.id)) {
        try {
            $Body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($User.id)" }
            Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members/`$ref" -Method "POST" -Body $Body
            Write-Output "User $($User.userPrincipalName) added to group ${GroupName}"
        } catch {
            Write-Output "Error adding $($User.userPrincipalName): $_"
        }
    }
}

# REMOVE: users who are in the group but do NOT have the required EmployeeId
foreach ($Member in $GroupMembers) {
    if ($Member.userPrincipalName -and -not $UsersToKeepDict.ContainsKey($Member.id)) {
        try {
            Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members/$($Member.id)/`$ref" -Method "DELETE"
            Write-Output "User $($Member.userPrincipalName) removed from group ${GroupName} (incorrect or empty EmployeeId)"
        } catch {
            Write-Output "Error removing $($Member.userPrincipalName): $_"
        }
    }
}

Write-Output "Sync completed."
