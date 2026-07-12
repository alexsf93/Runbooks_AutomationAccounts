<#
.SYNOPSIS
    Automates the naming convention enforcement for iOS and iPadOS devices in Microsoft Intune.

.DESCRIPTION
    This runbook is designed to run on a scheduled basis within Azure Automation.
    It utilizes the Automation Account's System-Assigned Managed Identity to authenticate
    against Microsoft Graph and enforces a strict sequential naming pattern (IPHONELVLXXXX / IPADLVLXXXX)
    for corporate devices enrolled via Automated Device Enrollment (ADE).

    PREREQUISITES:
    The Automation Account's System-Assigned Managed Identity MUST be granted the following
    Microsoft Graph Application Permissions:
    - DeviceManagementManagedDevices.Read.All
    - DeviceManagementManagedDevices.PrivilegedOperations.All

.NOTES
    Author:   Alejandro Suárez (@alexsf93)
    Date:     2026-05-03
    Version:  2.4
#>

$Config = @{
    Padding          = 4
    PrefixiPad       = "IPAD-"
    PrefixiPhone     = "IPHONE-"
    EnrollmentFilter = "appleBulkWithUser"
    GraphBaseUri     = "https://graph.microsoft.com/beta/deviceManagement/managedDevices"
}

$RegexiPad   = "^$([regex]::Escape($Config.PrefixiPad))\d{$($Config.Padding)}$"
$RegexiPhone = "^$([regex]::Escape($Config.PrefixiPhone))\d{$($Config.Padding)}$"

Import-Module Az.Accounts -ErrorAction Stop

try {
    $null = Connect-AzAccount -Identity -ErrorAction Stop
    $TokenResult = (Get-AzAccessToken -ResourceTypeName MSGraph -ErrorAction Stop).Token
    
    if ($TokenResult -is [System.Security.SecureString]) {
        $Token = (New-Object System.Management.Automation.PSCredential("Token", $TokenResult)).GetNetworkCredential().Password
    }
    else {
        $Token = $TokenResult
    }
    
    $RequestHeaders = @{
        Authorization  = "Bearer $Token"
        "Content-Type" = "application/json"
    }
}
catch {
    Write-Error "CRITICAL: Authentication failed for Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

$ManagedDevices = [System.Collections.Generic.List[object]]::new()
$FetchUri = "$($Config.GraphBaseUri)?`$filter=operatingSystem eq 'iOS'&`$top=1000"

while ($FetchUri) {
    try {
        $Response = Invoke-RestMethod -Method GET -Uri $FetchUri -Headers $RequestHeaders -ErrorAction Stop
        if ($Response.value) {
            $ManagedDevices.AddRange($Response.value)
        }
        $FetchUri = $Response.'@odata.nextLink'
    }
    catch {
        Write-Error "CRITICAL: Error querying Microsoft Graph API: $($_.Exception.Message)"
        exit 1
    }
}

if ($ManagedDevices.Count -eq 0) {
    return
}

$LastNumberiPad   = 0
$LastNumberiPhone = 0

$ExistingiPads = $ManagedDevices | Where-Object { $_.deviceName -match $RegexiPad }
if ($ExistingiPads) {
    $LastNumberiPad = ($ExistingiPads.deviceName | ForEach-Object {
        [int]($_ -replace [regex]::Escape($Config.PrefixiPad), "")
    } | Measure-Object -Maximum).Maximum
}

$ExistingiPhones = $ManagedDevices | Where-Object { $_.deviceName -match $RegexiPhone }
if ($ExistingiPhones) {
    $LastNumberiPhone = ($ExistingiPhones.deviceName | ForEach-Object {
        [int]($_ -replace [regex]::Escape($Config.PrefixiPhone), "")
    } | Measure-Object -Maximum).Maximum
}

$TargetDevices = $ManagedDevices | Where-Object {
    $_.deviceEnrollmentType -eq $Config.EnrollmentFilter -and (
        (($_.model -match "iPad") -and ($_.deviceName -notmatch $RegexiPad)) -or
        (($_.model -match "iPhone") -and ($_.deviceName -notmatch $RegexiPhone))
    )
}

if (-not $TargetDevices) {
    return
}

foreach ($Device in $TargetDevices) {
    $IsiPad = $Device.model -match "iPad"
    $IsiPhone = $Device.model -match "iPhone"
    
    if (-not ($IsiPad -or $IsiPhone)) {
        continue
    }

    if ($IsiPad) {
        $LastNumberiPad++
        $CurrentSequence = $LastNumberiPad
        $CurrentPrefix   = $Config.PrefixiPad
    }
    else {
        $LastNumberiPhone++
        $CurrentSequence = $LastNumberiPhone
        $CurrentPrefix   = $Config.PrefixiPhone
    }

    $NewDeviceName = "{0}{1}" -f $CurrentPrefix, $CurrentSequence.ToString().PadLeft($Config.Padding, '0')

    try {
        $RenameUri = "$($Config.GraphBaseUri)/$($Device.id)/setDeviceName"
        $BodyPayload = @{ deviceName = $NewDeviceName } | ConvertTo-Json
        $null = Invoke-RestMethod -Method POST -Uri $RenameUri -Headers $RequestHeaders -Body $BodyPayload -ErrorAction Stop

        $SyncUri = "$($Config.GraphBaseUri)/$($Device.id)/syncDevice"
        $null = Invoke-RestMethod -Method POST -Uri $SyncUri -Headers $RequestHeaders -ErrorAction Stop

        Write-Output "SUCCESS: Device renamed to $NewDeviceName (ID: $($Device.id))"
    }
    catch {
        Write-Error "FAILED: Failed to process device with ID $($Device.id). Reason: $($_.Exception.Message)"
        
        if ($IsiPad) { $LastNumberiPad-- } else { $LastNumberiPhone-- }
    }
}
