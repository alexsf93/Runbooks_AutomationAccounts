<#
.SYNOPSIS
    Sincroniza dinamicamente dos grupos en Microsoft Entra ID segun si los usuarios tienen o no dispositivos gestionados en Intune.

.DESCRIPTION
    Este runbook esta diseñado para ejecutarse en Azure Automation. Se conecta a Microsoft Graph
    utilizando la Managed Identity (Identidad Administrada) del Automation Account y distribuye a los 
    usuarios del tenant en dos grupos:
      - "NaxvanCorp - Gente con dispositivo corporativo" (usuarios con al menos 1 dispositivo en Intune)
      - "NaxvanCorp - Gente sin dispositivo corporativo" (usuarios sin ningun dispositivo en Intune)

    El script calcula los cambios necesarios (diferenciales) para optimizar el rendimiento y evitar llamadas API redundantes.
    Por defecto se ejecuta en modo simulacion (Dry Run). Para aplicar los cambios reales, configure el parametro Commit a $true.

.PARAMETER GroupNameWithDevices
    Nombre del grupo para usuarios con dispositivo gestionado. Valor por defecto: "NaxvanCorp - Gente con dispositivo corporativo"

.PARAMETER GroupNameWithoutDevices
    Nombre del grupo para usuarios sin dispositivo gestionado. Valor por defecto: "NaxvanCorp - Gente sin dispositivo corporativo"

.PARAMETER Commit
    Booleano para aplicar los cambios en el tenant. Por defecto es $false (modo simulacion / Dry Run).

.PARAMETER ExcludeGuests
    Excluir a los usuarios invitados (Guest) de la sincronizacion. Valor por defecto: True.

.PARAMETER OnlyActiveUsers
    Incluir solo a los usuarios con la cuenta habilitada. Valor por defecto: True.

.PARAMETER AllowedDomain
    Dominio permitido para los usuarios. Valor por defecto: "naxvan.es"

.PARAMETER ExcludedUserPatterns
    Patrones de nombres de usuario a excluir de la evaluacion. Valor por defecto: "admin", "test", "prueba", "poc".

.PARAMETER EnforceNamingPattern
    Validar que el UPN cumpla con el patron de nombre de usuario de la organizacion. Valor por defecto: True.

.PARAMETER RequireMail
    Requerir que el usuario tenga un correo electronico configurado. Valor por defecto: True.

.PARAMETER ForceWithDeviceEmails
    Lista de correos de usuarios (ej. CEOs) que se forzara a considerar con dispositivo.

.PARAMETER ExceptedNamingPatternEmails
    Lista de correos de usuarios exceptuados de cumplir con la validacion del patron de nombre.

.REQUIREMENTS
    - Modulo Az.Accounts instalado en el Automation Account.
    - Managed Identity con los siguientes permisos de Microsoft Graph (Application):
        DeviceManagementManagedDevices.Read.All, User.Read.All, Group.ReadWrite.All, GroupMember.ReadWrite.All

    Script PowerShell para asignar estos permisos a la Managed Identity (ejecutar como Administrador de Aplicaciones / Global):
    --------------------------------------------------------------------------------------------------
    # Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"
    #
    # $MSIName    = "Nombre-De-Tu-Automation-Account"
    # $Roles      = @("DeviceManagementManagedDevices.Read.All", "User.Read.All", "Group.ReadWrite.All", "GroupMember.ReadWrite.All")
    # $GraphAppId = "00000003-0000-0000-c000-000000000000"
    #
    # $MSI     = Get-MgServicePrincipal -Filter "displayName eq '$MSIName'"
    # $GraphSP = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'"
    #
    # foreach ($Role in $Roles) {
    #     $AppRole = $GraphSP.AppRoles | Where-Object { $_.Value -eq $Role -and $_.AllowedMemberTypes -contains "Application" }
    #     if ($AppRole) {
    #         New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MSI.Id -PrincipalId $MSI.Id -ResourceId $GraphSP.Id -AppRoleId $AppRole.Id
    #         Write-Output "Asignado: $Role"
    #     }
    # }
    --------------------------------------------------------------------------------------------------

.NOTES
    Name: Runbook_EntraID-SyncDeviceGroups.ps1
    Author: Alejandro Suarez (@alexsf93)
    Version: 3.0.0
    Date: 2026-07-02
#>

param(
    [string]$GroupNameWithDevices = "NaxvanCorp - Gente con dispositivo corporativo",
    [string]$GroupNameWithoutDevices = "NaxvanCorp - Gente sin dispositivo corporativo",
    [bool]$Commit = $false,
    [bool]$ExcludeGuests = $true,
    [bool]$OnlyActiveUsers = $true,
    [string]$AllowedDomain = "naxvan.es",
    [string[]]$ExcludedUserPatterns = @("admin", "test", "prueba", "poc"),
    [bool]$EnforceNamingPattern = $true,
    [bool]$RequireMail = $true,
    [string[]]$ForceWithDeviceEmails = @("alejandro@naxvan.es", "alexsf93@naxvan.es"),
    [string[]]$ExceptedNamingPatternEmails = @(
        "Acerrato@naxvan.es",
        "alejandro@naxvan.es"
    )
)

$DryRun = -not $Commit

Import-Module Az.Accounts -ErrorAction Stop

Write-Output "========================================================================="
Write-Output " Sincronizacion de Grupos de Usuarios segun Dispositivos Intune"
Write-Output "========================================================================="
if ($DryRun) {
    Write-Output " MODO: SIMULACION (Dry Run) - No se realizaran cambios en el tenant."
    Write-Output " Para aplicar los cambios reales, configure el parametro Commit a `$true."
}
else {
    Write-Output " MODO: PRODUCCION (Commit) - Se aplicaran cambios reales en el tenant."
}
Write-Output "=========================================================================`n"

# 1. Conectar a Azure y obtener Token de Microsoft Graph
try {
    Write-Output "Conectando a Azure con Managed Identity..."
    $null = Connect-AzAccount -Identity -ErrorAction Stop
    
    Write-Output "Obteniendo Access Token para Microsoft Graph..."
    $TokenResult = (Get-AzAccessToken -ResourceTypeName MSGraph -ErrorAction Stop).Token
    
    if ($TokenResult -is [System.Security.SecureString]) {
        $Token = ConvertFrom-SecureString -SecureString $TokenResult -AsPlainText
    }
    else {
        $Token = $TokenResult
    }
    
    $RequestHeaders = @{
        Authorization  = "Bearer $Token"
        "Content-Type" = "application/json; charset=utf-8"
    }
    Write-Output "Autenticacion completada con exito.`n"
}
catch {
    Write-Error "Error de autenticacion con Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

# Wrapper de peticiones REST directas a Microsoft Graph
function Invoke-GraphRest {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        [string]$Body = $null
    )
    $FullUri = $Uri
    if ($Uri -notlike "http*") {
        $FullUri = "https://graph.microsoft.com/$Uri"
    }
    
    $Params = @{
        Method      = $Method
        Uri         = $FullUri
        Headers     = $RequestHeaders
        ErrorAction = "Stop"
    }
    if ($Body) {
        $Params.Body = $Body
    }
    
    try {
        return Invoke-RestMethod @Params
    }
    catch {
        $msg = $_.Exception.Message
        $details = ""
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $details = $_.ErrorDetails.Message
        }
        elseif ($_.Exception -and $_.Exception.Response) {
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                if ($stream) {
                    $reader = [System.IO.StreamReader]::new($stream)
                    $details = $reader.ReadToEnd()
                    $reader.Dispose()
                }
            }
            catch {}
        }
        Write-Error "Error en REST Request ($Method $FullUri): $msg. Detalles: $details"
        throw $_
    }
}

# Paginacion Graph API
function Invoke-GraphPaginatedRequest {
    param(
        [string]$Uri
    )
    $results = @()
    $nextUri = $Uri
    while ($nextUri) {
        try {
            $response = Invoke-GraphRest -Method GET -Uri $nextUri
            if ($response.value) {
                $results += $response.value
            }
            $nextUri = $null
            if ($response) {
                $nextUri = $response.'@odata.nextLink'
            }
        }
        catch {
            Write-Error "Error en la consulta Graph API ($nextUri): $($_.Exception.Message)"
            throw $_
        }
    }
    return $results
}

# Validar patron UPN (inicial + apellido)
function Test-UserNamingPattern {
    param(
        [string]$UserPrincipalName,
        [string]$DisplayName
    )
    if ([string]::IsNullOrEmpty($UserPrincipalName) -or [string]::IsNullOrEmpty($DisplayName)) {
        return $false
    }
    
    $username = ($UserPrincipalName.Split('@')[0].ToLower() -replace '[^a-z0-9]', '')
    
    $normalized = $DisplayName.Normalize([System.Text.NormalizationForm]::FormD)
    $cleanName = ($normalized -replace '\p{Mn}', '') -replace '[^a-zA-Z0-9 ]', ''
    
    $parts = $cleanName.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
    if ($parts.Count -lt 2) {
        return $false
    }
    
    $firstName = $parts[0].ToLower()
    $lastNameParts = $parts[1..($parts.Count - 1)] | ForEach-Object { $_.ToLower() }
    
    $lastNameCandidates = @()
    foreach ($part in $lastNameParts) {
        $lastNameCandidates += $part
    }
    if ($lastNameParts.Count -gt 1) {
        $lastNameCandidates += ($lastNameParts -join '')
    }
    
    foreach ($lastName in $lastNameCandidates) {
        $cleanLastName = $lastName -replace '[^a-z0-9]', ''
        
        if ($username.EndsWith($cleanLastName)) {
            $prefix = $username.Substring(0, $username.Length - $cleanLastName.Length)
            if ($prefix.Length -ge 1 -and $prefix.Length -le 3) {
                $firstLetter = $firstName.Substring(0, 1)
                if ($prefix.StartsWith($firstLetter)) {
                    return $true
                }
            }
        }
    }
    
    return $false
}

# Buscar o crear grupo
function Get-OrCreateGroup {
    param(
        [string]$GroupName,
        [string]$Description
    )
    
    Write-Host "Buscando el grupo '$GroupName'..."
    $group = $null
    try {
        $uri = "v1.0/groups?`$filter=displayName eq '$($GroupName -replace "'", "''")'"
        $res = Invoke-GraphRest -Method GET -Uri $uri
        if ($res.value -and $res.value.Count -gt 0) {
            $group = $res.value[0]
        }
    }
    catch {
        Write-Warning "Error al buscar el grupo '$GroupName': $($_.Exception.Message)"
    }

    if ($group) {
        $typesStr = if ($group.groupTypes) { $group.groupTypes -join ", " } else { "Ninguno" }
        Write-Host "Grupo '$GroupName' encontrado con ID: $($group.id)"
        Write-Host "  - Tipo (groupTypes): $typesStr"
        Write-Host "  - Habilitado para Mail (mailEnabled): $($group.mailEnabled)"
        Write-Host "  - Seguridad (securityEnabled): $($group.securityEnabled)"
        Write-Host "  - Sincronizado desde On-Premises (onPremisesSyncEnabled): $($group.onPremisesSyncEnabled)"
        
        # Validar si es un tipo de grupo soportado
        if ($group.securityEnabled -ne $true -and ($group.groupTypes -notcontains "Unified")) {
            Write-Warning "¡ATENCION! El grupo '$GroupName' no parece ser un grupo de seguridad estándar ni un grupo de Microsoft 365. Las consultas de miembros podrían fallar (404/400)."
        }
        if ($group.onPremisesSyncEnabled -eq $true) {
            Write-Warning "¡ATENCION! El grupo '$GroupName' está sincronizado desde on-premises. Su membresía es de solo lectura en la nube."
        }
        return $group.id
    }

    if ($DryRun) {
        Write-Host "[SIMULACION] El grupo '$GroupName' no existe. Se creara en el modo real."
        return "SIMULATION_GROUP_ID_$($GroupName -replace '[^a-zA-Z0-9]', '')"
    }
    else {
        Write-Host "El grupo '$GroupName' no existe. Creandolo..."
        $groupParams = @{
            DisplayName     = $GroupName
            Description     = $Description
            MailEnabled     = $false
            MailNickname    = ($GroupName -replace '[^a-zA-Z0-9]', '')
            SecurityEnabled = $true
            GroupTypes      = @()
        }
        
        try {
            $body = $groupParams | ConvertTo-Json -Depth 5
            $newGroupJson = Invoke-GraphRest -Method POST -Uri "v1.0/groups" -Body $body
            Write-Host "Grupo creado con exito. ID: $($newGroupJson.id)"
            return $newGroupJson.id
        }
        catch {
            Write-Error "No se pudo crear el grupo '$GroupName': $($_.Exception.Message)"
            throw $_
        }
    }
}

# 2. Obtener o crear los grupos
$groupWithId = Get-OrCreateGroup -GroupName $GroupNameWithDevices -Description "Usuarios con al menos un dispositivo gestionado en Intune (Sincronizacion Automatica)"
$groupWithoutId = Get-OrCreateGroup -GroupName $GroupNameWithoutDevices -Description "Usuarios sin ningun dispositivo gestionado en Intune (Sincronizacion Automatica)"
Write-Output ""

# 3. Obtener dispositivos de Intune
Write-Output "Obteniendo lista de dispositivos gestionados en Intune..."
$devices = @()
try {
    $devices = Invoke-GraphPaginatedRequest -Uri "v1.0/deviceManagement/managedDevices?`$select=id,userId,userPrincipalName,operatingSystem"
    Write-Output "Se encontraron $($devices.Count) dispositivos registrados en Intune."
}
catch {
    Write-Error "Fallo critico al leer los dispositivos gestionados en Intune. Abortando."
    exit 1
}

# Hash de dispositivos
$userWithDeviceIds = @{}
foreach ($d in $devices) {
    $uid = $d.userId
    if (-not [string]::IsNullOrEmpty($uid)) {
        $userWithDeviceIds[$uid] = $true
    }
}
Write-Output "Usuarios unicos con dispositivos gestionados: $($userWithDeviceIds.Count)`n"

# 4. Obtener usuarios del tenant
Write-Output "Obteniendo lista de usuarios del tenant..."
$allUsersRaw = @()
try {
    $allUsersRaw = Invoke-GraphPaginatedRequest -Uri "v1.0/users?`$select=id,displayName,userPrincipalName,accountEnabled,userType,mail"
    Write-Output "Se encontraron $($allUsersRaw.Count) usuarios en el tenant."
}
catch {
    Write-Error "Fallo critico al leer los usuarios del tenant. Abortando."
    exit 1
}

$allUsersMap = @{}
foreach ($u in $allUsersRaw) {
    $uid = $u.id
    if (-not [string]::IsNullOrEmpty($uid)) { $allUsersMap[$uid] = $u }
}

# Filtrar usuarios elegibles
$eligibleUsers = @()
foreach ($u in $allUsersRaw) {
    $uid = $u.id
    $upn = $u.userPrincipalName
    $enabled = $u.accountEnabled
    $uType = $u.userType
    $mail = $u.mail
    $displayName = $u.displayName
    
    # Excepcion CEOs
    if ($ForceWithDeviceEmails -contains $upn) {
        Write-Output "DEBUG: Usuario '$upn' es forzado (CEO). Pasa filtro."
        $eligibleUsers += $u
        continue
    }
    
    if ($null -eq $enabled) { $enabled = $true }
    if ($null -eq $uType) { $uType = "Member" }

    $keep = $true
    $filterReason = ""
    
    if ($ExcludeGuests -and ($uType -ne "Member" -or $upn -like "*#EXT#*")) { 
        $keep = $false 
        $filterReason += " [No es Member o es Invitado, tipo=$uType]"
    }
    if ($OnlyActiveUsers -and -not $enabled) { 
        $keep = $false 
        $filterReason += " [Inactivo]"
    }
    if ($keep -and $RequireMail -and [string]::IsNullOrEmpty($mail)) {
        $keep = $false
        $filterReason += " [Sin mail]"
    }
    if ($keep -and -not [string]::IsNullOrEmpty($AllowedDomain)) {
        if ($upn -notlike "*@$AllowedDomain") { 
            $keep = $false 
            $filterReason += " [Dominio no permitido: $upn]"
        }
    }
    if ($keep -and $ExcludedUserPatterns) {
        $username = $upn.Split('@')[0]
        foreach ($pattern in $ExcludedUserPatterns) {
            if ($username -like "*$pattern*") {
                $keep = $false
                $filterReason += " [Excluido por patron: $pattern]"
                break
            }
        }
    }
    if ($keep -and $EnforceNamingPattern) {
        if ($ExceptedNamingPatternEmails -contains $upn) {
            Write-Output "DEBUG: Usuario '$upn' es excepcion del patron de nombre. Pasa filtro."
        }
        elseif (-not (Test-UserNamingPattern -UserPrincipalName $upn -DisplayName $displayName)) {
            $keep = $false
            $filterReason += " [Fallo patron de nombre (UPN=$upn, DisplayName=$displayName)]"
        }
    }
    
    if (-not $keep) {
        Write-Output "DEBUG: Usuario '$upn' filtrado debido a:$filterReason"
    }
    else {
        Write-Output "DEBUG: Usuario '$upn' pasa el filtro."
        $eligibleUsers += $u
    }
}
$exclusionesStr = if ($ExcludedUserPatterns) { $ExcludedUserPatterns -join ',' } else { "Ninguna" }
Write-Output "Usuarios elegibles para evaluar tras filtros (Activos=$OnlyActiveUsers, ExcluirInvitados=$ExcludeGuests, Dominio=$AllowedDomain, Exclusiones=$exclusionesStr, ValidarPatronNombre=$EnforceNamingPattern, RequerirEmail=$RequireMail): $($eligibleUsers.Count)`n"

# 5. Obtener miembros actuales de ambos grupos
$currentMembersWith = @()
$currentMembersWithout = @()

if ($groupWithId -notlike "SIMULATION_GROUP_ID_*") {
    Write-Output "Obteniendo miembros actuales del grupo '$GroupNameWithDevices'..."
    try {
        $currentMembersWith = Invoke-GraphPaginatedRequest -Uri "v1.0/groups/$groupWithId/members?`$select=id,userPrincipalName,displayName"
        Write-Output "Miembros actuales en '$GroupNameWithDevices': $($currentMembersWith.Count)"
    }
    catch {
        Write-Warning "No se pudo leer la membresia de '$GroupNameWithDevices'."
    }
}

if ($groupWithoutId -notlike "SIMULATION_GROUP_ID_*") {
    Write-Output "Obteniendo miembros actuales del grupo '$GroupNameWithoutDevices'..."
    try {
        $currentMembersWithout = Invoke-GraphPaginatedRequest -Uri "v1.0/groups/$groupWithoutId/members?`$select=id,userPrincipalName,displayName"
        Write-Output "Miembros actuales en '$GroupNameWithoutDevices': $($currentMembersWithout.Count)"
    }
    catch {
        Write-Warning "No se pudo leer la membresia de '$GroupNameWithoutDevices'."
    }
}
Write-Output ""

# Crear tablas hash de miembros actuales para busquedas eficientes O(1)
$currentMembersWithIds = @{}
foreach ($m in $currentMembersWith) {
    $mid = $m.id
    if (-not [string]::IsNullOrEmpty($mid)) { $currentMembersWithIds[$mid] = $true }
}

$currentMembersWithoutIds = @{}
foreach ($m in $currentMembersWithout) {
    $mid = $m.id
    if (-not [string]::IsNullOrEmpty($mid)) { $currentMembersWithoutIds[$mid] = $true }
}

# 6. Calcular diferencias (Deltas)
$toAddWith = @()
$toRemoveWith = @()
$toAddWithout = @()
$toRemoveWithout = @()

foreach ($u in $eligibleUsers) {
    $uid = $u.id
    $upn = $u.userPrincipalName
    $name = $u.displayName
    
    $hasDevice = $userWithDeviceIds.ContainsKey($uid) -or ($ForceWithDeviceEmails -contains $upn)
    
    if ($hasDevice) {
        $reasonAdd = if ($ForceWithDeviceEmails -contains $upn) { "Forzado (CEO)" } else { "Tiene dispositivo en Intune" }
        $reasonRemove = if ($ForceWithDeviceEmails -contains $upn) { "Forzado (CEO, pasa a con dispositivo)" } else { "Tiene dispositivo en Intune (pasa a con dispositivo)" }
        
        if (-not $currentMembersWithIds.ContainsKey($uid)) {
            $toAddWith += [PSCustomObject]@{ Id = $uid; UPN = $upn; DisplayName = $name; Reason = $reasonAdd }
        }
        if ($currentMembersWithoutIds.ContainsKey($uid)) {
            $toRemoveWithout += [PSCustomObject]@{ Id = $uid; UPN = $upn; DisplayName = $name; Reason = $reasonRemove }
        }
    }
    else {
        if (-not $currentMembersWithoutIds.ContainsKey($uid)) {
            $toAddWithout += [PSCustomObject]@{ Id = $uid; UPN = $upn; DisplayName = $name; Reason = "No tiene dispositivo en Intune" }
        }
        if ($currentMembersWithIds.ContainsKey($uid)) {
            $toRemoveWith += [PSCustomObject]@{ Id = $uid; UPN = $upn; DisplayName = $name; Reason = "No tiene dispositivo en Intune (pasa a sin dispositivo)" }
        }
    }
}

# Limpiar no elegibles
$eligibleUserIds = @{}
foreach ($u in $eligibleUsers) {
    $uid = $u.id
    if (-not [string]::IsNullOrEmpty($uid)) { $eligibleUserIds[$uid] = $true }
}

foreach ($mem in $currentMembersWith) {
    $mid = $mem.id
    if (-not $eligibleUserIds.ContainsKey($mid)) {
        $mupn = $mem.userPrincipalName
        $mname = $mem.displayName
        
        $finalUpn = if ($mupn) { $mupn } else { $mid }
        $finalName = if ($mname) { $mname } else { "Inactivo/Invitado" }
        
        $reason = "No elegible"
        $rawUser = $allUsersMap[$mid]
        if ($null -eq $rawUser) {
            $reason = "Eliminado del tenant"
        }
        else {
            $uEnabled = $rawUser.accountEnabled
            $uUserType = $rawUser.userType
            $uMail = $rawUser.mail
            
            $isExcludedPattern = $false
            if ($ExcludedUserPatterns) {
                $username = $finalUpn.Split('@')[0]
                foreach ($pat in $ExcludedUserPatterns) {
                    if ($username -like "*$pat*") {
                        $isExcludedPattern = $true
                        break
                    }
                }
            }
            
            if ($OnlyActiveUsers -and $uEnabled -eq $false) {
                $reason = "Cuenta deshabilitada"
            }
            elseif ($ExcludeGuests -and ($uUserType -ne "Member" -or $finalUpn -like "*#EXT#*")) {
                $reason = "Usuario tipo Guest/Invitado"
            }
            elseif ($RequireMail -and [string]::IsNullOrEmpty($uMail)) {
                $reason = "Sin correo electronico"
            }
            elseif (-not [string]::IsNullOrEmpty($AllowedDomain) -and $finalUpn -notlike "*@$AllowedDomain") {
                $reason = "Dominio no permitido"
            }
            elseif ($isExcludedPattern) {
                $reason = "Excluido por palabra clave (admin/test/poc)"
            }
            elseif ($EnforceNamingPattern -and $ExceptedNamingPatternEmails -notcontains $finalUpn -and -not (Test-UserNamingPattern -UserPrincipalName $finalUpn -DisplayName $finalName)) {
                $reason = "No cumple patron de nombre"
            }
        }
        
        $toRemoveWith += [PSCustomObject]@{ Id = $mid; UPN = $finalUpn; DisplayName = $finalName; Reason = $reason }
    }
}

foreach ($mem in $currentMembersWithout) {
    $mid = $mem.id
    if (-not $eligibleUserIds.ContainsKey($mid)) {
        $mupn = $mem.userPrincipalName
        $mname = $mem.displayName
        
        $finalUpn = if ($mupn) { $mupn } else { $mid }
        $finalName = if ($mname) { $mname } else { "Inactivo/Invitado" }
        
        $reason = "No elegible"
        $rawUser = $allUsersMap[$mid]
        if ($null -eq $rawUser) {
            $reason = "Eliminado del tenant"
        }
        else {
            $uEnabled = $rawUser.accountEnabled
            $uUserType = $rawUser.userType
            $uMail = $rawUser.mail
            
            $isExcludedPattern = $false
            if ($ExcludedUserPatterns) {
                $username = $finalUpn.Split('@')[0]
                foreach ($pat in $ExcludedUserPatterns) {
                    if ($username -like "*$pat*") {
                        $isExcludedPattern = $true
                        break
                    }
                }
            }
            
            if ($OnlyActiveUsers -and $uEnabled -eq $false) {
                $reason = "Cuenta deshabilitada"
            }
            elseif ($ExcludeGuests -and ($uUserType -ne "Member" -or $finalUpn -like "*#EXT#*")) {
                $reason = "Usuario tipo Guest/Invitado"
            }
            elseif ($RequireMail -and [string]::IsNullOrEmpty($uMail)) {
                $reason = "Sin correo electronico"
            }
            elseif (-not [string]::IsNullOrEmpty($AllowedDomain) -and $finalUpn -notlike "*@$AllowedDomain") {
                $reason = "Dominio no permitido"
            }
            elseif ($isExcludedPattern) {
                $reason = "Excluido por palabra clave (admin/test/poc)"
            }
            elseif ($EnforceNamingPattern -and $ExceptedNamingPatternEmails -notcontains $finalUpn -and -not (Test-UserNamingPattern -UserPrincipalName $finalUpn -DisplayName $finalName)) {
                $reason = "No cumple patron de nombre"
            }
        }
        
        $toRemoveWithout += [PSCustomObject]@{ Id = $mid; UPN = $finalUpn; DisplayName = $finalName; Reason = $reason }
    }
}

# 7. Resumen de acciones
Write-Output "========================================= RESUMEN DE CAMBIOS ========================================="
Write-Output "Grupo: '$GroupNameWithDevices'"
Write-Output "  [+] A adicionar: $($toAddWith.Count)"
Write-Output "  [-] A eliminar: $($toRemoveWith.Count)"
Write-Output "Grupo: '$GroupNameWithoutDevices'"
Write-Output "  [+] A adicionar: $($toAddWithout.Count)"
Write-Output "  [-] A eliminar: $($toRemoveWithout.Count)"
Write-Output "======================================================================================================"
Write-Output ""

# Añadir miembro
function Add-GroupMember {
    param(
        [string]$GroupId,
        [string]$UserId,
        [string]$UserUPN,
        [string]$GroupName,
        [string]$Reason
    )
    if ($DryRun) {
        Write-Output "[SIMULACION] Adicionaria '$UserUPN' al grupo '$GroupName' (Motivo: $Reason)"
        return
    }
    try {
        $body = @{
            '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId"
        } | ConvertTo-Json
        $null = Invoke-GraphRest -Method POST -Uri "v1.0/groups/$GroupId/members/`$ref" -Body $body
        Write-Output "[+] Adicionado con exito: '$UserUPN' al grupo '$GroupName' (Motivo: $Reason)"
    }
    catch {
        if ($_.Exception.Message -match "One or more added object references already exist") {
            Write-Output "[i] El usuario '$UserUPN' ya es miembro de '$GroupName'"
        }
        else {
            Write-Warning "[X] Error al adicionar '$UserUPN' (Motivo: $Reason): $($_.Exception.Message)"
        }
    }
}

# Eliminar miembro
function Remove-GroupMember {
    param(
        [string]$GroupId,
        [string]$UserId,
        [string]$UserUPN,
        [string]$GroupName,
        [string]$Reason
    )
    if ($DryRun) {
        Write-Output "[SIMULACION] Eliminaria '$UserUPN' del grupo '$GroupName' (Motivo: $Reason)"
        return
    }
    try {
        $null = Invoke-GraphRest -Method DELETE -Uri "v1.0/groups/$GroupId/members/$UserId/`$ref"
        Write-Output "[-] Eliminado con exito: '$UserUPN' del grupo '$GroupName' (Motivo: $Reason)"
    }
    catch {
        Write-Warning "[X] Error al eliminar '$UserUPN' (Motivo: $Reason): $($_.Exception.Message)"
    }
}

# 8. Aplicar cambios
if ($toAddWith.Count -gt 0 -or $toRemoveWith.Count -gt 0) {
    Write-Output "Actualizando membresias del grupo '$GroupNameWithDevices'..."
    foreach ($u in $toAddWith) {
        Add-GroupMember -GroupId $groupWithId -UserId $u.Id -UserUPN $u.UPN -GroupName $GroupNameWithDevices -Reason $u.Reason
    }
    foreach ($u in $toRemoveWith) {
        Remove-GroupMember -GroupId $groupWithId -UserId $u.Id -UserUPN $u.UPN -GroupName $GroupNameWithDevices -Reason $u.Reason
    }
}

if ($toAddWithout.Count -gt 0 -or $toRemoveWithout.Count -gt 0) {
    Write-Output "Actualizando membresias del grupo '$GroupNameWithoutDevices'..."
    foreach ($u in $toAddWithout) {
        Add-GroupMember -GroupId $groupWithoutId -UserId $u.Id -UserUPN $u.UPN -GroupName $GroupNameWithoutDevices -Reason $u.Reason
    }
    foreach ($u in $toRemoveWithout) {
        Remove-GroupMember -GroupId $groupWithoutId -UserId $u.Id -UserUPN $u.UPN -GroupName $GroupNameWithoutDevices -Reason $u.Reason
    }
}

Write-Output "`nSincronizacion finalizada correctamente."
if ($DryRun) {
    Write-Output "Recuerda: Los cambios anteriores fueron simulados. Configure el parametro Commit a `$true para aplicarlos."
}