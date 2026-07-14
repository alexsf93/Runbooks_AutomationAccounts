<#
.SYNOPSIS
    Sincroniza dinamicamente dos grupos en Microsoft Entra ID segun si los usuarios tienen o no dispositivos gestionados en Intune.

.DESCRIPTION
    Este runbook esta diseñado para ejecutarse en Azure Automation. Se conecta a Microsoft Graph
    utilizando la Managed Identity (Identidad Administrada) del Automation Account y distribuye a los 
    usuarios del tenant en dos grupos:
      - "NaxvanCorp - Gente con dispositivo corporativo" (usuarios con al menos 1 dispositivo en Intune)
      - "NaxvanCorp - Gente sin dispositivo corporativo" (usuarios sin ningun dispositivo en Intune)

    NOTA IMPORTANTE: Los grupos de destino deben ser de tipo Microsoft 365 (Unified Groups) para poder ser gestionados 
    al 100% mediante Microsoft Graph API sin requerir roles administrativos de Exchange Online. Si los grupos no existen 
    en el tenant, el script los creará automáticamente en la primera ejecución.

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
    Patrones de nombres de usuario a excluir de la evaluacion. Valor por defecto: "admin", "test", "prueba", "poc", "noreply", "no-reply".

.PARAMETER EnforceNamingPattern
    Validar que el UPN cumpla con el patron de nombre de usuario de la organizacion. Valor por defecto: True.

.PARAMETER RequireMail
    Requerir que el usuario tenga un correo electronico configurado. Valor por defecto: True.

.PARAMETER ForceWithDeviceEmails
    Lista de correos de usuarios (ej. CEOs) que se forzara a considerar con dispositivo.

.PARAMETER ExceptedNamingPatternEmails
    Lista de correos de usuarios exceptuados de cumplir con la validacion del patron de nombre.

.REQUIREMENTS
    - Los grupos de destino deben ser de tipo Microsoft 365 (Unified Groups) para permitir que la API de Microsoft Graph actualice su membresía sin requerir privilegios de Exchange. El script intentará crearlos automáticamente si no existen en el tenant.
    - Módulos instalados en el Automation Account:
        * Az.Accounts
    - Managed Identity con los siguientes permisos de Microsoft Graph (tipo Aplicación / Application permissions):
        1. DeviceManagementManagedDevices.Read.All (Lectura de dispositivos Intune)
        2. User.Read.All (Lectura de usuarios del tenant)
        3. Group.ReadWrite.All (Creación, lectura y edición de grupos de Microsoft 365 y sus miembros)

    Nota: Al utilizar la API de Microsoft Graph para la gestión de grupos y miembros, no se requiere ningún rol de Microsoft Entra ID (como Administrador de Exchange) ni permisos de Exchange Online (como Exchange.ManageAsApp).

    Script PowerShell para asignar los permisos a la Managed Identity (Ejecutar en Azure Cloud Shell):
    --------------------------------------------------------------------------------------------------
    # Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"
    #
    # $MSIName    = "Nombre-De-Tu-Automation-Account"
    # $MSI        = Get-MgServicePrincipal -Filter "displayName eq '$MSIName'"
    #
    # # Asignar permisos de Microsoft Graph
    # $GraphRoles = @("DeviceManagementManagedDevices.Read.All", "User.Read.All", "Group.ReadWrite.All")
    # $GraphAppId = "00000003-0000-0000-c000-000000000000"
    # $GraphSP    = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'"
    #
    # foreach ($Role in $GraphRoles) {
    #     $AppRole = $GraphSP.AppRoles | Where-Object { $_.Value -eq $Role -and $_.AllowedMemberTypes -contains "Application" }
    #     if ($AppRole) {
    #         New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MSI.Id -PrincipalId $MSI.Id -ResourceId $GraphSP.Id -AppRoleId $AppRole.Id
    #     }
    # }

.NOTES
    Name: Runbook_EntraID-SyncDeviceGroups-M365Groups.ps1
    Author: Alejandro Suarez (@alexsf93)
    Version: 3.1.0
    Date: 2026-07-14
#>

param(
    [string]$GroupNameWithDevices = "NaxvanCorp - Gente con dispositivo corporativo",
    [string]$GroupNameWithoutDevices = "NaxvanCorp - Gente sin dispositivo corporativo",
    [bool]$Commit = $false,
    [bool]$ExcludeGuests = $true,
    [bool]$OnlyActiveUsers = $true,
    [string]$AllowedDomain = "naxvan.es",
    [object]$ExcludedUserPatterns = $null,
    [bool]$EnforceNamingPattern = $true,
    [bool]$RequireMail = $true,
    [object]$ForceWithDeviceEmails = $null,
    [object]$ExceptedNamingPatternEmails = $null
)

try {
    Write-Verbose "Iniciando el script. Estableciendo arrays por defecto..."
    
    # Asignar valores por defecto o procesar comas para ExcludedUserPatterns
    if ($null -eq $ExcludedUserPatterns -or [string]::IsNullOrEmpty($ExcludedUserPatterns)) {
        $ExcludedUserPatterns = [string[]]@("admin", "test", "prueba", "poc", "noreply", "no-reply")
    }
    elseif ($ExcludedUserPatterns -is [string]) {
        $ExcludedUserPatterns = [string[]]($ExcludedUserPatterns.Split([char[]]',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() })
    }

    # Asignar valores por defecto o procesar comas para ForceWithDeviceEmails
    if ($null -eq $ForceWithDeviceEmails -or [string]::IsNullOrEmpty($ForceWithDeviceEmails)) {
        $ForceWithDeviceEmails = [string[]]@("alexsf93@naxvan.es")
    }
    elseif ($ForceWithDeviceEmails -is [string]) {
        $ForceWithDeviceEmails = [string[]]($ForceWithDeviceEmails.Split([char[]]',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() })
    }

    # Asignar valores por defecto o procesar comas para ExceptedNamingPatternEmails
    if ($null -eq $ExceptedNamingPatternEmails -or [string]::IsNullOrEmpty($ExceptedNamingPatternEmails)) {
        $ExceptedNamingPatternEmails = [string[]]@(
            "alejandro@naxvan.es"
        )
    }
    elseif ($ExceptedNamingPatternEmails -is [string]) {
        $ExceptedNamingPatternEmails = [string[]]($ExceptedNamingPatternEmails.Split([char[]]',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() })
    }

    Write-Verbose "Sanitizando parametros..."
    if ($AllowedDomain) { 
        $AllowedDomain = $AllowedDomain.Trim("`"'") 
        Write-Verbose "AllowedDomain sanitizado: $AllowedDomain"
    }
    if ($GroupNameWithDevices) { 
        $GroupNameWithDevices = $GroupNameWithDevices.Trim("`"'") 
        Write-Verbose "GroupNameWithDevices sanitizado: $GroupNameWithDevices"
    }
    if ($GroupNameWithoutDevices) { 
        $GroupNameWithoutDevices = $GroupNameWithoutDevices.Trim("`"'") 
        Write-Verbose "GroupNameWithoutDevices sanitizado: $GroupNameWithoutDevices"
    }
    
    if ($ExcludedUserPatterns) {
        Write-Verbose "Sanitizando ExcludedUserPatterns..."
        $ExcludedUserPatterns = [string[]]$ExcludedUserPatterns
        for ($i = 0; $i -lt $ExcludedUserPatterns.Count; $i++) {
            if ($ExcludedUserPatterns[$i]) { $ExcludedUserPatterns[$i] = $ExcludedUserPatterns[$i].Trim("`"'") }
        }
    }
    if ($ForceWithDeviceEmails) {
        Write-Verbose "Sanitizando ForceWithDeviceEmails..."
        $ForceWithDeviceEmails = [string[]]$ForceWithDeviceEmails
        for ($i = 0; $i -lt $ForceWithDeviceEmails.Count; $i++) {
            if ($ForceWithDeviceEmails[$i]) { $ForceWithDeviceEmails[$i] = $ForceWithDeviceEmails[$i].Trim("`"'") }
        }
    }
    if ($ExceptedNamingPatternEmails) {
        Write-Verbose "Sanitizando ExceptedNamingPatternEmails..."
        $ExceptedNamingPatternEmails = [string[]]$ExceptedNamingPatternEmails
        for ($i = 0; $i -lt $ExceptedNamingPatternEmails.Count; $i++) {
            if ($ExceptedNamingPatternEmails[$i]) { $ExceptedNamingPatternEmails[$i] = $ExceptedNamingPatternEmails[$i].Trim("`"'") }
        }
    }

    Write-Output "[+] Conectando a Microsoft Graph..."
    
    Write-Verbose "Cargando modulo Az.Accounts..."
    Import-Module Az.Accounts -ErrorAction Stop
    Write-Verbose "Modulo Az.Accounts cargado con exito."
    
    $null = Connect-AzAccount -Identity -ErrorAction Stop
    Write-Verbose "Conexion a Azure establecida con exito."
    
    Write-Verbose "Obteniendo Access Token para Microsoft Graph..."
    $TokenResult = (Get-AzAccessToken -ResourceTypeName MSGraph -ErrorAction Stop).Token
    Write-Verbose "Access Token obtenido."
    
    if ($TokenResult -is [System.Security.SecureString]) {
        $Token = [System.Net.NetworkCredential]::new('', $TokenResult).Password
    }
    else {
        $Token = $TokenResult
    }
    
    $RequestHeaders = @{
        Authorization  = "Bearer $Token"
        "Content-Type" = "application/json; charset=utf-8"
    }
    Write-Output "    -> Conectado a Microsoft Graph [OK]"
    Write-Output ""
}
catch {
    Write-Verbose "Ocurrio un fallo critico en la inicializacion: $($_.Exception.Message)"
    Write-Error "Fallo critico en la inicializacion o conexion: $($_.Exception.Message)"
    throw $_
}

$DryRun = -not $Commit

Write-Output "--------------------------------------------------------------------------------"
Write-Output " Sincronizacion de Grupos de Usuarios segun Dispositivos Intune"
Write-Output "--------------------------------------------------------------------------------"
if ($DryRun) {
    Write-Output " MODO: SIMULACION (Dry Run) - No se realizaran cambios en el tenant."
}
else {
    Write-Output " MODO: PRODUCCION (Commit) - Se aplicaran cambios reales en el tenant."
}
Write-Output "--------------------------------------------------------------------------------`n"

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
    
    $parts = $cleanName.Split([char[]]' ', [System.StringSplitOptions]::RemoveEmptyEntries)
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
function Get-RequiredGroup {
    param(
        [string]$GroupName
    )
    
    Write-Verbose "Buscando el grupo '$GroupName'..."
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
        Write-Verbose "Grupo '$GroupName' encontrado con ID: $($group.id)"
        Write-Verbose "  - Tipo (groupTypes): $typesStr"
        Write-Verbose "  - Habilitado para Mail (mailEnabled): $($group.mailEnabled)"
        Write-Verbose "  - Seguridad (securityEnabled): $($group.securityEnabled)"
        Write-Verbose "  - Sincronizado desde On-Premises (onPremisesSyncEnabled): $($group.onPremisesSyncEnabled)"
        
        # Validar si es un tipo de grupo soportado
        if ($group.groupTypes -notcontains "Unified") {
            Write-Warning "¡ATENCION! El grupo '$GroupName' no es de tipo Microsoft 365 (Unified Group). Si es un grupo de seguridad habilitado para correo (Mail-enabled Security Group), la actualización de miembros mediante Graph API fallará con error 400 (Bad Request) debido a restricciones de Microsoft."
        }
        if ($group.onPremisesSyncEnabled -eq $true) {
            Write-Warning "¡ATENCION! El grupo '$GroupName' está sincronizado desde on-premises. Su membresía es de solo lectura en la nube."
        }
        return $group.id
    }

    # Si no existe y estamos en DryRun, simulamos su creacion
    if ($DryRun) {
        Write-Verbose "[SIMULACION] El grupo '$GroupName' no existe."
        return "SIMULATION_GROUP_ID_$($GroupName -replace '[^a-zA-Z0-9]', '')"
    }

    # Intentar crearlo en Microsoft Graph como grupo de Microsoft 365 (Unified)
    Write-Verbose "El grupo '$GroupName' no existe en el tenant. Intentando crearlo automaticamente en Microsoft Graph como Grupo de Microsoft 365..."
    try {
        # Generar un mailNickname válido a partir del nombre del grupo
        # Solo caracteres alfanuméricos
        $mailNickname = ($GroupName -replace '[^a-zA-Z0-9]', '').ToLower()
        if ($mailNickname.Length -gt 64) {
            $mailNickname = $mailNickname.Substring(0, 64)
        }
        
        $Body = @{
            displayName = $GroupName
            description = "Grupo de sincronizacion de dispositivos Intune ($GroupName)"
            groupTypes = @("Unified")
            mailEnabled = $true
            securityEnabled = $false
            mailNickname = $mailNickname
            resourceBehaviorOptions = @("WelcomeEmailDisabled")
        } | ConvertTo-Json -Compress
        
        $newGroup = Invoke-GraphRest -Method POST -Uri "v1.0/groups" -Body $Body
        $groupId = $newGroup.id
        
        Write-Verbose "[+] Grupo de Microsoft 365 '$GroupName' creado con exito con ID: $groupId. Esperando 15 segundos para la replicacion..."
        Start-Sleep -Seconds 15
        return $groupId
    }
    catch {
        throw "Fallo critico: No se pudo crear el grupo de Microsoft 365 '$GroupName' en Microsoft Graph. Error: $($_.Exception.Message)"
    }
}

# 2. Obtener los grupos requeridos
Write-Output "[+] Buscando grupos de destino..."
$groupWithId = Get-RequiredGroup -GroupName $GroupNameWithDevices
$groupWithoutId = Get-RequiredGroup -GroupName $GroupNameWithoutDevices
Write-Verbose ""

# 3. Obtener dispositivos de Intune
Write-Output "[+] Consultando dispositivos en Microsoft Intune..."
$devices = @()
try {
    $devices = Invoke-GraphPaginatedRequest -Uri "v1.0/deviceManagement/managedDevices?`$select=id,userId,userPrincipalName,operatingSystem"
}
catch {
    throw "Fallo critico al leer los dispositivos gestionados en Intune. Error: $($_.Exception.Message)"
}

# Hash de dispositivos
$userWithDeviceIds = @{}
foreach ($d in $devices) {
    $uid = $d.userId
    if (-not [string]::IsNullOrEmpty($uid)) {
        $userWithDeviceIds[$uid] = $true
    }
}
Write-Output "    -> Dispositivos detectados: $($devices.Count) ($($userWithDeviceIds.Count) usuarios unicos)"

# 4. Obtener usuarios del tenant
Write-Output "[+] Consultando usuarios de Microsoft Entra ID..."
$allUsersRaw = @()
try {
    $allUsersRaw = Invoke-GraphPaginatedRequest -Uri "v1.0/users?`$select=id,displayName,userPrincipalName,accountEnabled,userType,mail"
}
catch {
    throw "Fallo critico al leer los usuarios del tenant. Error: $($_.Exception.Message)"
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
        Write-Verbose "Usuario '$upn' es forzado (CEO). Pasa filtro."
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
            Write-Verbose "Usuario '$upn' es excepcion del patron de nombre. Pasa filtro."
        }
        elseif (-not (Test-UserNamingPattern -UserPrincipalName $upn -DisplayName $displayName)) {
            $keep = $false
            $filterReason += " [Fallo patron de nombre (UPN=$upn, DisplayName=$displayName)]"
        }
    }
    
    if (-not $keep) {
        Write-Verbose "Usuario '$upn' filtrado debido a:$filterReason"
    }
    else {
        Write-Verbose "Usuario '$upn' pasa el filtro."
        $eligibleUsers += $u
    }
}
$exclusionesStr = if ($ExcludedUserPatterns) { $ExcludedUserPatterns -join ',' } else { "Ninguna" }
Write-Output "[+] Evaluando usuarios del tenant..."
Write-Output "    -> Usuarios elegibles tras filtros: $($eligibleUsers.Count) (de $($allUsersRaw.Count) totales)"

# 5. Obtener miembros actuales de ambos grupos
Write-Output "[+] Consultando miembros actuales de los grupos..."
$currentMembersWith = @()
$currentMembersWithout = @()

if ($groupWithId -notlike "*SIMULATION_GROUP_ID_*") {
    try {
        $currentMembersWith = Invoke-GraphPaginatedRequest -Uri "v1.0/groups/$groupWithId/members?`$select=id,userPrincipalName,displayName"
        Write-Output "    -> Miembros en '$GroupNameWithDevices': $($currentMembersWith.Count)"
    }
    catch {
        Write-Warning "No se pudo leer la membresia de '$GroupNameWithDevices'."
    }
}

if ($groupWithoutId -notlike "*SIMULATION_GROUP_ID_*") {
    try {
        $currentMembersWithout = Invoke-GraphPaginatedRequest -Uri "v1.0/groups/$groupWithoutId/members?`$select=id,userPrincipalName,displayName"
        Write-Output "    -> Miembros en '$GroupNameWithoutDevices': $($currentMembersWithout.Count)"
    }
    catch {
        Write-Warning "No se pudo leer la membresia de '$GroupNameWithoutDevices'."
    }
}
Write-Verbose ""

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
Write-Output "--------------------------------------------------------------------------------"
Write-Output "                              RESUMEN DE CAMBIOS"
Write-Output "--------------------------------------------------------------------------------"
Write-Output " Grupo: '$GroupNameWithDevices'"
Write-Output "   [+] Miembros a añadir: $($toAddWith.Count)"
Write-Output "   [-] Miembros a eliminar: $($toRemoveWith.Count)"
Write-Output " Grupo: '$GroupNameWithoutDevices'"
Write-Output "   [+] Miembros a añadir: $($toAddWithout.Count)"
Write-Output "   [-] Miembros a eliminar: $($toRemoveWithout.Count)"
Write-Output "--------------------------------------------------------------------------------`n"

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
        Write-Output "    [SIMULACION] Añadir '$UserUPN' ($Reason)"
        return
    }
    try {
        $Body = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId"
        } | ConvertTo-Json
        
        $null = Invoke-GraphRest -Method POST -Uri "v1.0/groups/$GroupId/members/`$ref" -Body $Body
        Write-Output "    [+] Añadido: '$UserUPN' ($Reason)"
    }
    catch {
        if ($_.Exception.Message -match "already exists" -or $_.Exception.Message -match "ya existe" -or $_.Exception.Message -match "One or more added object references already exist") {
            Write-Output "    [i] Ya es miembro: '$UserUPN'"
        }
        else {
            Write-Warning "    [X] Error al añadir '$UserUPN' ($Reason): $($_.Exception.Message)"
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
        Write-Output "    [SIMULACION] Eliminar '$UserUPN' ($Reason)"
        return
    }
    try {
        $null = Invoke-GraphRest -Method DELETE -Uri "v1.0/groups/$GroupId/members/$UserId/`$ref"
        Write-Output "    [-] Eliminado: '$UserUPN' ($Reason)"
    }
    catch {
        Write-Warning "    [X] Error al eliminar '$UserUPN' ($Reason): $($_.Exception.Message)"
    }
}

# 8. Aplicar cambios
if ($toAddWith.Count -gt 0 -or $toRemoveWith.Count -gt 0) {
    Write-Output "[+] Sincronizando miembros en el grupo '$GroupNameWithDevices'..."
    foreach ($u in $toAddWith) {
        Add-GroupMember -GroupId $groupWithId -UserId $u.Id -UserUPN $u.UPN -GroupName $GroupNameWithDevices -Reason $u.Reason
    }
    foreach ($u in $toRemoveWith) {
        Remove-GroupMember -GroupId $groupWithId -UserId $u.Id -UserUPN $u.UPN -GroupName $GroupNameWithDevices -Reason $u.Reason
    }
}

if ($toAddWithout.Count -gt 0 -or $toRemoveWithout.Count -gt 0) {
    Write-Output "[+] Sincronizando miembros en el grupo '$GroupNameWithoutDevices'..."
    foreach ($u in $toAddWithout) {
        Add-GroupMember -GroupId $groupWithoutId -UserId $u.Id -UserUPN $u.UPN -GroupName $GroupNameWithoutDevices -Reason $u.Reason
    }
    foreach ($u in $toRemoveWithout) {
        Remove-GroupMember -GroupId $groupWithoutId -UserId $u.Id -UserUPN $u.UPN -GroupName $GroupNameWithoutDevices -Reason $u.Reason
    }
}

Write-Output "Sincronizacion finalizada correctamente."
if ($DryRun) {
    Write-Output "Nota: Los cambios anteriores fueron simulados (Dry Run). Para aplicarlos, configure Commit a `$true."
}

# Sincronización completada
