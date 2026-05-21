<#
.SYNOPSIS
    Automates the scaling of an Azure SQL Database based on performance metrics.
 
.DESCRIPTION
    This runbook is designed to be triggered by Azure Monitor Alerts via Action Groups.
    It uses the Automation Account's System-Assigned Managed Identity to authenticate
    and dynamically adjusts the vCores of an Azure SQL Database using step-scaling logic.
 
    PREREQUISITES:
    The Automation Account's System-Assigned Managed Identity MUST be granted
    the 'SQL DB Contributor' role on the target Azure SQL Server (IAM).
 
.PARAMETER ResourceGroupName
    The name of the resource group containing the Azure SQL Server.
 
.PARAMETER ServerName
    The name of the Azure SQL Server.
 
.PARAMETER DatabaseName
    The name of the Azure SQL Database to scale.
 
.PARAMETER Action
    The scaling action to perform. Valid values are 'ScaleUp' and 'ScaleDown'.

.PARAMETER WebhookData
    Automatically injected by Azure Monitor Action Groups when triggered by an alert.
 
.NOTES
    Author:      Alejandro Suarez
    Date:        2026-05-21
    Version:     1.5
#>
 
[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$ResourceGroupName,
 
    [Parameter(Mandatory=$false)]
    [string]$ServerName,
 
    [Parameter(Mandatory=$false)]
    [string]$DatabaseName,
 
    [Parameter(Mandatory=$false)]
    [string]$Action,

    [Parameter(Mandatory=$false)]
    [object]$WebhookData
)
 
$ErrorActionPreference = "Stop"
 
function Write-Log {
    param (
        [Parameter(Mandatory=$true)][string]$Message,
        [Parameter(Mandatory=$false)][ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")][string]$Level = "INFO"
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ssZ")
    Write-Output "[$timestamp] [$Level] $Message"
}
 
try {
    Write-Log "Initializing AutoScale Runbook..." "INFO"
 
    # Process inputs if triggered via webhook (Azure Monitor)
    if ($null -ne $WebhookData) {
        Write-Log "Runbook triggered by Azure Monitor Webhook. Analyzing data structure..." "INFO"
        
        $AlertContext = $null
        
        # Check if Common Alert Schema is directly on the root object
        if ($null -ne $WebhookData.schemaId -and $WebhookData.schemaId -eq "azureMonitorCommonAlertSchema") {
            $AlertContext = $WebhookData
            Write-Log "Common Alert Schema detected at Webhook root." "INFO"
        }
        # Fallback to RequestBody if serialized
        elseif ($null -ne $WebhookData.RequestBody) {
            $AlertContext = ConvertFrom-Json $WebhookData.RequestBody
            Write-Log "Common Alert Schema extracted from RequestBody." "INFO"
        }
        
        # Extract resource data and operation from alert payload
        if ($null -ne $AlertContext -and $null -ne $AlertContext.data -and $null -ne $AlertContext.data.essentials) {
            $Essentials = $AlertContext.data.essentials
            
            # Extract target database details from resource ID
            if ($null -ne $Essentials.alertTargetIDs -and $Essentials.alertTargetIDs.Count -gt 0) {
                $ResourceID = $Essentials.alertTargetIDs[0]
                Write-Log "Analyzing alert target resource ID: $ResourceID" "INFO"
                
                if ($ResourceID -match "resourceGroups/(?<rg>[^/]+)/providers/Microsoft.Sql/servers/(?<server>[^/]+)/databases/(?<db>[^/]+)") {
                    $ResourceGroupName = $Matches['rg']
                    $ServerName        = $Matches['server']
                    $DatabaseName      = $Matches['db']
                    Write-Log "Resources automatically detected -> RG: '$ResourceGroupName', Server: '$ServerName', DB: '$DatabaseName'" "SUCCESS"
                }
            }
            
            # Determine scale direction from alert rule name
            if ($null -ne $Essentials.alertRule) {
                $RuleName = $Essentials.alertRule
                if ($RuleName -match "scaleup") {
                    $Action = "ScaleUp"
                } elseif ($RuleName -match "scaledown") {
                    $Action = "ScaleDown"
                }
                Write-Log "Action determined via alert rule name ($RuleName) -> '$Action'" "SUCCESS"
            }
        }
    }

    # Validate final parameters
    if ([string]::IsNullOrEmpty($ResourceGroupName) -or [string]::IsNullOrEmpty($ServerName) -or [string]::IsNullOrEmpty($DatabaseName) -or [string]::IsNullOrEmpty($Action)) {
        throw "Missing required parameters. Evaluated State -> RG: '$ResourceGroupName', Server: '$ServerName', DB: '$DatabaseName', Action: '$Action'"
    }

    $ResourceGroupName = $ResourceGroupName.Trim()
    $ServerName        = $ServerName.Trim()
    $DatabaseName      = $DatabaseName.Trim()
    $Action            = $Action.Trim()
 
    # Strip FQDN suffix if present
    if ($ServerName -like "*.database.windows.net*") {
        $ServerName = $ServerName.Split('.')[0]
    }

    Write-Log "PROCESSING REAL RESOURCE CHANGE: DB '$DatabaseName' | Server '$ServerName' | Action: '$Action'" "INFO"
 
    # Connect to Azure
    Write-Log "Authenticating to Azure via System-Assigned Managed Identity..." "INFO"
    Disable-AzContextAutosave -Scope Process | Out-Null
    $AzureContext = (Connect-AzAccount -Identity).context
    Write-Log "Authentication successful. Subscription: $($AzureContext.Subscription.Name)" "SUCCESS"
 
    # Get current DB scale info
    Write-Log "Retrieving current database configuration..." "INFO"
    $db = Get-AzSqlDatabase -ResourceGroupName $ResourceGroupName -ServerName $ServerName -DatabaseName $DatabaseName
    $currentCapacity = $db.Capacity
   
    if ([string]::IsNullOrEmpty($currentCapacity)) {
        if ($null -ne $db.Sku -and $null -ne $db.Sku.Capacity) { $currentCapacity = $db.Sku.Capacity }
        elseif ($db.CurrentServiceObjectiveName -match "_(\d+)$") { $currentCapacity = [int]$matches[1] }
        else { $currentCapacity = 0 }
    }
 
    Write-Log "Current detected capacity: $currentCapacity vCores." "INFO"
 
    # Scaling thresholds
    $MaxVcores  = 6
    $MinVcores  = 4
    $StepVcores = 2
 
    # Execute scaling logic
    if ($Action -eq "ScaleUp") {
        if ($currentCapacity -lt $MaxVcores) {
            $targetCapacity = $currentCapacity + $StepVcores
            if ($targetCapacity -gt $MaxVcores) { $targetCapacity = $MaxVcores }
 
            Write-Log "High CPU threshold reached. Scaling up to $targetCapacity vCores..." "WARNING"
            Set-AzSqlDatabase -ResourceGroupName $ResourceGroupName -ServerName $ServerName -DatabaseName $DatabaseName -Capacity $targetCapacity | Out-Null
            Write-Log "Scale up completed successfully to $targetCapacity vCores." "SUCCESS"
        }
        else {
            Write-Log "Database is already at the maximum configured capacity ($currentCapacity vCores)." "INFO"
        }
    }
    elseif ($Action -eq "ScaleDown") {
        if ($currentCapacity -gt $MinVcores) {
            $targetCapacity = $currentCapacity - $StepVcores
            if ($targetCapacity -lt $MinVcores) { $targetCapacity = $MinVcores }
 
            Write-Log "CPU usage normalized. Scaling down to $targetCapacity vCores..." "INFO"
            Set-AzSqlDatabase -ResourceGroupName $ResourceGroupName -ServerName $ServerName -DatabaseName $DatabaseName -Capacity $targetCapacity | Out-Null
            Write-Log "Scale down completed successfully to $targetCapacity vCores." "SUCCESS"
        }
        else {
            Write-Log "Database is already at the minimum configured capacity ($currentCapacity vCores)." "INFO"
        }
    }
    else {
        throw "Invalid action value: '$Action'. Must be 'ScaleUp' or 'ScaleDown'."
    }
 
    Write-Log "Runbook execution completed successfully." "SUCCESS"
}
catch {
    Write-Log "An unexpected error occurred during execution: $($_.Exception.Message)" "ERROR"
    throw
}
