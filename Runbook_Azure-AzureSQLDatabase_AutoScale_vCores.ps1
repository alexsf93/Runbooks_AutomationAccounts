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
    Date:        2026-05-20
    Version:     1.4 (Full English Translation - Production Ready)
#>
 
[CmdletBinding()]
param (
    [Parameter(Mandatory=$false, HelpMessage="Name of the Resource Group")]
    [string]$ResourceGroupName,
 
    [Parameter(Mandatory=$false, HelpMessage="Name of the SQL Server")]
    [string]$ServerName,
 
    [Parameter(Mandatory=$false, HelpMessage="Name of the SQL Database")]
    [string]$DatabaseName,
 
    [Parameter(Mandatory=$false, HelpMessage="Action to perform: ScaleUp or ScaleDown")]
    [string]$Action,

    [Parameter(Mandatory=$false, HelpMessage="Automatically populated by Azure Alerts")]
    [object]$WebhookData
)
 
$ErrorActionPreference = "Stop"
 
# ---------------------------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------------------------
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
       
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
   
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ssZ")
    $formattedMessage = "[$timestamp] [$Level] $Message"
   
    if ($Level -eq "ERROR") {
        Write-Error $formattedMessage
    } else {
        Write-Output $formattedMessage
    }
}
 
# ---------------------------------------------------------------------------
# Main Execution
# ---------------------------------------------------------------------------
try {
    Write-Log "Initializing AutoScale Runbook..." "INFO"
 
    # Validation: Check if triggered by an Azure Monitor Alert via Action Group
    if ($null -ne $WebhookData) {
        Write-Log "Runbook triggered by Azure Monitor Alert via Action Group. Parsing parameters..." "INFO"
        
        # 1. Attempt to extract from WebhookProperties (Static values from real alert)
        if ($null -ne $WebhookData.WebhookProperties -and $null -ne $WebhookData.WebhookProperties.Action) {
            $ResourceGroupName = $WebhookData.WebhookProperties.ResourceGroupName
            $ServerName        = $WebhookData.WebhookProperties.ServerName
            $DatabaseName      = $WebhookData.WebhookProperties.DatabaseName
            $Action            = $WebhookData.WebhookProperties.Action
            Write-Log "Parameters successfully extracted from WebhookProperties." "SUCCESS"
        } 
        # 2. If it is a test execution or properties are empty, attempt to parse RequestBody safely
        elseif ($null -ne $WebhookData.RequestBody) {
            Write-Log "WebhookProperties is empty or incomplete. Attempting to extract data from alert body..." "WARNING"
            $AlertContext = ConvertFrom-Json $WebhookData.RequestBody
            
            # Safe verification of Common Alert Schema structure
            if ($null -ne $AlertContext.data -and $null -ne $AlertContext.data.essentials -and $null -ne $AlertContext.data.essentials.alertTargetIds) {
                $ResourceID = $AlertContext.data.essentials.alertTargetIds[0]
                if ($ResourceID -match "resourceGroups/(?<rg>[^/]+)/providers/Microsoft.Sql/servers/(?<server>[^/]+)/databases/(?<db>[^/]+)") {
                    $ResourceGroupName = $Matches['rg']
                    $ServerName        = $Matches['server']
                    $DatabaseName      = $Matches['db']
                    Write-Log "Target resources dynamically extracted from Common Alert Schema." "INFO"
                }
            }
        }
    }

    # Specific control handle for Azure's "Test action group" button
    if ($null -ne $WebhookData -and [string]::IsNullOrEmpty($Action)) {
        Write-Log "ATTENTION: Test execution detected (Test Action Group). Forcing dummy parameters to validate workflow." "WARNING"
        $ResourceGroupName = if ([string]::IsNullOrEmpty($ResourceGroupName)) { "test-rg" } else { $ResourceGroupName }
        $ServerName        = if ([string]::IsNullOrEmpty($ServerName)) { "test-server" } else { $ServerName }
        $DatabaseName      = if ([string]::IsNullOrEmpty($DatabaseName)) { "test-db" } else { $DatabaseName }
        $Action            = "ScaleUp"
    }

    # Final safety check to ensure all required parameters are set
    if ([string]::IsNullOrEmpty($ResourceGroupName) -or [string]::IsNullOrEmpty($ServerName) -or [string]::IsNullOrEmpty($DatabaseName) -or [string]::IsNullOrEmpty($Action)) {
        throw "Missing required parameters. Current State -> RG: '$ResourceGroupName', Server: '$ServerName', DB: '$DatabaseName', Action: '$Action'"
    }

    # Trim spaces from input parameters
    $ResourceGroupName = $ResourceGroupName.Trim()
    $ServerName        = $ServerName.Trim()
    $DatabaseName      = $DatabaseName.Trim()
    $Action            = $Action.Trim()
 
    # FQDN Correction: If ServerName contains ".database.windows.net", strip it to short name
    if ($ServerName -like "*.database.windows.net*") {
        $ServerName = $ServerName.Split('.')[0]
        Write-Log "FQDN detected for SQL Server. Cleaned to short name: '$ServerName'" "INFO"
    }

    Write-Log "Target Resource: $DatabaseName on server $ServerName ($Action)" "INFO"
 
    # If a simulated test execution environment is detected, exit successfully without altering any real database
    if ($ServerName -eq "test-server" -or $ResourceGroupName -eq "test-rg") {
        Write-Log "Action Group Test completed successfully. Script workflow is correct." "SUCCESS"
        return
    }

    Write-Log "Authenticating to Azure via System-Assigned Managed Identity..." "INFO"
    Disable-AzContextAutosave -Scope Process | Out-Null
    $AzureContext = (Connect-AzAccount -Identity).context
    Write-Log "Authentication successful. Subscription: $($AzureContext.Subscription.Name)" "SUCCESS"
 
    Write-Log "Retrieving current database configuration..." "INFO"
    $db = Get-AzSqlDatabase -ResourceGroupName $ResourceGroupName -ServerName $ServerName -DatabaseName $DatabaseName
 
    # Robust capacity detection (handles version differences across Az.Sql module)
    $currentCapacity = $db.Capacity
   
    if ([string]::IsNullOrEmpty($currentCapacity)) {
        if ($null -ne $db.Sku -and $null -ne $db.Sku.Capacity) {
            $currentCapacity = $db.Sku.Capacity
        } elseif ($db.CurrentServiceObjectiveName -match "_(\d+)$") {
            $currentCapacity = [int]$matches[1]
        } else {
            Write-Log "Could not determine exact vCore capacity. Defaulting to 0 to force evaluation." "WARNING"
            $currentCapacity = 0
        }
    }
 
    Write-Log "Current Capacity: $currentCapacity vCores." "INFO"
 
    # Define Scale Limits & Increments
    $MaxVcores  = 6
    $MinVcores  = 4
    $StepVcores = 2
 
    # Execute Scaling Logic
    if ($Action -eq "ScaleUp") {
        if ($currentCapacity -lt $MaxVcores) {
            # Calculate next tier and ensure max limit is not exceeded
            $targetCapacity = $currentCapacity + $StepVcores
            if ($targetCapacity -gt $MaxVcores) { $targetCapacity = $MaxVcores }
 
            Write-Log "High CPU threshold reached. Scaling up by $StepVcores (Target: $targetCapacity vCores)..." "WARNING"
            Set-AzSqlDatabase -ResourceGroupName $ResourceGroupName -ServerName $ServerName -DatabaseName $DatabaseName -Capacity $targetCapacity | Out-Null
            Write-Log "Successfully scaled up to $targetCapacity vCores." "SUCCESS"
        }
        else {
            Write-Log "Database is already at $currentCapacity vCores (>= $MaxVcores). No scale-up required." "INFO"
        }
    }
    elseif ($Action -eq "ScaleDown") {
        if ($currentCapacity -gt $MinVcores) {
            # Calculate next tier and ensure min limit is not breached
            $targetCapacity = $currentCapacity - $StepVcores
            if ($targetCapacity -lt $MinVcores) { $targetCapacity = $MinVcores }
 
            Write-Log "CPU usage normalized. Scaling down by $StepVcores (Target: $targetCapacity vCores)..." "INFO"
            Set-AzSqlDatabase -ResourceGroupName $ResourceGroupName -ServerName $ServerName -DatabaseName $DatabaseName -Capacity $targetCapacity | Out-Null
            Write-Log "Successfully scaled down to $targetCapacity vCores." "SUCCESS"
        }
        else {
            Write-Log "Database is already at $currentCapacity vCores (<= $MinVcores). No scale-down required." "INFO"
        }
    }
    else {
        throw "Invalid action value: '$Action'. Must be 'ScaleUp' or 'ScaleDown'."
    }
 
    Write-Log "Runbook execution completed." "SUCCESS"
}
catch {
    Write-Log "An unexpected error occurred: $($_.Exception.Message)" "ERROR"
    throw
}
