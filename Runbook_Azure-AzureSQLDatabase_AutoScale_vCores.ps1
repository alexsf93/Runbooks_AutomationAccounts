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

.NOTES
    Author:     Alejandro Suarez
    Date:       2026-04-25
    Version:    1.1
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true, HelpMessage="Name of the Resource Group")]
    [ValidateNotNullOrEmpty()]
    [string]$ResourceGroupName,

    [Parameter(Mandatory=$true, HelpMessage="Name of the SQL Server")]
    [ValidateNotNullOrEmpty()]
    [string]$ServerName,

    [Parameter(Mandatory=$true, HelpMessage="Name of the SQL Database")]
    [ValidateNotNullOrEmpty()]
    [string]$DatabaseName,

    [Parameter(Mandatory=$true, HelpMessage="Action to perform: ScaleUp or ScaleDown")]
    [ValidateSet("ScaleUp", "ScaleDown")]
    [string]$Action
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

    # Clean input parameters to prevent formatting issues from Azure Action Groups
    $ResourceGroupName = $ResourceGroupName.Trim()
    $ServerName        = $ServerName.Trim()
    $DatabaseName      = $DatabaseName.Trim()

    Write-Log "Target Resource: $DatabaseName on server $ServerName ($Action)" "INFO"

    Write-Log "Authenticating to Azure via System-Assigned Managed Identity..." "INFO"
    Disable-AzContextAutosave -Scope Process | Out-Null
    $AzureContext = (Connect-AzAccount -Identity).context
    Write-Log "Authentication successful. Subscription: $($AzureContext.Subscription.Name)" "SUCCESS"

    Write-Log "Retrieving current database configuration..." "INFO"
    $db = Get-AzSqlDatabase -ResourceGroupName $ResourceGroupName -ServerName $ServerName -DatabaseName $DatabaseName

    # Robust capacity detection (handles differences across Az.Sql module versions)
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
    $MinVcores  = 2
    $StepVcores = 2

    # Execute Scaling Logic
    if ($Action -eq "ScaleUp") {
        if ($currentCapacity -lt $MaxVcores) {
            # Calculate next tier and ensure we don't exceed the max limit
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
            # Calculate next tier and ensure we don't drop below the min limit
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

    Write-Log "Runbook execution completed." "SUCCESS"
}
catch {
    Write-Log "An unexpected error occurred: $($_.Exception.Message)" "ERROR"
    throw
}
