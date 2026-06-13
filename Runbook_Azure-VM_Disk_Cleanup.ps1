<#
.SYNOPSIS
    Automates disk space remediation on Linux VMs triggered by Azure Monitor alerts.
 
.DESCRIPTION
    This runbook is designed to be triggered by Azure Monitor Alerts via Action Groups.
    It uses the Automation Account's System-Assigned Managed Identity to authenticate
    and dynamically cleans up disk space on a Linux Virtual Machine via the Run Command API.
 
    PREREQUISITES:
    The Automation Account's System-Assigned Managed Identity MUST be granted
    the 'Virtual Machine Contributor' role on the target resource group or VMs (IAM).
 
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
    Write-Log "Initializing Disk Maintenance Runbook..." "INFO"
    
    $ResourceGroupName = $null
    $VMName = $null
 
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
        
        # Extract target database details from resource ID
        if ($null -ne $AlertContext -and $null -ne $AlertContext.data -and $null -ne $AlertContext.data.essentials) {
            $Essentials = $AlertContext.data.essentials
            
            if ($null -ne $Essentials.alertTargetIDs -and $Essentials.alertTargetIDs.Count -gt 0) {
                $ResourceID = $Essentials.alertTargetIDs[0]
                Write-Log "Target resource trace ID: $ResourceID" "INFO"
                
                # Match pattern for Azure Virtual Machines
                if ($ResourceID -match "resourceGroups/(?<rg>[^/]+)/providers/Microsoft.Compute/virtualMachines/(?<vm>[^/]+)") {
                    $ResourceGroupName = $Matches['rg']
                    $VMName            = $Matches['vm']
                    Write-Log "Target infrastructure identified -> RG: '$ResourceGroupName', VM: '$VMName'" "SUCCESS"
                }
            }
        }
    }
 
    # Validate final parameters
    if ([string]::IsNullOrEmpty($ResourceGroupName) -or [string]::IsNullOrEmpty($VMName)) {
        throw "Failed to extract target VM properties from the incoming alert payload."
    }
 
    $ResourceGroupName = $ResourceGroupName.Trim()
    $VMName            = $VMName.Trim()
 
    # Define embedded cleanup script
    $BashScript = @'
#!/bin/bash
echo "=== Starting automated disk cleanup ==="
if [ -x "$(command -v apt-get)" ]; then
    apt-get clean -y && apt-get autoremove -y
elif [ -x "$(command -v dnf)" ]; then
    dnf clean all -y
elif [ -x "$(command -v yum)" ]; then
    yum clean all -y
fi
if [ -x "$(command -v logrotate)" ]; then
    logrotate -f /etc/logrotate.conf 2>/dev/null
fi
if [ -x "$(command -v journalctl)" ]; then
    journalctl --vacuum-time=3d
fi
find /tmp -type f -atime +7 -delete 2>/dev/null
if [ -x "$(command -v docker)" ]; then
    docker system prune -af --volumes >/dev/null 2>&1
fi
echo "=== Disk cleanup completed successfully ==="
'@

    # Connect to Azure
    Write-Log "Authenticating to Azure via System-Assigned Managed Identity..." "INFO"
    Disable-AzContextAutosave -Scope Process | Out-Null
    $AzureContext = (Connect-AzAccount -Identity).context
    Write-Log "Authentication successful. Subscription: $($AzureContext.Subscription.Name)" "SUCCESS"
 
    # Run bash script inside the VM guest OS
    Write-Log "Dispatching execution command to Linux VM '$VMName' in Resource Group '$ResourceGroupName'..." "INFO"
    
    $CommandResult = Invoke-AzVMRunCommand -ResourceGroupName $ResourceGroupName `
                                           -VMName $VMName `
                                           -CommandId 'RunShellScript' `
                                           -ScriptString $BashScript
 
    # Log guest execution feedback
    Write-Log "Execution output from VM guest OS:" "INFO"
    foreach ($Value in $CommandResult.Value) {
        Write-Output $Value.Message
    }
 
    Write-Log "Disk maintenance process completed safely." "SUCCESS"
}
catch {
    Write-Log "An unexpected error occurred during execution: $($_.Exception.Message)" "ERROR"
    throw
}