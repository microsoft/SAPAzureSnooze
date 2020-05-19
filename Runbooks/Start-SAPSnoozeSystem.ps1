<#
.SYNOPSIS
Runbook does the following:
1) Parse the resourcegroupname from webhookdata
2) Start all VMs in the resourcegroup in parallel

The runbook requires an Azure automation Run As account which has start VM access on this resource group.
Please find instructions for setting up Azure automation Run As account here. https://docs.microsoft.com/en-us/azure/automation/manage-runas-account

.DESCRIPTION
.NOTES
Author:
DateUpdated: 03/25/2017
Version: 1.2

.PARAMETER WebhookData
Webhookdata passed to the webhook

.OUTPUTS

.EXAMPLE
#>

param(
[Parameter(Mandatory=$false)]
[object] $WebhookData
)

try{
    if($WebhookData){
        #Login to Azure using AzureRsAsConnection
        $connection = Get-AutomationConnection -Name AzureRunAsConnection
        Connect-AzureRmAccount `
            -ServicePrincipal `
            -Tenant $connection.TenantID `
            -ApplicationID $connection.ApplicationID `
            -CertificateThumbprint $connection.CertificateThumbprint

        #Parse JSON request body from WebhookData
        $WebhookBody = (ConvertFrom-Json -InputObject $WebhookData.RequestBody)
        $ResourceGroupName = $WebhookBody.SID
        
        Write-Output $ResourceGroupName
        
        #Script block to start a VM
        $ScriptBlockCopy = {
            param ($VM)
            try{
                $ErrorActionPreference = "Stop"
                $null = $VM | Start-AzureRmVM
                $Output = [PSCustomObject]@{Name=$VM.Name;Status="OK"}
            }
            catch{
                $Output = [PSCustomObject]@{Name=$VM.Name;Status=$_.Exception.Message}
            }
        }

        Write-Output ((Get-AzureRmVM -ResourceGroupName $ResourceGroupName) | out-string)
        #Start all VMs in the resource group in parallel
        $Output = Start-ParallelJobs -ScriptBlockCopy $ScriptBlockCopy -ArgumentList (Get-AzureRmVM -ResourceGroupName $ResourceGroupName) -MaxRunSpacePool 100 -MaxWaitTimeInSeconds 1200
        Write-Output ($Output | ft -AutoSize -Wrap | Out-String)
    }
}
catch{
    $_.Exception
}
Finally{
    #collect garbage
    [gc]::Collect()
}
