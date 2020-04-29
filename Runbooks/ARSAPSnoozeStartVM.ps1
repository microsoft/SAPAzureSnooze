param(
[Parameter(Mandatory=$false)]
[object] $WebhookData
)

<#
.SYNOPSIS
Function to run parallel threads using runspacepool

.DESCRIPTION


.NOTES
Author:
DateUpdated: 03/25/2017
Version: 1.2

.PARAMETER ScriptBlockCopy
Script block that you want to execute in parallel

.PARAMETER ArgumentList
Argument list array. Each element in the array would be passed as argument to each thread

.PARAMETER MaxRunSpacePool
Maximum number of parallel threads

.PARAMETER MaxWaitTimeInSeconds
Maximum waiting time is seconds for threads

.PARAMETER MonitoringIntervalInSeconds
Interal in seconds between monitoring of threads

.OUTPUTS
Returns the outputs from all threads in an array

.EXAMPLE
.\Start-ActivityInParallel.ps1 ScriptBlockCopy "{param($ComputerName); get-service -ComputerName $ComputerName}" -ArgumentList @("Server1","Server2") -MaxRunspacePool 2

Runs the specified script block in parallel with two threads for the specified argument array
#>

 Function Start-ActivityInParallel{
    [cmdletbinding()]
    param(
    [parameter(mandatory=$true)]
    $ScriptBlockCopy,
    [parameter(mandatory=$true)]
    $ArgumentList,
    $MaxRunSpacePool=20,
    $MaxWaitTimeInSeconds = 3600,
    $MonitoringIntervalInSeconds=1)

    Write-Output "Creating background jobs..."
    $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $MaxRunSpacePool)
    $RunspacePool.Open()
    $Jobs = @()
    $Counter = 1
    foreach ($Argument in $ArgumentList){
        $Job = [powershell]::Create().AddScript($ScriptBlockCopy).AddArgument($Argument)
        $Job.RunspacePool = $RunspacePool
        $Jobs += New-Object PSObject -Property @{Pipe = $Job;Result = $Job.BeginInvoke()}
        $Counter++
    }
    
    $Counter = 1
    $ActiveJobs = $Jobs
    $FirstTimePrintFlag = $true
    do{
        if($Jobs.Count -ne 0){
            $CompletedJobCount = ($Jobs | ?{$_.Result.IsCompleted -eq $true}).Count
            $JobCount = $Jobs.Count
            Write-Output "Completed% $CompletedJobCount/$JobCount"
        }
        
        Sleep $MonitoringIntervalInSeconds
        if(($Counter*$MonitoringIntervalInSeconds) -ge $MaxWaitTimeInSeconds){
            Write-Warning "Waiting threshold $MaxWaitTimeInSeconds seconds exceeded for following:"
            $Jobs | ?{$_.Result.Iscompleted -ne $true} | %{Write-Output $_.Pipe.Commands.Commands.parameters.Value}
            break
        }
        $Counter++
    }While ( $Jobs.Result.IsCompleted -contains $false)
    
    $Output = @()
    ForEach ($Job in $Jobs){
        if($Job.Result.IsCompleted){
            $Output += $Job.Pipe.EndInvoke($Job.Result)
        }
    }

    #Close runspacepool
    $RunspacePool.close()
    #collect garbage
    [gc]::Collect()

    return $Output
}

<#
SAPSnooze runbook to start system using PowerApps and web hook
#>
try{
    if($WebhookData){
        #-------------------------- Import Azure run as certificate on the hybrid worker server and login to Azure --------------------------
        # Generate the password used for this certificate
        Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue | Out-Null
        $Password = [System.Web.Security.Membership]::GeneratePassword(25, 10)

        # Stop on errors
        $ErrorActionPreference = 'stop'

        # Get the management certificate that will be used to make calls into Azure Service Management resources
        $RunAsCert = Get-AutomationCertificate -Name "AzureRunAsCertificate"

        # location to store temporary certificate in the Automation service host
        $CertPath = Join-Path $env:temp  "AzureRunAsCertificate.pfx"

        # Save the certificate
        $Cert = $RunAsCert.Export("pfx",$Password)
        Set-Content -Value $Cert -Path $CertPath -Force -Encoding Byte | Write-Verbose

        Write-Output ("Importing certificate into $env:computername local machine root store from " + $CertPath)
        $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        Import-PfxCertificate -FilePath $CertPath -CertStoreLocation Cert:\LocalMachine\My -Password $SecurePassword -Exportable | Write-Verbose

        # Test that authentication to Azure Resource Manager is working
        $RunAsConnection = Get-AutomationConnection -Name "AzureRunAsConnection"

        Connect-AzureRmAccount `
            -ServicePrincipal `
            -TenantId $RunAsConnection.TenantId `
            -ApplicationId $RunAsConnection.ApplicationId `
            -CertificateThumbprint $RunAsConnection.CertificateThumbprint | Write-Verbose

        Set-AzureRmContext -SubscriptionId $RunAsConnection.SubscriptionID | Write-Verbose

        #----------------------------- Main function starts here -----------------------------
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
        $Output = Start-ActivityInParallel -ScriptBlockCopy $ScriptBlockCopy -ArgumentList (Get-AzureRmVM -ResourceGroupName $ResourceGroupName) -MaxRunSpacePool 100 -MaxWaitTimeInSeconds 1200
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
