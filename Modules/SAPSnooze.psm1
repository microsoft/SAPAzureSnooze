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
.\Start-ParallelJobs.ps1 ScriptBlockCopy "{param($ComputerName); get-service -ComputerName $ComputerName}" -ArgumentList @("Server1","Server2") -MaxRunspacePool 2

Runs the specified script block in parallel with two threads for the specified argument array
#>

Function Start-ParallelJobs{
    [cmdletbinding()]
    param(
    [parameter(mandatory=$true)]
    $ScriptBlockCopy,
    
    [parameter(mandatory=$true)]
    $ArgumentList,
    
    [int] $MaxRunSpacePool=20,
    
    [int] $MaxWaitTimeInSeconds = 3600,
    
    [int] $MonitoringIntervalInSeconds=1)

    #Write-Output "Creating background jobs..."
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
            #Write-Output "Completed% $CompletedJobCount/$JobCount"
        }
        
        Sleep $MonitoringIntervalInSeconds
        if(($Counter*$MonitoringIntervalInSeconds) -ge $MaxWaitTimeInSeconds){
            #Write-Warning "Waiting threshold $MaxWaitTimeInSeconds seconds exceeded for following:"
            #$Jobs | ?{$_.Result.Iscompleted -ne $true} | %{Write-Output $_.Pipe.Commands.Commands.parameters.Value}
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
.SYNOPSIS
Function to get SAP system status

.DESCRIPTION
.NOTES
Author:
DateUpdated: 03/25/2017
Version: 1.2

.PARAMETER SID
Script block that you want to execute in parallel

.PARAMETER MsgServerHost
SAP message server host

.PARAMETER MsgServerHTTPPort
SAP message server http port

.OUTPUTS
Returns the status of SAP system

.EXAMPLE
.\Get-SAPSnoozeSAPActiveInstances -SID T01 -MsgServerHost sapt01ms -MsgServerHTTPPort 8101

Gets status of SAP system T01 with message server sapt01ms and http port 8101
#>
Function Get-SAPSnoozeSAPActiveInstances{
    param(
    [cmdletbinding()]
    [Parameter(Mandatory=$true)]
    [string] $SID, 
    
    [Parameter(Mandatory=$true)]
    [string] $MsgServerHost,

    [Parameter(Mandatory=$true)]
    [string] $MsgServerHTTPPort
    )
	try{
        #Connect to the messageserver HTTP port to get the list of application servers
        $ErrorActionPreference = "stop"
        $webClient = New-Object System.Net.WebClient
	    $webClient.Headers.Add("user-agent", "PowerShell Script")
	    $msgurl = "http://${MsgServerHost}:${MsgServerHTTPPort}/msgserver/text/aslist"
        $output = $webClient.DownloadString($msgurl)
        return ($output -split "`n" | ?{$_ -notmatch "^version"} | select @{Name="InstanceName";Expression={($_ -split "`t")[0]}})
    }
    catch{
        return $null
    }
}

<#
.SYNOPSIS
Function to get status of VMs in a resource group

.DESCRIPTION
.NOTES
Author:
DateUpdated: 03/25/2017
Version: 1.2

.PARAMETER ResourceGroupName
Azure resource group name

.OUTPUTS
Overall status of VMs in the resource group. Even if one VM is offline, the overall status will be offline

.EXAMPLE
Get-SAPSnoozeVMStatus -ResourceGroupName SAP_SBX_T11_T11

Gets overall VM status of VMs in the resource group SAP_SBX_T11_T11
#>
Function Get-SAPSnoozeVMStatus{
    param(
    [cmdletbinding()]
    [Parameter(Mandatory=$true)]
    [string[]]$ResourceGroupName
    )
    try{
        $ResourceGroupName = $ResourceGroupName -split ","
        foreach($RG in $ResourceGroupName){
            if(Get-AzureRMVM -ResourceGroupName $RG -Status | ?{$_.PowerState -ne "VM Running"}){
                return "Offline"
            }
            else{
                return "Online"
            }
        }
    }
    catch{
        return $_.Exception.Message
    }
}