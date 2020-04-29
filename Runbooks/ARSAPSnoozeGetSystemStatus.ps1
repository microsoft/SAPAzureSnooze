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


#Fill in values for the following variables
$SharePointURL = "https://microsoft.sharepoint.com/teams/<SharePointSite>"
$SharePointListName = "SharePointListName"
$SharePointUserName = "Domain user with edit access on the SharePointlist"

<#
SAPSnooze runbook to get status of VMs for all SIDs in the SharePoint list and update the SharePoint list with the status: Online, Offline or Unknown
SharePointSDK module should be installed on the jump boxes
#>
try{
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
    #Populate sharepoint list
    $PsCred = Get-AutomationPSCredential -Name "REDMOND_${SharePointUserName}"
    $Password = $PsCred.GetNetworkCredential().Password
    $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
    $SharePointCred = New-Object System.Management.Automation.PSCredential ("${SharePointUserName}@microsoft.com", $secpasswd)
    $ListItems = Get-SPListItem -SiteUrl $SharePointURL -Credential $SharePointCred -IsSharePointOnlineSite $true -ListName $SharePointListName
    $SharePointList = @()
    $ListItems | %{
        $SharePointList += [PSCustomObject]@{
            ID=$_.ID;
            SID= $_.SID;
            Status=$_.Status;
            SAPStatus=$_.SAPStatus;
            VMStatus=$_.VMStatus;
            ResourceGroupName=$_.ResourceGroupName;
            MESSAGESERVERHOST=$_.MESSAGESERVERHOST;
            MESSAGESERVERPORT=$_.MESSAGESERVERPORT;
            GWSERV=$_.GWSERV;
            MESSAGESERVERHTTPPORT=$_.MESSAGESERVERHTTPPORT;
            User=$_.User
        }
    }

    Write-Output ($SharePointList| ft | out-string)
    
    #Save previous status to find out delta change
    $PreviousStatus = $SharePointList |  select SID, Status, SAPStatus, VMStatus

    #If MESSAGESERVERHOST and MESSAGESERVERHTTPPORT columns have values, then consider that as an SAP system. If not non-sap system
    $SharePointList | %{
        if($_.MESSAGESERVERHOST -ne $null -and $_.MESSAGESERVERHTTPPORT -ne $null){
            $_ | Add-Member -MemberType NoteProperty -Name SystemType -Value "SAP"
        }
        else{
            $_ | Add-Member -MemberType NoteProperty -Name SystemType -Value "NonSAP"
        }

        #Reset SAP status and VM status to null
        $_.SAPStatus = $null
        $_.VMStatus = $null
    }

    #Scriptblock to get System status in parallel
    $GetSAPStatusScriptBlock = {
        param ($item)
        try{
            $ErrorActionPreference = "Stop"
            Function Get-SAPSnoozeSAPActiveInstances{
                param(
                [Parameter(Mandatory=$true)]
                [string] $SID, 
                [Parameter(Mandatory=$true)]
                $MsgServerHost, 
                [Parameter(Mandatory=$true)]
                $MsgServerHTTPPort
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
            
            $SID = $item.SID
            $ActiveInstances = Get-SAPSnoozeSAPActiveInstances -SID $SID -MsgServerHost $item.MESSAGESERVERHOST -MsgServerHTTPPort $item.MESSAGESERVERHTTPPORT
            if($ActiveInstances){
                $Output = [PSCustomobject]@{SID=$SID;Status="Online";Details=$null}
            }
            else{
                $Output = [PSCustomobject]@{SID=$SID;Status="Offline";Details=$null}
            }
        }
        catch{
            $Output = [PSCustomobject]@{SID=$SID;Status="Error";Details=$_.Exception.Message}
        }
        return $Output
    }

    #Get status of all SAP systems in parallel
    $SAPStatus = Start-ActivityInParallel -ScriptBlockCopy $GetSAPStatusScriptBlock -ArgumentList ($SharePointList | ?{$_.SystemType -eq "SAP"}) -MaxRunSpacePool 100 -MaxWaitTimeInSeconds 60
    
    #For SAP systems that're timed out, set status as offline
    ($SharePointList | ?{$_.SystemType -eq "SAP"}) | ?{$_.SID -notin $SAPStatus.SID} | %{$SAPStatus += [PSCustomobject]@{SID=$_.SID;Status="Offline";Details="Message server connection timed out"}}
    
    #Update SAP status and VM Status in SharePointList object
    foreach($item in $SAPStatus){
        try{
            ($SharePointList | ?{$_.SID -eq $item.SID}).SAPStatus = $item.Status
            ($SharePointList | ?{$_.SID -eq $item.SID}).VMStatus = $item.Status
        }
        catch{
            break
        }
    }

    #Get list of resource groups for non-SAP systems and SAP systems with status offline, check VM status
    $RGsForVMStatusCheck = ($SharePointList | ?{$_.SystemType -eq "NonSAP" -or $_.SAPStatus -ne "Online"}).ResourceGroupName
    #Scriptblock to get System status in parallel
    $GetVMStatusScriptBlock = {
        param ($ResourceGroupName)
        try{
            $ErrorActionPreference = "Stop"
            Function Get-SAPSnoozeVMStatus{
                param(
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
            
            $Status = Get-SAPSnoozeVMStatus -ResourceGroupName $ResourceGroupName
            $Output = [PSCustomobject]@{ResourceGroupName=$ResourceGroupName;Status=$Status;Details=$null}
        }
        catch{
            $Output = [PSCustomobject]@{ResourceGroupName=$ResourceGroupName;Status="Error";Details=$_.Exception.Message}
        }
        return $Output
    }

    #Get status of VMs in parallel 
    $VMStatus = Start-ActivityInParallel -ScriptBlockCopy $GetVMStatusScriptBlock -ArgumentList $RGsForVMStatusCheck -MaxRunSpacePool 100 -MaxWaitTimeInSeconds 60
    #Update SAP status and VM Status in SharePointList object
    foreach($item in $VMStatus){
        ($SharePointList | ?{$_.ResourceGroupName -eq $item.ResourceGroupName}).VMStatus = $item.Status
        #If SAPStaus is null then 
        if(!(($SharePointList | ?{$_.ResourceGroupName -eq $item.ResourceGroupName}).SAPStatus)){
            ($SharePointList | ?{$_.ResourceGroupName -eq $item.ResourceGroupName}).SAPStatus = "NA"
        }
    }

    <#
    Update Status column:
    For SAP systems
        if SAPStatus is online then Online (Green)
        else
            If VM status is online then Starting (Yellow)
            else then Offline (Grey)
    Else
        If VMStatus is online then Online (Green)
        Else Offline (Grey)
    #>
    $SharePointList | %{
        if($_.SystemType -eq "SAP"){
            if($_.SAPStatus -eq "Online"){
                $_.Status = "Online"
            }
            else{
                if($_.VMStatus -eq "Online"){
                    $_.Status = "Starting"
                }
                else{
                    $_.Status = "Offline"
                }
            }
        }
        else{
            if($_.VMStatus -eq "Online"){
                $_.Status = "Online"
            }
            else{
                $_.Status = "Offline"
            }
        }
    }

    #Find list items for which status changed by comparing with the PreviousStatus variable
    $SIDsWithStatusChange = ((($SharePointList | select SID, Status, SAPStatus, VMStatus) + $PreviousStatus) | group SID | ?{($_.Group[0].SAPStatus -ne $_.Group[1].SAPStatus) -or ($_.Group[0].VMStatus -ne $_.Group[1].VMStatus) -or ($_.Group[0].Status -ne $_.Group[1].Status)}).Name
    $SIDsWithStatusChange = $SharePointList | ?{$_.SID -in $SIDsWithStatusChange}

    #Update systems for which status changed
    if($SIDsWithStatusChange){
        Write-Output "Following SIDs have status change. Updating SharePoint list"
        Write-Output ($SIDsWithStatusChange | ft -AutoSize -Wrap | Out-String)
        #Update SharePoint list
        #Scriptblock to update sharepoint list in parallel
        $SharepointListUpdateScriptBlock = {
            param ($SharePointListObject)
            try{
                $ErrorActionPreference = "Stop"

                #Update SharePoint list
                $SharePointURL = $SharePointListObject.SharePointURL
                $SharePointCred = $SharePointListObject.SharePointCred
                $SharePointListName = $SharePointListObject.SharePointListName
                $ListitemID = $SharePointListObject.ListitemID
                $Status = $SharePointListObject.Status
                $HashTable = @{}
                $Status.psobject.properties | Foreach { $HashTable[$_.Name] = $_.Value }
                $Null = Update-SPListItem -SiteUrl $SharePointURL -Credential $SharePointCred -IsSharePointOnlineSite $true -ListName $SharePointListName -ListItemID $ListItemID -ListFieldsValues $HashTable
                $Output = New-Object PSObject -Property @{ListItemID=$ListitemID;Status="OK"}
            }
            catch{
                $Output = New-Object PSObject -Property @{ListItemID=$ListitemID;Status=$_.Exception.Message}
            }
            return $Output
        }

        #Build hashtable for SharePoint list update
        $SharePointListObjects = @()
        foreach($Status in $SIDsWithStatusChange){
            $ListItemID = ($SharePointList | ?{$_.SID -eq $Status.SID}).ID
            $SharePointListObjects += [PSCustomobject]@{SharePointURL=$SharePointURL;SharePointCred=$SharePointCred;SharePointListName=$SharePointListName;ListItemID=$ListItemID;Status=($Status | select Status, SAPStatus, VMStatus)}
        }

        #Update SharePoinst list in parallel
        $SharePointListUpdateStatus= Start-ActivityInParallel -ScriptBlockCopy $SharepointListUpdateScriptBlock -ArgumentList $SharePointListObjects -MaxRunSpacePool 100 -MaxWaitTimeInSeconds 60
        
        #Display status
        Write-Output $SharePointListUpdateStatus
    }
    else{
        Write-Output "No SID with status change found"
    }
}
#Exception
catch{
    Write-Output $_.Exception
}
Finally{
    #collect garbage
    [gc]::Collect()
}
