<#
.SYNOPSIS
Runbook does the following:
1) Pull system list from SharePoint list
2) If messageserver data is maintained then consider it as an SAP system and get SAP system status
3) If messageserver data is not maintained then consider it as a non-SAP system and get VM status
4) For SAP systems that'e offline, get VM status
5) Update SharePoint list for systems that have status change

The runbook requires the following modules:
SharePointSDK (Available on Azure Automation module library)
SAPSnooze (Available in SAPSnooze GitHub repo)

If the message server has a public IP, then you can use public IP as messageserver host in the SharePointlist and you can schedule the runbook to run as a Azure runbook
If the message server doesn't have a public IP, then you'll need to setup a Windows Hybrid Runbook Worker that can access the messageserver host and schedule the runbook to run on the hybrid worker.
Instructions on settuping Windows Hybrid Worker - https://docs.microsoft.com/en-us/azure/automation/automation-windows-hrw-install

.DESCRIPTION
.NOTES
Author:
DateUpdated: 03/25/2017
Version: 1.2

.PARAMETER SharePointURL
SharePoint URL name. E.g. https://microsoft.sharepoint.com/teams/SAPSnooze

.PARAMETER SharePointUserName
An account with edit access on the SharePointlist. E.g. sapsnooze@microsoft.com

.OUTPUTS

.EXAMPLE

#>

[cmdletbinding()]
param(
    [parameter(mandatory=$true)]
    [string] $SharePointURL,

    [parameter(mandatory=$true)]
    [string] $SharePointUserName
)

try{
    #Login to Azure using AzureRsAsConnection
    $connection = Get-AutomationConnection -Name AzureRunAsConnection
    Connect-AzureRmAccount `
        -ServicePrincipal `
        -Tenant $connection.TenantID `
        -ApplicationID $connection.ApplicationID `
        -CertificateThumbprint $connection.CertificateThumbprint

    #SharePoint list name should be "SAP System List" for PowerApps to work
    $SharePointListName = "SAP System List"

    #Import SAPSnooze module. This is required only if this is for a hybrid worker
    Import-Module c:\SAPSnooze\SAPSnooze.psm1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    
    #Populate sharepoint list
    $PsCred = Get-AutomationPSCredential -Name "SharePoint_${SharePointUserName}"
    $Password = $PsCred.GetNetworkCredential().Password
    $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
    $SharePointCred = New-Object System.Management.Automation.PSCredential ($SharePointUserName, $secpasswd)
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
            #Import SAPSnooze module. This is required only if this is for a hybrid worker
            Import-Module c:\SAPSnooze\SAPSnooze.psm1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
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
    $SAPStatus = Start-ParallelJobs -ScriptBlockCopy $GetSAPStatusScriptBlock -ArgumentList ($SharePointList | ?{$_.SystemType -eq "SAP"}) -MaxRunSpacePool 100 -MaxWaitTimeInSeconds 60
    
    #For SAP systems that're timed out, set status as offline
    ($SharePointList | ?{$_.SystemType -eq "SAP"}) | ?{$_.SID -notin $SAPStatus.SID} | %{$SAPStatus += [PSCustomobject]@{SID=$_.SID;Status="Offline";Details="Message server connection timed out"}}
    
    #Update SAP status and VM Status in SharePointList object
    foreach($item in $SAPStatus){
        ($SharePointList | ?{$_.SID -eq $item.SID}).SAPStatus = $item.Status
        ($SharePointList | ?{$_.SID -eq $item.SID}).VMStatus = $item.Status
    }

    #Get list of resource groups for non-SAP systems and SAP systems with status offline, check VM status
    $RGsForVMStatusCheck = ($SharePointList | ?{$_.SystemType -eq "NonSAP" -or $_.SAPStatus -ne "Online"}).ResourceGroupName
    #Scriptblock to get System status in parallel
    $GetVMStatusScriptBlock = {
        param ($ResourceGroupName)
        try{
            $ErrorActionPreference = "Stop"
            $Status = Get-SAPSnoozeVMStatus -ResourceGroupName $ResourceGroupName
            $Output = [PSCustomobject]@{ResourceGroupName=$ResourceGroupName;Status=$Status;Details=$null}
        }
        catch{
            $Output = [PSCustomobject]@{ResourceGroupName=$ResourceGroupName;Status="Error";Details=$_.Exception.Message}
        }
        return $Output
    }

    #Get status of VMs in parallel 
    $VMStatus = Start-ParallelJobs -ScriptBlockCopy $GetVMStatusScriptBlock -ArgumentList $RGsForVMStatusCheck -MaxRunSpacePool 100 -MaxWaitTimeInSeconds 60
    #Update SAP status and VM Status in SharePointList object
    foreach($item in $VMStatus){
        ($SharePointList | ?{$_.ResourceGroupName -eq $item.ResourceGroupName}).VMStatus = $item.Status
        #If SAPStaus is null then setup SAPStatus as Not Applicable (NA)
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
    $SIDsWithStatusChange = (([PSCustomObject[]]($SharePointList | select SID, Status, SAPStatus, VMStatus) + $PreviousStatus) | group SID | ?{($_.Group[0].SAPStatus -ne $_.Group[1].SAPStatus) -or ($_.Group[0].VMStatus -ne $_.Group[1].VMStatus) -or ($_.Group[0].Status -ne $_.Group[1].Status)}).Name
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
        $SharePointListUpdateStatus= Start-ParallelJobs -ScriptBlockCopy $SharepointListUpdateScriptBlock -ArgumentList $SharePointListObjects -MaxRunSpacePool 100 -MaxWaitTimeInSeconds 60
        
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
