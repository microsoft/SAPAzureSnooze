<p align="left">
<img width="200" height="40" src="MD%20image/1.png"> 
</p>  
  

<div align="center">

# SAPSnooze PowerApp
</div>

### **Application Overview:**
#### **Prerequisites:**
1.	Azure automation account
2.	Windows Hybrid worker (if SAP message server doesn't have a public IP)
3.	Microsoft PowerApps 
4.	SharePoint List

#### **What’re the features of the SAPSnooze Application?** 
You can use SAPSnooze PowerApps application to get status of the SAP system and VMs. 
You can also use it start systems that are stopped in Azure.

#### **SharePoint List**
Create a SharePoint List "SAP System List" with the following properties
IMPORTANT: Please use the same list name and column names as they're referenced in multiple places in the PowerApps application)
#### **SharePoint List properties:** 
A SharePoint list with the following columns needs to be created for maintain system information for PowerApps application. 
1.	Title (Type: Single line of text) – Title of the application. PowerApps application landing page will group the systems according to their application. 
2.	SID (Type: Single line of text) – SAP System ID 
3.	ResourceGroupName (Type: Single line of text) – Azure resource group that hosts the virtual machines for the SAP system
4.	Status (Type: Single line of text) – Status of the SAP system. Allowed values are (Online, Offline, Unknown, Starting) 
5.	SAPStatus (Type: Single line of text) – Status of the SAP system 
6.	VMStatus (Type: Single line of text) – Status of virtual machines in Azure
7.	User (Type: Single line of text) – Email address of the user who started the system 
8.	MESSAGESERVERHOST (Type: Single line of text) – Message server host name of the SAP system
10.	MESSAGESERVERHTTPPORT (Type: Single line of text) – Message server HTTP port of the SAP system (81XX) 

Create the system list by filling in the above properties

#### **SharePoint List permissions:** 
The system account used in Azure Automation should have edit access on the SharePoint list.
All users of the SAPSnooze application should have edit access on the SharePoint list. This can be granted either individually or by an active directory security group.

#### **Setting up Azure runbooks:** 
Following runbooks need to be imported from \Runbooks in your Azure automation account
1.	Get-SAPSnoozeSystemStatus.ps1 – Runbook to get status of SAP systems. This should be scheduled to run every 15 mins. You will need to create 4 schedules with 1 hour frequency to start at 15 minutes intervals. 
2.	Start-SAPSnoozeSystem.ps1 – Runbook to start VMs. Once this runbook is created, create a webhook and capture the URL. (The webhook will be of the following format
https://<AzureWebHookHost>/webhooks?token=<token>. Capture the AzureWebHookHost name and the token. Please note if the token has %3d at the end, then replace that with "=")

By default the credential for the user with edit access to the SharePointList is referenced as below in the runbook. If you name it differently, then please update the runbook as well. 
SharePoint_<UserName> (E.g. SharePoint_user@something.com)

If the message server has a public IP:
1.	You can use public IP as messageserver host in the SharePointlist. If you're using Public IP, please make sure that the message server http port is allowed in NSG. 
2.	Import SharePointSDK module in your Azure Automation Account from "Modules gallery"
3.	Import SAPSnooze module from \Modules in your Azure Automation Account
4.	You can schedule the runbook to run as an Azure runbook. The runbook will require two input parameters:1) SharePointURL (SharePoint URL that has the 'SAP System List'. E.g. https://microsoft.sharepoint.com/teams/SomeSite) 2) SharePointUserName (An account with edit access on the SharePointlist. E.g. user@something.com)

If the message server doesn't have a public IP
1.	Please setup a Windows Hybrid Runbook Worker that can access the messageserver host by following the instructions here. https://docs.microsoft.com/en-us/azure/automation/automation-windows-hrw-install
2.	Install SharePointSDK module on all your Hybrid worker (Install-Module SharePointSDK -Force)
3.	Import SAPSnooze module from the repo in your Azure Automation Account
4.	Copy SAPSnooze module from the repo to the directory C:\SAPSnooze on all hybrid workers
5.	Please schedule the runbook to run on the hybrid worker. The runbook will require two input parameters:1) SharePointURL (SharePoint URL that has the 'SAP System List'. E.g. https://microsoft.sharepoint.com/teams/SomeSite) 2) SharePointUserName (An account with edit access on the SharePointlist. E.g. user@something.com)

The runbook Start-SAPSnoozerunbook requires an Azure automation Run As account which has start VM access the resource groups in scope.
Please find instructions for setting up Azure automation Run As account here. https://docs.microsoft.com/en-us/azure/automation/manage-runas-account

### **Creating PowerApps application** 
#### **Create PowerApps Connectors:** 
1.	Go to PowerApps Studio:[https://preview.create.powerapps.com/studio/#](https://preview.create.powerapps.com/studio/#)
2.	Create a new Connection for SharePoint. 
3.	Use the credentials of a user who has read and write permissions on the SharePoint List that you’ve created in the above step.
4.	Open the file /PowerApps/Start-SAP-Systems.swagger.json and update value AzureWebHookHost with the host name in the webhook URL
5.	Create a new custom connector “**Start SAP Systems**” using the file /PowerApps/Start-SAP-Systems.swagger.json
6.	Create the connection "**Start SAP Systems**" using the newly created custom connector 
7.	Once the connectors are created, then import the PowerApps package from \PowerApps\SAPSnooze.zip by using the "**Import canvas app**" button to create the application
8.	Click on the Action button to select the name of application and connectors and then click on import to finish importing the application.  

#### **Update configuration:** 
Once created, edit the application in PowerApps studio
1) Go to “**Start**” button function and update the web hook token that you’ve created in the step to setup Azure runbooks (only the token is required not the whole URL. Also as mentioned in the setup above, if the token has a %3d at the end, then please replace that with "=")
2) Go to “**View --> Data Sources**”. Click on "Add data source", select the SharePoint connection that you created above, enter the SharePoint site URL and choose the SharePoint list "SAP System List"

#### **Permission:** 
Permission to PowerApps can be given to individual users or using active directory security groups. Please note people who’ve been access to PowerApps application should also be given edit access to the SharePoint list that’s created in the previous step.

> **Note**:This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact **opencode@microsoft.com** with any additional questions or comments.

