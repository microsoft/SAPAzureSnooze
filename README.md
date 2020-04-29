This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/)
or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.


------------ SAPSnooze PowerApps Application Overview ------------

------------ Prerequisites ------------
Azure automation account
Microsoft PowerApps subscription
SharePoint List
SQL Azure Database

------------ What is snoozing an SAP system ------------
Snoozing an SAP system means the Azure VMs hosting that system are stopped and deallocated. Stopping VMs saves compute cost.

------------ What’re the features of the SAPSnooze Application ------------
SAPSnooze PowerApps application provides the status of the SAP system and the VMs.
It also provides a start button using which systems that are snoozed can be brought back online 


------------ Setting up Azure runbooks, SQL Azure Tables, Hybrid worker group servers and SAP User account ------------

Following is the list of runbooks that’re needed:
1. ARSAPSnoozeGetSystemStatus.ps1 – Runbook to get status of SAP systems. This should be scheduled to run every 15 mins. 
2. ARSAPSnoozeStartVM.ps1 – Runbook to start VMs. Once this runbook is created, create a webhook and capture the URL. 

Hybrid worker group servers need to have the following modules installed:
• SharePointSDK
Use command Install-Module SharePointSDK -Force to install
 
------------ Automation run as account and permissions ------------
An automation run as account needs to be created that has permission to start the VMs 

------------ Telemetry/Usage statistics ------------
A SQL Azure table with the following structure needs to be created to capture usage statics of the PowerApps application. PowerApps would log the details each time a user logs in or start/stop/extend a system. 
CREATE TABLE [dbo].[UserTelemetry](
	[Timestamp] [datetime] NOT NULL,
	[Name] [varchar](256) NOT NULL,
	[Email] [varchar](256) NOT NULL,
	[System_SID] [varchar](10) NULL,
	[Action] [varchar](256) NULL,
 CONSTRAINT [PK_UserTelemetry] PRIMARY KEY CLUSTERED 
(
	[Timestamp] ASC,
	[Name] ASC,
	[Email] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

TimeStamp – Timestamp of the entry
Name – Name of the user
Email – Email address of the user
Action – Action by the user (Login/Start/Stop/Extend)

------------ SharePoint List ------------
SharePoint List properties
A SharePoint list with the following columns needs to be created for maintain system information for PowerApps application. 
•	Title – Title of the application. PowerApps application landing page will group the systems according to their application. 
•	SID – SAP System ID
•	ResourceGroupName – Azure resource group that hosts the virtual machines for the SAP system 
•	Status – Status of the SAP system. Allowed values are (Online, Offline, Unknown, Starting)
•	SAPStatus – Status of the SAP system
•	VMStatus – Status of virtual machines in Azure
•	User – Email address of the user who started the system
•	MESSAGESERVERHOST – Message server host name of the SAP system
•	MESSAGESERVERPORT – Message server port of the SAP system (36XX)
•	MESSAGESERVERHTTPPORT – Message server HTTP port of the SAP system (81XX)
•	GWSERV – Gateway port of the SAP system (33XX)

------------ How to onboard a new system on the application ------------
Fill the following details for the SAP system in the SharePoint list
Title, SID, ResourceGroupName, MESSAGESERVERHOST, MESSAGESERVERPORT, MESSAGESERVERHTTPPORT, GWSERV

------------ SharePoint List permissions ------------
Users who will be accessing the PowerApps application need edit access on the SharePoint list. This can be granted by an active directory security group. 
The active directory user that’s used by the Azure runbooks should have edit access on the SharePoint list.

------------ Creating PowerApps application ------------
------ Create PowerApps Connectors ------
Go to PowerApps Studio: https://preview.create.powerapps.com/studio/#
Create a new Connection for SharePoint. Use the credentials of a user who has read and write permissions on the SharePoint List that you’ve created in the above step
Create a new custom connector “Start SAP Systems” using the json file under /PowerApps/Start-SAP-Systems.swagger.json
Create the connection "Start SAP Systems" using the newly created custom connector
Create a new SQL Server Connection for collecting usage telemetry
Once the connectors are created, then import the PowerApps package from \PowerApps\SAPSnooze.zip by using the "Import canvas app" button to create the application:
Click on the Action button to select the name of application and connectors and then click on import to finish importing the application. 

------------ Update configuration ------------
Once created, edit the application in PowerApps studio, go to “Start” button function and update the web hook token that you’ve created in the step to setup Azure runbooks. 
Similarly, update telemetry database table also, in the “Start” button function and in the App lan

------------ Permission ------------
Permission to PowerApps can be given to individual users or using active directory security groups. Please note people who’ve been access to PowerApps application should also be given edit access to the SharePoint list that’s created in the previous step. 

