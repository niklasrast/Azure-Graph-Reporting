# Azure-Graph-Reporting

![GitHub repo size](https://img.shields.io/github/repo-size/niklasrast/Azure-Graph-Reporting)

![GitHub issues](https://img.shields.io/github/issues-raw/niklasrast/Azure-Graph-Reporting)

![GitHub last commit](https://img.shields.io/github/last-commit/niklasrast/Azure-Graph-Reporting)

This repo contains an powershell scripts to create an Excel report for Azure services from data exported through Microsoft Graph API.
My recommendation for you is to create an schedule task to auto-run this script every month automatically so that you have zero input for the creation of your reports.

## Report Workflow Picture:
![Alt text](https://github.com/niklasrast/Azure-Graph-Reporting/blob/main/img/workflow.png "Workflow details")

## Create report:
```powershell
C:\Windows\SysNative\WindowsPowershell\v1.0\PowerShell.exe -ExecutionPolicy Bypass -Command .\AzureReportWPS.ps1
```

## Report Customer settings:
```powershell
$Customer = "CUSTOMERNAME-HERE"
```

## Report MailSend settings:
```powershell
$ReportRecipient = 'someone@mydomain.tld'

function SendReportMailGraph {
    $MailTenantID = 'YOURTENANTID'
    $MailClientID = 'YOURAPPID'
    $MailClientsecret = 'YOURAPPSECRET'
    $MailSender = "someone@mydomain.tld"
```

## Report Feature Enablement:
Comment or uncomment the folowing lines to enable oder #disable report data
```powershell
DefenderAlerts
AzurePrinter
AzureADDevices
WindowsCloudPC
AzureADLicenses
AzureADUsers
AzureADGroups
IntuneApplicationList
IntuneCreatedPackages
AutopilotEvents
WindowsUpdateForBusinessDeployments
IntuneAuditLogs
```

## Report App registration settings:
Place the Azure AD App registration from the tenant where you want to grab the reporting data from
```powershell
$tenantId = 'TENANTID-HERE'
$appId = 'AZUREADAPPID-HERE'
$appSecret = 'AZUREADAPPSECRET-HERE'
```
![Alt text](https://github.com/niklasrast/Azure-Graph-Reporting/blob/main/img/azure-ad-app-registration-01.png "App registration details")

### Azure AD App permissions:
Create an Azure AD App registration and assign following rights:
![Alt text](https://github.com/niklasrast/Azure-Graph-Reporting/blob/main/img/azure-ad-app-permissions.png "App permissions details")

## Schedule task
Create a Basic schedule task with following configuration:
![Alt text](https://github.com/niklasrast/Azure-Graph-Reporting/blob/main/img/schedule-taks-01.png "Schedule task configuration")
![Alt text](https://github.com/niklasrast/Azure-Graph-Reporting/blob/main/img/schedule-taks-02.png "Schedule task configuration")
![Alt text](https://github.com/niklasrast/Azure-Graph-Reporting/blob/main/img/schedule-taks-03.png "Schedule task configuration")
![Alt text](https://github.com/niklasrast/Azure-Graph-Reporting/blob/main/img/schedule-taks-04.png "Schedule task configuration")
![Alt text](https://github.com/niklasrast/Azure-Graph-Reporting/blob/main/img/schedule-taks-05.png "Schedule task configuration")


## Requirements:
- PowerShell 5.0
- Azure AD
- Azure AD App registration

# Feature requests
If you have an idea for a new feature in this repo, send me an issue with the subject Feature request and write your suggestion in the text. I will then check the feature and implement it if necessary.

Created by @niklasrast 
