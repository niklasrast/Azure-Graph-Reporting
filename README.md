# Azure-Graph-Reporting

![GitHub repo size](https://img.shields.io/github/repo-size/niklasrast/Azure-Graph-Reporting)

![GitHub issues](https://img.shields.io/github/issues-raw/niklasrast/Azure-Graph-Reporting)

![GitHub last commit](https://img.shields.io/github/last-commit/niklasrast/Azure-Graph-Reporting)

This repo contains an powershell scripts to create an Excel report for Azure services from data exported through Microsoft Graph API.
My recommendation for you is to create an schedule task to auto-run this script every month automatically so that you have zero input for the creation of your reports.

## Create report:
```powershell
PowerShell.exe -ExecutionPolicy Bypass -Command .\AzureReport.ps1
```

## Report Customer settings:
```powershell
$Customer = "CUSTOMERNAME-HERE"
```

## Report SMTP Seder settings:
```powershell
$AzureSMTPUser = "M365SMTPADRESS-HERE"
$AzureSMTPPassword = ConvertTo-SecureString "M365SMTPPASSWORD-HERE" -AsPlainText -Force
$ReportRecipient = 'RECIPIENT-HERE'
$ReportGenerators  = 'REPLYTO-HERE'
```

## Report Teams settings:
```powershell
$TeamsURL = "TEAMSWEBHOOKURL-HERE"
```

## Report IGEL Clients settings:
```powershell
$IgelServer = "SERVERFQDN-HERE"
$IgelUser = "UMSADMIN-HERE"
$IgelPassword = (ConvertTo-SecureString "UMSADMINPASSWORD-HERE" -AsPlainText -Force) 
```

## Report App registration settings:
```powershell
$tenantId = 'TENANTID-HERE'
$appId = 'AZUREADAPPID-HERE'
$appSecret = 'AZUREADAPPSECRET-HERE'
```

## Azure AD App permissions:
![Alt text](https://github.com/niklasrast/Azure-Graph-Reporting/blob/main/azure-ad-app-permissions.png "App permissions details")

## Requirements:
- PowerShell 5.0
- Azure AD
- Azure AD App registration

# Feature requests
If you have an idea for a new feature in this repo, send me an issue with the subject Feature request and write your suggestion in the text. I will then check the feature and implement it if necessary.

Created by @niklasrast 