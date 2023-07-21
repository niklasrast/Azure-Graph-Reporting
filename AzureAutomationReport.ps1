<#
    .SYNOPSIS 
    Modern Workplace Management Reporting

    .DESCRIPTION
    Install:   .\AzureAutomationReport.ps1 

    .ENVIRONMENT
    PowerShell 7.1

    .AUTHOR
    Niklas Rast
#>

#Runbook Parameters
Param (
  [Parameter (Mandatory= $true)]
  [String] $appId,
  [Parameter (Mandatory= $true)]
  [String] $appSecret,
  [Parameter (Mandatory= $true)]
  [String] $Customer,
  [Parameter (Mandatory= $true)]
  [String] $ReportRecipient,
  [Parameter (Mandatory= $true)]
  [String] $tenantId
)

#Settings
$ErrorActionPreference = "SilentlyContinue"

#Variable for last month
$LastMonth = (Get-Date -Format "MM").ToString() -1
if ($LastMonth -lt 10) {
    $LastMonth = "0" + $LastMonth
}
$LastMonthYear = (Get-Date -Format "yyyy")
$Month = $LastMonthYear.ToString() + "-" + $LastMonth.ToString()

$OutFile = "$PSSCRIPTROOT\$Month-$Customer-ModernWorkplaceReport.xlsx"

#Azure login token
$resourceAppIdUri = 'https://graph.microsoft.com'
$oAuthUri = "https://login.microsoftonline.com/$TenantId/oauth2/token"
$body = [Ordered] @{
    resource = "$resourceAppIdUri"
    client_id = "$appId"
    client_secret = "$appSecret"
    grant_type = 'client_credentials'
}
$response = Invoke-RestMethod -Method Post -Uri $oAuthUri -Body $body -ErrorAction Stop
$aadToken = $response.access_token

#Modules
if (Get-Module -ListAvailable -Name ImportExcel) {
    Import-Module -Name ImportExcel
    Write-Host "Imported ImportExcel Module" -ForegroundColor Green
} 
else {
    Install-Module -Name ImportExcel
    Import-Module -Name ImportExcel
    Write-Host "Installed and Imported ImportExcel Module" -ForegroundColor Green
}

function DefenderAlerts {
     
    $SheetName = "Defender Alerts" 
    $url = "https://graph.microsoft.com/v1.0/security/alerts"

    # Set the WebRequest headers
    $headers = @{
        'Content-Type' = 'application/json'
        Accept = 'application/json'
        Authorization = "Bearer $aadToken"
    }

    # Send the webrequest and get the results. 
    $response = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop
    #Extract the alerts from the results. 
    $DefenderAlerts = ($response | ConvertFrom-Json).value | ConvertTo-Json

    foreach ($item in ($DefenderAlerts | ConvertFrom-Json) ) {
        $DefenderAlert = [PSCustomObject]@{
            EventTime = $item.eventDateTime
            Category = $item.category
            Severity = $item.severity
            Description = $item.description
        }
    $DefenderAlert | Where-Object EventTime -match $Month  | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append -AutoSize
    }
    Write-Host "Function DefenderAlerts finished." -ForegroundColor Green
}

function AzurePrinter {

    $SheetName = "Azure Universal Print" 
    $url = "https://graph.microsoft.com/beta/print/printers"

    # Set the WebRequest headers
    $headers = @{
        'Content-Type' = 'application/json'
        Accept = 'application/json'
        Authorization = "Bearer $aadToken"
    }

    # Send the webrequest and get the results. 
    $response = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop
    #Extract the alerts from the results. 
    $AUPDevices = ($response | ConvertFrom-Json).value | ConvertTo-Json

    foreach ($item in ($AUPDevices | ConvertFrom-Json) ) {
        $AUPDevice = [PSCustomObject]@{
            Printername = $item.name
            Model = $item.model
            Active = $item.isShared
        }
    $AUPDevice | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append -AutoSize
    }
    Write-Host "Function AzurePrinter finished." -ForegroundColor Green
}

function AzureADDevices {

    $SheetName = "Azure AD Devices" 
    $url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"

    # Set the WebRequest headers
    $headers = @{
        'Content-Type' = 'application/json'
        Accept = 'application/json'
        Authorization = "Bearer $aadToken"
    }

    # Send the webrequest and get the results. 
    $response = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop
    #Extract the alerts from the results. 
    $AADDevices = ($response | ConvertFrom-Json).value | ConvertTo-Json

    foreach ($item in ($AADDevices | ConvertFrom-Json) ) {
        $AADDevice = [PSCustomObject]@{
            SerialNumber = $item.serialNumber
            DeviceName = $item.deviceName
            OperatingSystem = $item.operatingSystem
            Version = $item.osVersion
            Manufacturer = $item.manufacturer
            Model = $item.model
            PrimaryUser = $item.emailAddress
            LastIntuneSync = $item.lastSyncDateTime
        }
    $AADDevice | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append -AutoSize
    }
    Write-Host "Function AzureADDevices finished." -ForegroundColor Green
}

function WindowsCloudPC {

    $SheetName = "Windows 365" 
    $url = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs"

    # Set the WebRequest headers
    $headers = @{
        'Content-Type' = 'application/json'
        Accept = 'application/json'
        Authorization = "Bearer $aadToken"
    }

    # Send the webrequest and get the results. 
    $response = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop
    #Extract the SW List from the results. 
    $CloudPCList = ($response | ConvertFrom-Json).value | ConvertTo-Json

    foreach ($item in ($CloudPCList | ConvertFrom-Json) ) {
        $CPC = [PSCustomObject]@{
            Hostname = $item.managedDeviceName
            User = $item.userPrincipalName
            License = $item.servicePlanType
            LicenseType = $item.servicePlanName
            Image = $item.imageDisplayName
            State = $item.status
            LastUpdated = $item.lastModifiedDateTime
        }
    $CPC | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append -AutoSize
    }
    Write-Host "Function WindowsCloudPC finished." -ForegroundColor Green
}

function AutopilotEvents {

    $SheetName = "Windows Autopilot Logs" 
    $url = "https://graph.microsoft.com/beta/deviceManagement/autopilotEvents"

    # Set the WebRequest headers
    $headers = @{
        'Content-Type' = 'application/json'
        Accept = 'application/json'
        Authorization = "Bearer $aadToken"
    }

    # Send the webrequest and get the results. 
    $response = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop
    #Extract the alerts from the results. 
    $AutopilotEvents = ($response | ConvertFrom-Json).value | ConvertTo-Json

    foreach ($item in ($AutopilotEvents | ConvertFrom-Json) ) {
        $AutopilotEvent = [PSCustomObject]@{
            SerialNumber = $item.deviceSerialNumber
            DeviceName = $item.managedDeviceName
            Version = $item.osVersion
            DeploymentProfile = $item.windowsAutopilotDeploymentProfileDisplayName
            EnrollmentType = $item.enrollmentType
            EnrollmentState = $item.enrollmentState
            DeploymentStart = $item.deploymentStartDateTime
            DeploymentEnd = $item.deploymentEndDateTime
        }
    $AutopilotEvent | Where-Object DeploymentStart -match $Month | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append -AutoSize
    }
    Write-Host "Function AutopilotEvents finished." -ForegroundColor Green
}

function AzureADUsers {

    $SheetName = "Azure AD Users" 
    $url = "https://graph.microsoft.com/beta/users"

    # Set the WebRequest headers
    $headers = @{
        'Content-Type' = 'application/json'
        Accept = 'application/json'
        Authorization = "Bearer $aadToken"
    }

    # Send the webrequest and get the results. 
    $response = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop
    #Extract the alerts from the results. 
    $AADUsers = ($response | ConvertFrom-Json).value | ConvertTo-Json

    foreach ($item in ($AADUsers | ConvertFrom-Json) ) {
        $AssignedUserLicenses 
        foreach ($lic in $item.assignedLicenses) {
            $i = 0
            $AssignedUserLicenses = $AssignedUserLicenses + $lic + ", "
            $i = $i + 1
        }

        $AADUser = [PSCustomObject]@{
            Name = $item.givenName
            Surname = $item.surname
            Mail = $item.userPrincipalName
            OfficeLocation = $item.officeLocation
            Licenses = $($AssignedUserLicenses -replace '06ebc4ee-1bb5-47dd-8120-11324bc54e06','Microsoft 365 E5').Replace('05e9a617-0261-4cee-bb44-138d3ef5d965','Microsoft 365 E3').Replace('66b55226-6b4f-492c-910c-a3b7a3c9d993','Microsoft 365 F3').Replace('	4b585984-651b-448a-9e53-3b10f069cf7f','Office 365 F3').Replace('c5928f49-12ba-48f7-ada3-0d743a3601d5','Microsoft Visio (Plan 2)').Replace('4b244418-9658-4451-a2b8-b5e2b364e9bd','Microsoft Visio (Plan 1)').Replace('53818b1b-4a27-454b-8896-0dba576410e6','Microsoft Project (Plan 3)').Replace('f30db892-07e9-47e9-837c-80727f46fd3d','Microsoft Flow (Free)').Replace('440eaaa8-b3e0-484b-a8be-62870b9ba70a','Microsoft 365 Phone System Virtual User').Replace('MCOCAP','COMMON AREA PHONE').Replace('6070a4c8-34c6-4937-8dfb-39bbc6397a60','Microsoft Teams Rooms Standard').Replace('a403ebcc-fae0-4ca2-8c8c-7a907fd6c235','Microsoft Power BI Standard').Replace('c1d032e0-5619-4761-9b5c-75b6831e1711','Microsoft Power BI Premium').Replace('710779e8-3d4a-4c88-adb9-386c958d1fdf','Microsoft Teams Exploratory').Replace('d3b4fe1f-9992-4930-8acb-ca6ec609365e','Microsoft Skype for Business').Replace('8c4ce438-32a7-4ac5-91a6-e22ae08d9c8b','Rights Management Adhoc').Replace('5b631642-bd26-49fe-bd20-1daaa972ef80','Microsoft Power Apps for Developer').Replace('05e9a617-0261-4cee-bb44-138d3ef5d965','Microsoft 365 E3').Replace('66b55226-6b4f-492c-910c-a3b7a3c9d993','Microsoft 365 F3').Replace('	4b585984-651b-448a-9e53-3b10f069cf7f','Office 365 F3').Replace('c5928f49-12ba-48f7-ada3-0d743a3601d5','Microsoft Visio (Plan 2)').Replace('4b244418-9658-4451-a2b8-b5e2b364e9bd','Microsoft Visio (Plan 1)').Replace('53818b1b-4a27-454b-8896-0dba576410e6','Microsoft Project (Plan 3)').Replace('f30db892-07e9-47e9-837c-80727f46fd3d','Microsoft Flow (Free)').Replace('440eaaa8-b3e0-484b-a8be-62870b9ba70a','Microsoft 365 Phone System Virtual User').Replace('MCOCAP','COMMON AREA PHONE').Replace('6070a4c8-34c6-4937-8dfb-39bbc6397a60','Microsoft Teams Rooms Standard').Replace('a403ebcc-fae0-4ca2-8c8c-7a907fd6c235','Microsoft Power BI Standard').Replace('c1d032e0-5619-4761-9b5c-75b6831e1711','Microsoft Power BI Premium').Replace('710779e8-3d4a-4c88-adb9-386c958d1fdf','Microsoft Teams Exploratory').Replace('d3b4fe1f-9992-4930-8acb-ca6ec609365e','Microsoft Skype for Business').Replace('8c4ce438-32a7-4ac5-91a6-e22ae08d9c8b','Rights Management Adhoc').Replace('5b631642-bd26-49fe-bd20-1daaa972ef80','Microsoft Power Apps for Developer').Replace('71f21848-f89b-4aaa-a2dc-780c8e8aac5b','Windows 365 Business 2 vCPU 8 GB 128 GB').Replace('e2aebe6c-897d-480f-9d62-fff1381581f7','Windows 365 Enterprise 2 vCPU 8 GB 128 GB').Replace('@{disabledPlans=System.Object[]; skuId=','').Replace('}','')
            LastActivity = $item.signInSessionsValidFromDateTime
        }
        $AssignedUserLicenses = $null
    $AADUser | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append -AutoSize 
    }
    Write-Host "Function AzureADUsers finished." -ForegroundColor Green
}

function AzureADGroups {

    $SheetName = "Azure AD Groups" 
    $url = "https://graph.microsoft.com/v1.0/groups"

    # Set the WebRequest headers
    $headers = @{
        'Content-Type' = 'application/json'
        Accept = 'application/json'
        Authorization = "Bearer $aadToken"
    }

    # Send the webrequest and get the results. 
    $response = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop
    #Extract the alerts from the results. 
    $AADGroups = ($response | ConvertFrom-Json).value | ConvertTo-Json

    foreach ($item in ($AADGroups | ConvertFrom-Json) ) {
        $AADGroup = [PSCustomObject]@{
            CreatedDate = $item.createdDateTime
            GroupName = $item.displayName
            MembershipRule = $item.membershipRule
        }
    $AADGroup | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append -AutoSize
    }
    Write-Host "Function AzureADGroups finished." -ForegroundColor Green
}

function AzureADLicenses {

    $SheetName = "Azure AD Licenses" 
    $url = "https://graph.microsoft.com/v1.0/subscribedSkus"

    # Set the WebRequest headers
    $headers = @{
        'Content-Type' = 'application/json'
        Accept = 'application/json'
        Authorization = "Bearer $aadToken"
    }

    # Send the webrequest and get the results. 
    $response = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop
    #Extract the alerts from the results. 
    $AADLic = ($response | ConvertFrom-Json).value | ConvertTo-Json

    foreach ($item in ($AADLic | ConvertFrom-Json) ) {
        $AADLic = [PSCustomObject]@{
            #LicenseName = $item.skuPartNumber
            LicenseName = $($item.skuPartNumber -replace 'SPE_E5','Microsoft 365 E5').Replace('SPE_E3','Microsoft 365 E3').Replace('SPE_F1','Microsoft 365 F3').Replace('DESKLESSPACK','Office 365 F3').Replace('VISIOCLIENT','Microsoft Visio (Plan 2)').Replace('VISIOONLINE_PLAN1','Microsoft Visio (Plan 1)').Replace('PROJECTPROFESSIONAL','Microsoft Project (Plan 3)').Replace('FLOW_FREE','Microsoft Flow (Free)').Replace('PHONESYSTEM_VIRTUALUSER','Microsoft 365 Phone System Virtual User').Replace('MCOCAP','COMMON AREA PHONE').Replace('MEETING_ROOM','Microsoft Teams Rooms Standard').Replace('POWER_BI_STANDARD','Microsoft Power BI Standard').Replace('PBI_PREMIUM_PER_USER','Microsoft Power BI Premium').Replace('TEAMS_EXPLORATORY','Microsoft Teams Exploratory').Replace('MCOPSTN2','Microsoft Skype for Business').Replace('RIGHTSMANAGEMENT_ADHOC','Rights Management Adhoc').Replace('POWERAPPS_DEV','Microsoft Power Apps for Developer').Replace('CPC_B_2C_8RAM_128GB','Windows 365 Business 2 vCPU 8 GB 128 GB').Replace('CPC_E_2C_8GB_128GB','Windows 365 Enterprise 2 vCPU 8 GB 128 GB')
            Total = $item.prepaidUnits.enabled
            Assigned = $item.consumedUnits
        }
    $AADLic | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append -AutoSize
    }
    Write-Host "Function AADLicenses finished." -ForegroundColor Green
}

function IntuneApplicationList {

    $SheetName = "Software Inventory" 
    $url = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps"

    # Set the WebRequest headers
    $headers = @{
        'Content-Type' = 'application/json'
        Accept = 'application/json'
        Authorization = "Bearer $aadToken"
    }

    # Send the webrequest and get the results. 
    $response = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop
    #Extract the SW List from the results. 
    $MEMApplications = ($response | ConvertFrom-Json).value | ConvertTo-Json

    foreach ($item in ($MEMApplications | ConvertFrom-Json) ) {
        $MEMApplication = [PSCustomObject]@{
            PackageDate = $item.createdDateTime
            PackageName = $item.displayName
            Packager = $item.developer
            Order = $item.owner
            Class = $item.notes
            Owner = $item.publisher
        }
    $MEMApplication | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append -AutoSize
    }
    Write-Host "Function IntuneApplicationList finished." -ForegroundColor Green
}

function IntuneCreatedPackages {

    $SheetName = "Software Packaging last month" 
    $url = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps"

    # Set the WebRequest headers
    $headers = @{
        'Content-Type' = 'application/json'
        Accept = 'application/json'
        Authorization = "Bearer $aadToken"
    }

    # Send the webrequest and get the results. 
    $response = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop
    #Extract the SW List from the results. 
    $MEMCreatedPackages = ($response | ConvertFrom-Json).value | ConvertTo-Json

    foreach ($item in ($MEMCreatedPackages | ConvertFrom-Json) ) {
        $MEMPackage = [PSCustomObject]@{
            PackageDate = $item.createdDateTime
            PackageName = $item.displayName
            Packager = $item.developer
            Order = $item.owner
            Class = $item.notes
            Owner = $item.publisher
        }
    $MEMPackage | Where-Object PackageDate -match $Month | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append -AutoSize
    }
    Write-Host "Function IntuneCreatedPackages finished." -ForegroundColor Green
}

function WindowsUpdateForBusinessDeployments {

    $SheetName = "Windows Updates" 
    $url = "https://graph.microsoft.com/beta/admin/windows/updates/catalog/entries"

    # Set the WebRequest headers
    $headers = @{
        'Content-Type' = 'application/json'
        Accept = 'application/json'
        Authorization = "Bearer $aadToken"
    }

    # Send the webrequest and get the results. 
    $response = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop
    #Extract the SW List from the results. 
    $MEMCreatedUpdates = ($response | ConvertFrom-Json).value | ConvertTo-Json

    foreach ($item in ($MEMCreatedUpdates | ConvertFrom-Json) ) {
        $MEMUpdate = [PSCustomObject]@{
            Release = $item.releaseDateTime
            Update = $item.displayName
            Class = $item.qualityUpdateClassification
        }
    $MEMUpdate | Where-Object Release -match $Month | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append -AutoSize
    }
    Write-Host "Function WindowsUpdateForBusinessDeployments finished." -ForegroundColor Green
}

function IntuneAuditLogs {
    $SheetName = "Intune AuditLogs" 
    $url = "https://graph.microsoft.com/beta/deviceManagement/auditEvents"

    # Set the WebRequest headers
    $headers = @{
        'Content-Type' = 'application/json'
        Accept = 'application/json'
        Authorization = "Bearer $aadToken"
    }

    # Send the webrequest and get the results. 
    $response = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop
    #Extract the AuditLogs from the results. 
    $MEMAuditLogs = ($response | ConvertFrom-Json).value | ConvertTo-Json -Depth 3
    
    foreach ($item in ($MEMAuditLogs | ConvertFrom-Json) ) {
        $i = 0
        $AuditEvent = [PSCustomObject]@{
            EventTime = $item.activityDateTime
            EventType = $item.activityType
            Category = $item.category
            Actor = $item.actor.userPrincipalName
            Resource = $item.resources[$i].displayName
            ResourceType = $item.resources[$i].Type
        }
        $i = $i + 1
    $AuditEvent | Where-Object EventTime -match $Month | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append -AutoSize
    }
    Write-Host "Function IntuneAuditLogs finished." -ForegroundColor Green
}

function SendReportMailGraph {
    $MailTenantID = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
    $MailClientID = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
    $MailClientsecret = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
    $MailSender = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    $FileName=(Get-Item -Path $OutFile).name
    $base64string = [Convert]::ToBase64String([IO.File]::ReadAllBytes($OutFile))

    #Connect to GRAPH API
    $MailtokenBody = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $MailClientID
        Client_Secret = $MailClientsecret
    }
    $MailtokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$MailTenantID/oauth2/v2.0/token" -Method POST -Body $MailtokenBody
    $Mailheaders = @{
        "Authorization" = "Bearer $($MailtokenResponse.access_token)"
        "Content-type"  = "application/json"
    }

    #Send Mail    
    $URLsend = "https://graph.microsoft.com/v1.0/users/$MailSender/sendMail"
$BodyJsonsend = @"
                    {
                        "message": {
                          "subject": "Modern Workplace Reporting $Customer",
                          "body": {
                            "contentType": "HTML",
                            "content": "Hello,
                             <br><br>
                             attached a report for $Customer from $Month.
                             <br><br>
                             Best regards
                             <br>
                             Modern Workplace Services
                            "
                          },
                          "toRecipients": [
                            {
                              "emailAddress": {
                                "address": "$ReportRecipient"
                              }
                            }
                          ],
                          "ccRecipients": [
                          {
                            "emailAddress": {
                              "address": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
                            }
                          },
                          {
                            "emailAddress": {
                              "address": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
                            }
                          }
                          ],"attachments": [
                            {
                              "@odata.type": "#microsoft.graph.fileAttachment",
                              "name": "$FileName",
                              "contentType": "text/plain",
                              "contentBytes": "$base64string"
                            }
                          ]
                        },
                        "saveToSentItems": "true"
                      }
"@

    Invoke-RestMethod -Method POST -Uri $URLsend -Headers $Mailheaders -Body $BodyJsonsend
    Write-Host "Function SendReportMailGraph finished." -ForegroundColor Green
}

function MSTeamsAlert {

	$JSONBody = [PSCustomObject][Ordered]@{
    "@type" = "MessageCard"
    "@context" = "<http://schema.org/extensions>"
    "summary" = "Modern Workplace Reporting"
    "themeColor" = '0078D7'
    "title" = "Modern Workplace Reporting"
    "text" = "Monthly report for $Customer created and send."
	}

	$TeamMessageBody = ConvertTo-Json $JSONBody

	$parameters = @{
		"URI" = 'https://COMPANY.webhook.office.com/webhookb2/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/IncomingWebhook/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
		"Method" = 'POST'
		"Body" = $TeamMessageBody
		"ContentType" = 'application/json'
	}

	Invoke-RestMethod @parameters
	Write-Host "Function MSTeamsAlert finished." -ForegroundColor Green
}

#Cleanup if run more than once a month
#if ($true -eq (Test-Path ($OutFile))){
#    Remove-Item -Path $OutFile -Force
#}

#Debuging
Write-Host "----------------------------------------"
Write-Host "Customer: $Customer"
Write-Host "Recipient: $ReportRecipient"
Write-Host "Cutomer Tenant ID: $tenantId"
Write-Host "Customer App ID: $appId"
Write-Host "Customer App secret: $appSecret"
Write-Host "Directory: $PSSCRIPTROOT"
Write-Host "File: $OutFile"
Write-Host "----------------------------------------"

#Create and send report
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
SendReportMailGraph
MSTeamsAlert
