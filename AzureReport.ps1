<#
    .SYNOPSIS 
    Workplace Management Reporting

    .DESCRIPTION
    Install:   .\AzureReport.ps1 

    .ENVIRONMENT
    PowerShell 5.0

    .AUTHOR
    Niklas Rast
#>

#Settings
$ErrorActionPreference = "SilentlyContinue"
$Customer = ""
$AzureSMTPUser = ""
$AzureSMTPPassword = ConvertTo-SecureString "" -AsPlainText -Force 
$AzureCreds = New-Object System.Management.Automation.PSCredential -ArgumentList ($AzureSMTPUser, $AzureSMTPPassword)
$ReportRecipient = ''
$ReportGenerators  = '', ' or ', '' 

#Variable for last month
$LastMonth = (Get-Date -Format "MM").ToString() -1
if ($LastMonth -lt 10) {
    $LastMonth = "0" + $LastMonth
}
$LastMonthYear = (Get-Date -Format "yyyy")
$Month = $LastMonthYear.ToString() + "-" + $LastMonth.ToString()

$OutFile = "$PSSCRIPTROOT\$Month-$Customer-ModernWorkplaceReport.xlsx"

#Azure login token
$tenantId = ''
$appId = ''
$appSecret = ''
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

function AutopilotEvents {

    $SheetName = "Windows Autopilot (FOR WPS)" 
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
    $url = "https://graph.microsoft.com/v1.0/users"

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
        $AADUser = [PSCustomObject]@{
            Name = $item.givenName
            Surname = $item.surname
            Mail = $item.userPrincipalName
        }
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

function IntuneApplicationList {

    $SheetName = "Software Warenkorb" 
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

    $SheetName = "Software Paketierungen" 
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

    $SheetName = "Windows Updates (FOR WPS)" 
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
            Release = $item.releaseDate
            Update = $item.displayName
            Class = $item.qualityUpdateClassification
        }
    $MEMUpdate | Where-Object Release -match $Month | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append -AutoSize
    }
    Write-Host "Function WindowsUpdateForBusinessDeployments finished." -ForegroundColor Green
}

function IntuneAuditLogs {
    $SheetName = "Intune AuditLogs (FOR WPS)" 
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
 
function SendReportMail {

    $breakLines = "`r`n`r`n"
    $breakLine = "`r`n"

    ## Build parameters
    $mailParams = @{
        SmtpServer                 = 'smtp.office365.com'
        Port                       = '587'
        UseSSL                     = $true   
        Credential                 = $AzureCreds
        From                       = $AzureSMTPUser
        To                         = $ReportRecipient
        Subject                    = "Modern Workplace Reporting $Customer"
        Body                       = "Hallo," + $breakLine + "anbei der Report der Workplace Services des Kunden $Customer im Monat $Month." + $breakLines + "VG" + $breakLine + "Workplace Administration" + $breakLine + "Bei Fragen bitte an $ReportGenerators wenden."
        Attachment                 = $OutFile
        DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
    }

    ## Send the email
    Send-MailMessage @mailParams
    Write-Host "Function SendReportMail finished." -ForegroundColor Green
}

function SendReportMailGraph {
    $url = "https://graph.microsoft.com/v1.0/users/$AzureSMTPUser/sendMail"

    $FileName=(Get-Item -Path $OutFile).name
    $base64string = [Convert]::ToBase64String([IO.File]::ReadAllBytes($OutFile))

    # Set the WebRequest headers
    $headers = @{
        'Content-Type' = 'application/json'
        Accept = 'application/json'
        Authorization = "Bearer $aadToken"
    }

    $BodyJsonsend = @"
                    {
                        "message": {
                          "subject": "Modern Workplace Reporting $Customer",
                          "body": {
                            "contentType": "HTML",
                            "content": "Hallo, <br>
                            anbei der Report der Workplace Services des Kunden $Customer im Monat $Month. <br>
                            Bitte bei Fragen an $ReportGenerators wenden."
                            <br><br>
                            Viele Gruese <br>
                            Workplace Administration <br>
                            "
                          },
                          
                          "toRecipients": [
                            {
                              "emailAddress": {
                                "address": "$ReportRecipient"
                              }
                            }
                          ]
                          ,"attachments": [
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
    # Send mail
    Invoke-RestMethod -Method POST -Uri $url -Headers $headers -Body $BodyJsonsend
    Write-Host "Function SendReportMailGraph finished." -ForegroundColor Green
}

#Cleanup if run more than once a month
if ($true -eq (Test-Path ($OutFile))){
    Remove-Item -Path $OutFile -Force
}

#Create and send report
DefenderAlerts
AzurePrinter
AzureADDevices
AzureADUsers
AzureADGroups
IntuneApplicationList
IntuneCreatedPackages
AutopilotEvents
WindowsUpdateForBusinessDeployments
IntuneAuditLogs
SendReportMail
#SendReportMailGraph
