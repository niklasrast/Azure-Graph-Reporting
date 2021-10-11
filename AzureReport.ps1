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
$ReportGenerators  = ''
$TeamsURL = ""
$IgelServer = ""
$IgelUser = ""
$IgelPassword = (ConvertTo-SecureString "" -AsPlainText -Force) 

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
if ($IgelServer -ne "") {
    if (Get-Module -ListAvailable -Name PSIGEL) {
        Import-Module -Name PSIGEL
        Write-Host "Imported PSIGEL Module" -ForegroundColor Green

    } 
    else {
        Install-Module -Name PSIGEL
        Import-Module -Name PSIGEL
        Write-Host "Installed and Imported PSIGEL Module" -ForegroundColor Green
    }
}
if (Get-Module -ListAvailable -Name ImportExcel) {
    Import-Module -Name ImportExcel
    Write-Host "Imported ImportExcel Module" -ForegroundColor Green
} 
else {
    Install-Module -Name ImportExcel
    Import-Module -Name ImportExcel
    Write-Host "Installed and Imported ImportExcel Module" -ForegroundColor Green
}


function SendTeamsNotification {
    #Teams Notification
    $JSONBody = [PSCustomObject][Ordered]@{
        "@type"      = "MessageCard"
        "@context"   = "http://schema.org/extensions"
        "summary"    = "Modern Workplace Reporting Service"
        "themeColor" = '1683E0'
        "title"      = "Modern Workplace Reporting Service"
        "text"       = "Der monatliche Report ($Month) für den Kunden $Customer wurde erstellt und an die SDM versendet."
    }
    
    $TeamMessageBody = ConvertTo-Json $JSONBody -Depth 100
    Invoke-RestMethod -Uri $TeamsURL -Method Post -Body $TeamMessageBody -ContentType 'application/json' | Out-Null
    Write-Host "Function SendTeamsNotification finished." -ForegroundColor Green
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

    ($DefenderAlerts | ConvertFrom-Json) | Select-Object category, eventDateTime, description, severity | Where-Object eventDateTime -match $Month | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append 
    "--- NO MORE ENTRIES FOR $Month --- " | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append
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

    ($AUPDevices | ConvertFrom-Json) | Select-Object name, model, isShared | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append 
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

    ($AADDevices | ConvertFrom-Json) | Select-Object serialNumber, deviceName, operatingSystem, osVersion, manufacturer, model, emailAddress, lastSyncDateTime | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append
    Write-Host "Function AzureADDevices finished." -ForegroundColor Green
}

function IgelClientReport {

    if ($IgelServer -ne "") {
        $SheetName = "IGEL Clients"
        [pscredential]$IgelCreds = New-Object System.Management.Automation.PSCredential ($IgelUser, $IgelPassword)
        $WebSession = New-UMSAPICookie -Computername $IgelServer -Credential $IgelCreds
        Get-UMSDevice -Computername $IgelServer -WebSession $WebSession | Select-Object Name, Mac, LastIp | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append
        Remove-UMSAPICookie -Computername $IgelServer -WebSession $WebSession  
    } else {
        Write-Host "No IGEL UMS Server detected for $Customer - Skipped function"
    }
    Write-Host "Function IgelClientReport finished." -ForegroundColor Green
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

    ($AADUsers | ConvertFrom-Json) | Select-Object givenName, surname, userPrincipalName | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append
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

    ($AADGroups | ConvertFrom-Json) | Select-Object displayName, createdDateTime, isAssignableToRole, membershipRule | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append
    Write-Host "Function AzureADGroups finished." -ForegroundColor Green
}

function IntuneApplicationList {

    $SheetName = "$Customer Software Warenkorb" 
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

    ($MEMApplications | ConvertFrom-Json) | Select-Object displayName, createdDateTime, developer, owner, notes | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append
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

    ($MEMCreatedPackages | ConvertFrom-Json) | Select-Object displayName, createdDateTime, developer, owner, notes | Where-Object createdDateTime -match $Month | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append
    "--- NO MORE ENTRIES FOR $Month --- " | Export-Excel -Path $OutFile -MoveToEnd -WorksheetName $SheetName -Append
    Write-Host "Function IntuneCreatedPackages finished." -ForegroundColor Green
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

#Cleanup if run more than once a month
if ($true -eq (Test-Path ($OutFile))){
    Remove-Item -Path $OutFile -Force
}

#Create and send report
DefenderAlerts
AzurePrinter
AzureADDevices
IgelClientReport
AzureADUsers
AzureADGroups
IntuneApplicationList
IntuneCreatedPackages
SendReportMail
SendTeamsNotification