# Find-DevicesFromSignInLog.PS1
# Use the Entra ID sign in logs to find what registered devices are being used.
# https://github.com/12Knocksinna/Office365itpros/blob/master/Find-DevicesFromSignInLog.PS1
# V1.0 7-Mar-2024

# Connect to the Graph is needed to fetch information about devices and users
Connect-MgGraph -NoWelcome -Scopes 	Directory.Read.All

# File downloaded from Entra admin center with non-interactive sign ins 
# needs to be edited to remove the incoming token type column first
$InputDataFile = "c:\users\tonyr\downloads\SignInData.csv"

Write-Host "Loading data"
[array]$Data =  Import-Csv $InputDataFile | Sort-Object {$_.'Date (UTC)' -as [datetime]} -Descending
# Retrieve devices found in sign in logs
[array]$FoundDevices = $Data | Sort-Object 'Device ID' -Unique 
$FoundDevices = $FoundDevices | Where-Object {($_.'Device ID' -ne "{PII Removed}")}  | Select-Object -ExpandProperty 'Device ID'
# Retrieve known devices
[array]$KnownDevices = Get-MgDevice -All

Write-Host "Generating report"
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Device in $FoundDevices) {
    If (!([string]::IsNullOrWhiteSpace($Device))) {
        $DeviceDetails = $KnownDevices | Where-Object {$_.DeviceId -eq $Device}
        $DataDetails = $Data | Where-Object {$_.'Device ID' -eq $Device} | Select-Object -First 1
        $RegisteredOwnerId = $null; $RegisteredOwner = $null
        $RegisteredOwnerId = Get-MgDeviceRegisteredOwner -DeviceId $DeviceDetails.Id | Select-Object -ExpandProperty Id
        $RegisteredOwner = Get-MgUser -UserId $RegisteredOwnerId -ErrorAction SilentlyContinue
        $SignInDate = Get-Date $DataDetails.'Date (UTC)' -format 'dd-MMM-yyyy HH:mm'
        $RegisteredDate = Get-Date $DeviceDetails.RegistrationDateTime -format 'dd-MMM-yyyy HH:mm'
        $ReportLine = [PSCustomObject][Ordered]@{
            SignIn              = $SignInDate
            Device              = $Device
            'Device name'       = $DeviceDetails.displayName
            Id                  = $DeviceDetails.Id
            OS                  = $DeviceDetails.OperatingSystem
            Version             = $DeviceDetails.OperatingSystemVersion
            Registered          = $RegisteredDate
            'Registered owner'  = $RegisteredOwner.UserPrincipalName
            'User agent'        = $DataDetails.'User agent'
            'User signing in'   = $DataDetails.User
            'User sign in UPN'  = $DataDetails.userName
            Resource            = $DataDetails.Resource
            ClientApp           = $DataDetails.'Client App'
            }
        $Report.Add($ReportLine)
        }
}

$Report | Out-GridView

[array]$UnusedDevices = $KnownDevices | Where-Object {$_.Id -notin $FoundDevices} | Sort-Object DisplayName

Write-Host ""
Write-Host "The following devices cannot be found in a sign-in log"
Write-Host "------------------------------------------------------"

$UnusedDevices | Format-Table Id, DisplayName, OperatingSystem, RegistrationDateTime

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
