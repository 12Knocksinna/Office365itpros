
# Report-EntraRegisteredDevices.PS1
# A demo script to show how to generate a report about Entra registered devices using the 
# Microsoft Graph PowerShell SDK. The script is based on the example in the article
# described in the article https://office365itpros.com/2023/01/31/entra-id-registered-devices-os/
# This script is https://github.com/12Knocksinna/Office365itpros/blob/master/Report-EntraRegisteredDevices.PS1
# V1.0 24-June-2024

Connect-MgGraph -Scope User.Read.All, Directory.Read.All

Write-Host "Finding registered devices"
[array]$Devices = Get-MgDevice -All -PageSize 999
If (!($Devices)) { 
  Write-Host "No registered devices found - exiting" ; break 
}

Write-Host ("Processing details for {0} devices" -f $Devices.count)
$Report = [System.Collections.Generic.List[Object]]::new() 
$i = 0
ForEach ($Device in $Devices) {
  $i++
  Write-Host ("Reporting device {0} ({1}/{2})" -f $Device.DisplayName, $i, $Devices.count)
  $DeviceOwner = $Null
  Try {
    [array]$OwnerIds = Get-MgDeviceRegisteredOwner -DeviceId $Device.Id
    $DeviceOwner = Get-MgUser -UserId $OwnerIds[0].Id `
     -Property Id, displayName, Department, OfficeLocation, City, Country, UserPrincipalName
    }
  Catch { 
    Write-Host ("Error fetching owners for {0}" -f $Device.DisplayName)
  }

  $ReportLine = [PSCustomObject][Ordered]@{
   Device              = $Device.DisplayName
   Id                  = $Device.Id
  "Device last signed in"= $Device.ApproximateLastSignInDateTime
  "Days since sign in" = (New-TimeSpan($Device.ApproximateLastSignInDateTime)).Days
   Owner               = $DeviceOwner.DisplayName
   OwnerUPN            = $DeviceOwner.UserPrincipalName
   Department          = $DeviceOwner.Department
   Office              = $DeviceOwner.OfficeLocation
   City                = $DeviceOwner.City
   Country             = $DeviceOwner.Country
   "Operating System"  = $Device.OperatingSystem
   "O/S Version"       = $Device.OperatingSystemVersion
   Registered          = $Device.RegistrationDateTime
   "Account Enabled"   = $Device.AccountEnabled
   DeviceId            = $Device.DeviceId
   TrustType           = $Device.TrustType 
  }
  $Report.Add($ReportLine)

} #End Foreach Device

# Sort in order of last signed in date
$Report = $Report | Sort-Object {$_.LastSignIn -as [datetime]} -Descending
$Report | Out-GridView

[array]$OldDevices = $Report | Where-Object {$_.'Days Since sign in' -ge 365}
Write-Host ("There are {0} devices that have not signed in for over a year" -f $OldDevices.Count)
$OldDevices | Format-Table Device, 'Device last signed in', 'Days since sign in', Owner, OwnerUPN, Department -AutoSize

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.