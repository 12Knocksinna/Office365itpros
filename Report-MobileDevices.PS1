# Report-MobileDevices.PS1
# An example script to show how to extract mobile device statistics from devices registred with Exchange Online mailboxes
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-MobileDevices.PS1

$HtmlHead ="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.pass{background: #B7EB83;}
	   td.warn{background: #FFF275;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Microsoft 365 Mailboxes with Synchronized Mobile Devices</h1></p>
           <p><h3>Generated: " + (Get-Date -format 'dd-MMM-yyyy hh:mm tt') + "</h3></p></div>"

$Version = "1.0"
$HtmlReportFile = "c:\temp\MobileDevices.html"
$CSVReportFile = "c:\temp\MobileDevices.csv"

Connect-ExchangeOnline

$Organization = Get-OrganizationConfig | Select-Object -ExpandProperty DisplayName
[array]$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Sort-Object DisplayName
If (!($Mbx)) { Write-Host "Unable to find any user mailboxes..." ; break }

$Report = [System.Collections.Generic.List[Object]]::new() 

[int]$i = 0
ForEach ($M in $Mbx) {
 $i++
 Write-Host ("Scanning mailbox {0} for registered mobile devices... {1}/{2}" -f $M.DisplayName, $i, $Mbx.count)
 [array]$Devices = Get-MobileDevice -Mailbox $M.DistinguishedName
 ForEach ($Device in $Devices) {
   $DaysSinceLastSync = $Null; $DaySinceFirstSync = $Null; $SyncStatus = "OK"
   $DeviceStats = Get-ExoMobileDeviceStatistics -Identity $Device.DistinguishedName
   If ($Device.FirstSyncTime) {
      $DaysSinceFirstSync = (New-TimeSpan $Device.FirstSyncTime).Days }
   If (!([string]::IsNullOrWhiteSpace($DeviceStats.LastSuccessSync))) {
      $DaysSinceLastSync = (New-TimeSpan $DeviceStats.LastSuccessSync).Days }
   If ($DaysSinceLastSync -gt 30)  {
      $SyncStatus = ("Warning: {0} days since last sync" -f $DaysSinceLastSync) }
   If ($Null -eq $DaysSinceLastSync) {
      $SyncStatus = "Never synched" 
      $DeviceStatus = "Unknown" 
   } Else {
      $DeviceStatus =  $DeviceStats.Status }
   $ReportLine = [PSCustomObject]@{
     DeviceId            = $Device.DeviceId
     DeviceOS           = $Device.DeviceOS
     Model              = $Device.DeviceModel
     UA                 = $Device.DeviceUserAgent
     User               = $Device.UserDisplayName
     UPN                = $M.UserPrincipalName
     FirstSync          = $Device.FirstSyncTime
     DaysSinceFirstSync = $DaysSinceFirstSync
     LastSync           = $DeviceStats.LastSuccessSync
     DaysSinceLastSync  = $DaysSinceLastSync
     SyncStatus         = $SyncStatus
     Status             = $DeviceStatus
     Policy             = $DeviceStats.DevicePolicyApplied
     State              = $DeviceStats.DeviceAccessState
     LastPolicy         = $DeviceStats.LastPolicyUpdateTime
     DeviceDN           = $Device.DistinguishedName }
   $Report.Add($ReportLine)
 } #End Devices
} #End Mailboxes
[array]$SyncMailboxes = $Report | Sort-Object UPN -Unique | Select-Object UPN
[array]$SyncDevices = $Report | Sort-Object DeviceId -Unique | Select-Object DeviceId
[array]$SyncDevices30 = $Report | Where-Object {$_.DaysSinceLastSync -gt 30} 
$HtmlReport = $Report | Select-Object DeviceId, DeviceOS, Model, UA, User, UPN, FirstSync, DaysSinceFirstSync, LastSync, DaysSinceLastSync | Sort-Object UPN | ConvertTo-Html -Fragment

# Create the HTML report
$Htmltail = "<p>Report created for: " + ($Organization) + "</p><p>" +
             "<p>Number of mailboxes:                          " + $Mbx.count + "</p>" +
             "<p>Number of users synchronzing devices:         " + $SyncMailboxes.count + "</p>" +
             "<p>Number of synchronized devices:               " + $SyncDevices.count + "</p>" +
             "<p>Number of devices not synced in last 30 days: " + $SyncDevices30.count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------" +
             "<p>Microsoft 365 Mailboxes with Synchronized Mobile Devices<b>" + $Version + "</b>"	
$HtmlReport = $HtmlHead + $HtmlReport + $HtmlTail
$HtmlReport | Out-File $HtmlReportFile  -Encoding UTF8

Write-Host ""
Write-Host "All done"
Write-Host ""
Write-Host ("{0} Mailboxes with synchronized devices" -f $SyncMailboxes.count)
Write-Host ("{0} Individual devices found" -f $SyncDevices.count)

$Report | Export-CSV -NoTypeInformation $CSVReportFile
Write-Host ("Output files are available in {0} and {1}" -f $HtmlReportFile, $CSVReportFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
