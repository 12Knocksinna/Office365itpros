# ReportMembershipM365Group.PS1
# Report the membership of a Microsoft 365 Group
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportMembershipM365Group.PS1
CLS
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
       
$OrgName = (Get-OrganizationConfig).Name
$Version = "1.0"
$ReportFile = "c:\temp\M365MemberReport.html"
$CSVFile = "c:\temp\M365MemberReport.csv"

$CheckGroup = Read-Host "Enter the Microsoft 365 group to check"
$Group = Get-UnifiedGroup -Identity $CheckGroup
If (!($Group)) { Write-Host "Sorry - can't find the" $CheckGroup "group; exiting..." ; break}
$DisplayName = $Group.DisplayName
$CreationDate = Get-Date $Group.WhenCreated -format g
$MemberList = [System.Collections.Generic.List[Object]]::new()
$Members = Get-UnifiedGroupLinks -LinkType Members -Identity $Group.ExternalDirectoryObjectId
ForEach ($M in $Members) {
    $UserType = "Tenant account"
    If ($M.WindowsLiveID -like "*#EXT#*") { $UserType = "Guest account" }
    $MemberLine = [PSCustomObject][Ordered]@{
       Name             = $M.DisplayName
       UPN              = $M.WindowsLiveID
       Email            = $M.PrimarySmtpAddress
       Title            = $M.Title
       Department       = $M.Department
       Office           = $M.Office
       City             = $M.City
       PostCode         = $M.PostalCode
       Country          = $M.CountryOrRegion
       Type             = $UserType
       MemberType       = "Member" }
   $MemberList.Add($MemberLine) 
 }
$Owners = Get-UnifiedGroupLinks -LinkType Owners -Identity $Group.ExternalDirectoryObjectId
ForEach ($O in $Owners) {
    $UserType = "Tenant account"
    $OwnersCount++
    If ($O.WindowsLiveID -like "*#EXT#*") { $UserType = "Guest account" }
    $MemberLine = [PSCustomObject][Ordered]@{
       Name             = $O.DisplayName
       UPN              = $O.WindowsLiveID
       Email            = $O.PrimarySmtpAddress
       Title            = $O.Title
       Department       = $O.Department
       Office           = $O.Office
       City             = $O.City
       PostCode         = $O.PostalCode
       Country          = $O.CountryOrRegion
       Type             = $UserType
       MemberType       = "Owner" }
   $MemberList.Add($MemberLine) 
 }

# Create the HTML report
$htmlhead="<html>
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
           <p><h1>Microsoft 365 Groups Membership Listing</h1></p>
           <p><h2><b>$DisplayName</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format g) + "</h3></p></div>"

$htmlbody = $MemberList | ConvertTo-Html -Fragment
$htmltail = "<p>Report created for: " + $OrgName + "</p>" +
             "<p>Group name: " + $DisplayName + "</p>" +
             "<p>"+ $Group.Notes +"</p>" +
             "<p>Created: " + $CreationDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of members: " + $Group.GroupMemberCount + "</p>" +
             "<p>Number of guests:  " + $Group.GroupExternalMemberCount + "<p>"+
             "<p>Number of owners:  " + $OwnersCount + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Microsoft 365 Group Membership Report <b>" + $Version + "</b>"	
$htmlreport = $htmlhead + $htmlbody + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8
$MemberList | Export-CSV -NoTypeInformation $CSVFile
$MemberList| Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
