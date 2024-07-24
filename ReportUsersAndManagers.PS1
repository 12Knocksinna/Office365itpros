# ReportUsersAndManagers.PS1
# A script to show how to create a report about the links between managers and employees stored in Azure AD
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportUsersAndManagers.PS1

Connect-ExchangeOnline
Connect-MgGraph -Scopes Directory.Read.All

$HtmlHead="<html>
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
           <p><h1>Microsoft 365 Managers and Direct Reports</h1></p>
           <p><h3>Generated: " + (Get-Date -format 'dd-MMM-yyyy hh:mm tt') + "</h3></p></div>"

$Version = "1.0"
$HtmlReportFile = "c:\temp\UsersAndManagers.html"
$CSVReportFile = "c:\temp\UsersAndManagers.CSV"
$Organization = Get-MgOrganization

# Find user accounts
Write-Host "Finding user accounts..."
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All `
    -PageSize 999 -Property Id, displayName, userprincipalName, department, city, country
If (!($Users)) { 
   Write-Host "No licensed Entra ID accounts found - exiting" 
   break
}
Write-Host ("Found {0} user accounts - checking management information..." -f $Users.count)

# The accounts used for Room and shared mailboxes can hold licenses (for example, for Teams Rooms Devices or to get a larger quota for a shared mailbox)
# so we remove these accounts here
$OriginalUserAccounts = $Users.Count
Write-Host "Finding Entra ID accounts used for shared and room mailboxes..."
[array]$NonUserAccounts = Get-ExoMailbox -RecipientTypeDetails SharedMailbox, RoomMailbox -ResultSize Unlimited | Select-Object UserPrincipalName, ExternalDirectoryObjectId
Write-Host "Removing non-user accounts from set to be processed..."
$Users = $Users | Where-Object {$_.Id -notin $NonUserAccounts.ExternalDirectoryObjectId}
$RemovedAccounts = $OriginalUserAccounts - $Users.Count
Write-Host ("Proceeding to process {0} user accounts after removing {1} room and shared mailbox accounts" -f $Users.Count, $RemovedAccounts)

$UsersWithNoManagers = [System.Collections.Generic.List[Object]]::new() 
$ManagerReports = [System.Collections.Generic.List[Object]]::new() 

ForEach ($User in $Users) {
  $UserManager = Get-MgUserManager -UserId $User.Id -ErrorAction SilentlyContinue
  If ($Null -eq $UserManager) { # No manager
    $DataLine  = [PSCustomObject] @{
            User       = $User.DisplayName
            UPN        = $User.UserPrincipalName
            UserId     = $User.Id
            Department = $User.Department
            Office     = $User.OfficeLocation
            City       = $User.City
            Country    = $User.Country
         }
         $UsersWithNoManagers.Add($Dataline)
    Write-Host ("No manager found for {0}" -f $User.displayname) 
   }
  
  $DirectReports = Get-MgUserDirectReport -UserId $User.Id
  If ($Null -ne $DirectReports) { # Manager with direct reports
     Write-Host ("User {0} has {1} direct reports" -f $User.DisplayName, $DirectReports.count)
     ForEach ($Report in $DirectReports.AdditionalProperties) {
        $ReportUser = $Users | Where-Object {$_.UserPrincipalName -eq $Report.userPrincipalName}
        If (!($ReportUser)) { 
            $ReportUserRecord    = Get-MgUser -UserId $Report.userPrincipalName -ErrorAction SilentlyContinue 
            $ReportDisplayName   = $ReportUserRecord.DisplayName + " (unlicensed)"
            $ReportDepartment    = $ReportUserRecord.Department  
            $ReportCity          = $ReportUserRecord.City
            $ReportCountry       = $ReportUserRecord.Country 
        } Else {
            $ReportDisplayName   = $ReportUser.DisplayName
            $ReportDepartment    = $ReportUser.Department  
            $ReportCity          = $ReportUser.City
            $ReportCountry       = $ReportUser.Country 
        }
        $DataLine  = [PSCustomObject] @{
           ManagerId        = $User.Id
           Manager          = $User.DisplayName
           User             = $ReportDisplayName 
           Department       = $ReportDepartment 
           City             = $ReportCity
           Country          = $ReportCountry }
        $ManagerReports.Add($DataLine)   
    } # End ForEach Report
   } # End If Reports
}

# Add the information about managers and direct reports to the report 
[array]$Managers = $ManagerReports | Sort-Object ManagerId -Unique | Select-Object ManagerId, Manager | Sort-Object Manager
$HtmlReport = $HtmlHead
ForEach ($Manager in $Managers) {
   $DataToReport = $ManagerReports | Where-Object {$_.ManagerId -eq $Manager.ManagerId} 
   $HtmlHeading = ("<p><h1>Direct Reports for <b>{0}</b></h1>" -f $Manager.Manager)
   $DataToReport = $DataToReport | Select-Object User, Department, City, Country
   $HtmlData = $DataToReport | ConvertTo-Html -Fragment
   $HtmlReport = $HtmlReport + $HtmlHeading + $HtmlData 
}

# Add a section about users without managers
$HtmlHeading = "<p><h1>User Accounts Without a Manager</h1>"
$HtmlData = $UsersWithNoManagers | ConvertTo-Html -Fragment
$HtmlReport = $HtmlReport + $HtmlHeading + $HtmlData 

# Create the HTML report
$Htmltail = "<p>Report created for: " + ($Organization.DisplayName) + "</p><p>" +
             "<p>Number of users:                   " + $Users.count + "</p>" +
             "<p>Number of managers:                " + $Managers.count + "</p>" +
             "<p>Number of users without a manager: " + $NoManagers.count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------" +
             "<p>Microsoft 365 Managers and Direct Reports <b>" + $Version + "</b>"	
$HtmlReport = $HtmlHead + $HtmlReport + $HtmlTail
$HtmlReport | Out-File $HtmlReportFile  -Encoding UTF8

Write-Host ""
Write-Host "All done"
Write-Host ""
Write-Host ("{0} managers with direct reports found" -f $Managers.count)
Write-Host ("{0} user accounts found with no managers" -f $NoManagers.count)

$ManagerReports | Export-CSV -NoTypeInformation $CSVReportFile
Write-Host ("Output files are available in {0} and {1}" -f $HtmlReportFile, $CSVReportFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
