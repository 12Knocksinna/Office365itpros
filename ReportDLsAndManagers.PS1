# ReportDLsAndManagers.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportDLsAndManagers.PS1
# Quick and dirty script to report the managers of distribution lists
CLS
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}

$OrgName = (Get-OrganizationConfig).Name
$CreationDate = Get-Date -format g
$Version = "1.0"
$ReportFile = "c:\temp\DLManagersReport.html"
$CSVFileMembers = "c:\temp\DLManagersReport.csv"

Write-Host "Finding Distribution lists in" $OrgName "..."
[array]$DLs = Get-DistributionGroup -ResultSize Unlimited -Filter {RecipientTypeDetails -ne "Roomlist"} | Select DisplayName, ExternalDirectoryObjectId, ManagedBy
If (!($DLs)) { Write-Host "No distribution lists found - exiting" ; break }

$Report = [System.Collections.Generic.List[Object]]::new()
Write-Host "Reporting Distribution lists and managers..."
ForEach ($DL in $DLs) {
    $ManagerList = [System.Collections.Generic.List[Object]]::new()
    ForEach ($Manager in $DL.ManagedBy) {
       $Recipient = Get-Recipient -Identity $Manager -ErrorAction SilentlyContinue
       If (!($Recipient)) { # Can't resolve manager
           $Recipient = "Unknown user" }
       $ManagerLine = [PSCustomObject][Ordered]@{  
         DisplayName = $Recipient.DisplayName
         UPN         = $Recipient.WIndowsLiveID }
       $ManagerList.Add($ManagerLine) 
    } # End processing managers
    $Managers = $ManagerList.DisplayName -join ", " 
    $DLLine = [PSCustomObject][Ordered]@{    
         DisplayName  = $DL.DisplayName     
         Managers     = $Managers
         EMailAddress = $DL.PrimarySmtpAddress }
    $Report.Add($DLLine)
} # End processing DL


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
           <p><h1>Distribution List Manager Report</h1></p>
           <p><h2><b>For the " + $Orgname + " organization</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format g) + "</h3></p></div>"

$htmlbody1 = $Report | ConvertTo-Html -Fragment

$htmltail = "<p>Report created for: " + $OrgName + "</p>" +
             "<p>Created: " + $CreationDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of distribution lists found:    " + $DLs.Count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Distribution List Manager Report<b> " + $Version + "</b>"	

$htmlreport = $htmlhead + $htmlbody1 + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8

$Report | Export-CSV -NoTypeInformation $CSVFileMembers
CLS
Write-Host "All done. Output files are" $CSVFileMembers "and" $ReportFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
