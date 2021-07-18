# ReportManagersAndDirectReports.PS1
# A script to report the managers and their direct reports in a tenant
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportManagersAndDirectReports.PS1
CLS
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}

$OrgName = (Get-OrganizationConfig).Name
$CreationDate = Get-Date -format g
$Version = "1.0"
$ReportFile = "c:\temp\ManagersReport.html"
$CSVFileMembers = "c:\temp\ManagersReport.csv"
$NumberDR = 0
# Find people who have direct reports
Write-Host "Finding managers..."
[array]$Managers = Get-User -Filter {DirectReports -ne $null} | Select DisplayName, UserPrincipalName, ExternalDirectoryObjectId, DistinguishedName
IF (!($Managers)) { Write-Host "No managers with direct reports found!" ; break }
$ReportsList = [System.Collections.Generic.List[Object]]::new()
$ManagersWithoutDR = [System.Collections.Generic.List[Object]]::new()
# Loop through each manager to find their direct reports
Write-Host "Finding direct reports..."
ForEach ($Manager in $Managers) {
    $Dn = $Manager.DistinguishedName
    If ($Dn -like "*'*")  {
       $DNNew = "'" + "$($Dn.Replace("'","''''"))" + "'"
       $Cmd = "Get-User -Filter 'Members -eq '$DNnew''"
       [array]$Reports = Invoke-Expression $Cmd }
    Else {
       [array]$Reports = Get-User -Filter "Manager -eq '$Dn'" }
    If ($Reports.Count -eq 0) { $ManagersWithoutDR.Add($Manager.DisplayName) }
    $ReportLine = [PSCustomObject][Ordered]@{  # Write out details of the manager's reports
       Manager             = $Manager.DisplayName
       UPN                 = $Manager.UserPrincipalName
       "Number of reports" = $Reports.Count
       "Direct Reports"    = $Reports.DisplayName -join ", " }
    $ReportsList.Add($ReportLine)
    $NumberDR = $NumberDR + $Reports.Count
} # End ForEach
 
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
           <p><h1>Manager and Direct Reports Listing</h1></p>
           <p><h2><b>For the " + $Orgname + " organization</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format g) + "</h3></p></div>"

$htmlbody1 = $ReportsList | ConvertTo-Html -Fragment
$CountNoDR = $ManagersWithoutDR.Count
$ManagersWithoutDR = $ManagersWithoutDR -join ", "
$htmltail = "<p>Report created for: " + $OrgName + "</p>" +
             "<p>Created: " + $CreationDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of managers found:              " + $Managers.Count + "</p>" +
             "<p>Number of direct reports:              " + $NumberDR + "</p>" +
             "<p>Count of managers with no reports:     " + $CountNoDR +  "</p>" +
             "<p>Managers with no direct reports:       " + $ManagersWithoutDR + "</p>" +
             "<p>Average number of reports per manager: " + [math]::Round(($NumberDR/$Managers.Count),2) + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Tenant Manager and Reports Listing <b>" + $Version + "</b>"	

$htmlreport = $htmlhead + $htmlbody1 + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8

$ReportsList | Export-CSV -NoTypeInformation $CSVFileMembers
CLS
Write-Host "All done. Output files are" $CSVFileMembers "and" $ReportFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
