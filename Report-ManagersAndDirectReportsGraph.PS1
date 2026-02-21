# Report-ManagersAndDirectReportsGraph.PS1
# A script to report the managers and their direct reports in a tenant using the Microsoft Graph PowerShell SDK

# V1.0 20-Feb-2026
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ManagersAndDirectReportsGraph.PS1

Connect-Graph -Scopes User.Read.All, Organization.Read.All

[string]$Organization = (Get-MgOrganization).DisplayName

[int]$NumberDR = 0
# Find people who have direct reports
Write-Host "Finding managers..."

# Find all tenant users and expand the Manager property. A client-side filter then removes unlicensed member accounts
[array]$Users = Get-MgUser -Filter "userType eq 'Member'" -All -PageSize 500 -ExpandProperty Manager `
    -Select Id, displayName, Manager, assignedlicenses, userPrincipalName | Where-Object { $_.assignedlicenses.Count -gt 0 }
 
$ManagersList = [System.Collections.Generic.List[Object]]::new()
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($User in $Users) { 
    $ReportLine = [PSCustomObject][Ordered]@{  # Write out details of the manager's reports
        Manager             = $User.Manager.additionalProperties.displayName
        ManagerUPN          = $User.Manager.additionalProperties.userPrincipalName
        ManagerId           = $User.Manager.Id
        UPN                 = $User.UserPrincipalName
        UserId              = $User.Id
        UserDisplayName     = $User.DisplayName
    }
    $ManagersList.Add($ReportLine)
}

# Some users may not have a manager assigned, so we can report on those as well if needed
[array]$UsersWithoutManagers = $Users | Where-Object { $_.Manager.Id -eq $null } | Select-Object displayName, userPrincipalName

$Managers = $ManagersList | Where-Object {$_.ManagerId -ne $null} | Sort-Object -Property Manager -Unique | Select-Object ManagerId, Manager, ManagerUPN
ForEach ($Manager in $Managers) {
    $Reports = $ManagersList | Where-Object { $_.ManagerId -eq $Manager.ManagerId }
    $ReportLine = [PSCustomObject][Ordered]@{  # Write out details of the manager's reports
       Manager             = $Manager.Manager
       UPN                 = $Manager.ManagerUPN
       "Number of reports" = $Reports.Count
       "Direct Reports"    = $Reports.UserDisplayName -join ", " }
    $Report.Add($ReportLine)
    $NumberDR = $NumberDR + $Reports.Count
} # End ForEach

Write-Host ""
Write-Host "Managers and their direct reports:" -ForegroundColor Green
Write-Host "----------------------------------"
Write-Host ""
$Report | Format-Table Manager, 'Number of Reports', 'Direct Reports' -Wrap
Write-Host ""
Write-Host "Licensed user accounts without an assigned manager:" -ForegroundColor Yellow
$UsersWithoutManagers | Format-Table displayName, userPrincipalName -Wrap

$CreationDate = Get-Date -format 'dd-MMM-yyyy HH:mm:ss'
$Version = "1.0"

# Create the HTML report
$htmlhead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 28px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Manager and Direct Reports Listing</h1></p>
           <p><h2><b>For the " + $Organization + " organization</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format g) + "</h3></p></div>"

$htmlbody1 = $htmlHead + ($Report | ConvertTo-Html -Fragment)

$htmltail = "<p>Report created for: " + $Organization + "</p>" +
             "<p>Created: " + $CreationDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of managers found:              " + $Managers.Count + "</p>" +
             "<p>Number of direct reports:              " + $NumberDR + "</p>" +
             "<p>Average number of reports per manager: " + [math]::Round(($NumberDR/$Managers.Count),2) + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Tenant Manager and Reports Listing <b>" + $Version + "</b>"	

$htmlreport = $htmlbody1 + $htmltail
$HTMLReportFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\ManagersAndDirectReports.html"
$htmlreport | Out-File $HTMLReportFile -Encoding UTF8

# Generate the report in either Excel worksheet or CSV format, depending on if the ImportExcel module is available
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\ManagersAndDirectReports.xlsx"
    $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Managers and Reports" -Title ("Managers and Reports {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "ManagersAndReports" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\ManagersAndDirectReports.CSV"
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}   

Write-Host "HTML report available in file {0}" -f $HTMLReportFile

If ($ExcelGenerated -eq $true) {
    Write-Host ("Managers and Direct Reports data available in Excel workbook {0}" -f $ExcelOutputFile)
} Else {
    Write-Host ("Managers and Direct Reports data available in CSV file {0}" -f $CSVOutputFile)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.