# Report-ContainerManagementLabels.PS1
# A script to create a report about container management labels and highlight groups that don't have a label. Also
# highlights groups without owners
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ContainerManagementLabels.PS1

Connect-MgGraph -Scopes Directory.Read.All

$Version = "1.0"
$HtmlReportFile = "c:\temp\ContainerManagementLabels.html"
$CSVReportFile = "c:\temp\ContainerManagementLabels.CSV"
$Organization = Get-MgOrganization

Write-Host "Finding Microsoft 365 Groups to process..."
[array]$Groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'unified')" -All
If (!($Groups)) { Write-Host "Whoops - can't find any Microsoft 365 Groups" ; break }

Write-Host ("Found {0} groups - now checking container management labels" -f $Groups.count)
$Groups = $Groups | Sort-Object DisplayName

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Group in $Groups) {
 $LabelId = $Null; $LabelName = $Null; $GroupOwnerNames = $Null; $Status = $Null

 $Uri = ("https://graph.microsoft.com/v1.0/groups/{0}?`$select=assignedLabels" -f $Group.Id)
 $LabelData = Invoke-GraphRequest -Uri $Uri
 $LabelName = $LabelData.assignedLabels.displayName
 $LabelId   = $LabelData.assignedLabels.labelId
 [array]$GroupOwners = Get-MgGroupOwner -GroupId $Group.Id
 If ($GroupOwners) {
   [string]$GroupOwnerNames = $GroupOwners.additionalProperties.displayName -join ", " }

 # Issue warning if no label found
 If (!($LabelName)) {
    Write-Host ("The {0} group has no label. Owner(s) {1}" -f $Group.displayName, $GroupOwnerNames) -foregroundcolor Red
    $Status = "*** Check Label ***"
 }

 # Issue warning if no group owners found
 If (!($GroupOwners)) {
    Write-Host ("The {0} group has no owners." -f $Group.displayName) -foregroundcolor Red
    $Status = $Status + "*** No Group Owners ***"
 }

 $ReportLine = [PSCustomObject][Ordered]@{
    Id        = $Group.Id
    Name      = $Group.DisplayName
    Owners    = $GroupOwnerNames
    Label     = $LabelName
    LabelId   = $LabelId 
    Status    = $Status }                  
 $Report.Add($ReportLine) 

} # End processing groups

# Generate the report files
$HtmlHeading ="<html>
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
           <p><h1>Microsoft 365 Groups and Container Management Labels Report</h1></p>
           <p><h3>Generated: " + (Get-Date -format 'dd-MMM-yyyy hh:mm tt') + "</h3></p></div>"
$HtmlData = $Report | ConvertTo-Html -Fragment
$HtmlReport = $HtmlHeading + $HtmlData 

# Figure out labels used
$LabelData =  [System.Collections.Generic.List[Object]]::new()
[array]$LabelUnique = $Report | Where-Object {$_.Label -ne $Null} | Select-Object Label | Sort-Object Label -Unique
ForEach ($SingleLabel in $LabelUnique) {
   [array]$LData = $Report | Where-Object {$_.Label -eq $SingleLabel.Label} 
   $LabelInfo  = [PSCustomObject][Ordered]@{
      Label = $SingleLabel.Label
      Count = $LData.Count }
   $LabelData.Add($LabelInfo)
}

$HtmlHeading = "<p><h3>Label Summary</h3>"
$HtmlData = $LabelData | ConvertTo-Html -Fragment
$HtmlReport = "<p>" + $HtmlReport + "<p>" + $HtmlHeading + $HtmlData

# Add a section about groups missing labels if there are any
[array]$MissingLabels = $Report | Where-Object {$_.LabelId -eq $Null}
If ($MissingLabels) {
  $HtmlHeading = "<p><h3>Microsoft 365 Groups without a Sensitivity Label</h3>"
  $HtmlData = $MissingLabels | ConvertTo-Html -Fragment
  $HtmlReport = $HtmlReport + $HtmlHeading + $HtmlData 
}

# Add section about missing owners (if any)
[array]$MissingOwners = $Report | Where-Object {$_.Owners -eq $Null -or $_.Owners.length -eq 0}
If ($MissingOwners) {
  $HtmlHeading = "<p><h3>Microsoft 365 Groups without owners</h3>"
  $HtmlData = $MissingOwners | ConvertTo-Html -Fragment
  $HtmlReport = $HtmlReport + $HtmlHeading + $HtmlData 
}

# Create the HTML report
$Htmltail = "<p>Report created for: " + ($Organization.DisplayName) + "</p><p>" +
             "<p>Number of Microsoft 365 Group:    " + $Groups.count + "</p>" +
             "<p>Number of labels used:            " + $LabelUnique.count + "</p>" +
             "<p>Number of groups without a label: " + $MissingLabels.count + "</p>" +
             "<p>Number of groups without owners:  " + $MissingOwners.count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------" +
             "<p>Microsoft 365 Groups and Container Management Sensitivity Labels <b>" + $Version + "</b>"	
$HtmlReport = $HtmlHead + $HtmlReport + $HtmlTail
$HtmlReport | Out-File $HtmlReportFile  -Encoding UTF8

Write-Host ""
Write-Host "All done"
Write-Host ""
Write-Host ("{0} Microsoft 365 Groups found" -f $Groups.count)
Write-Host ("{0} Groups found without a label" -f $MissingLabels.count)
Write-Host ("{0} Groups found without owners" -f $MissingOwners.count)

$Report | Export-CSV -NoTypeInformation $CSVReportFile
Write-Host ("Output files are available in {0} and {1}" -f $HtmlReportFile, $CSVReportFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
