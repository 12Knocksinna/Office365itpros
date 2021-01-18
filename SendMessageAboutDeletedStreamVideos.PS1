# SendMessageAboutDeletedStreamVideos.PS1
# A script to show how to report videos waiting in the Stream recycle bin and send the report to an administrator account
# V1.1 11-June-2020
# https://github.com/12Knocksinna/Office365itpros/blob/master/SendMessageAboutDeletedStreamVideos.PS1
#
# Make sure that we have credentials to send the message
If (!$O365Cred) {$O365Cred = Get-Credential}

# And that we're connected to Exchange Online
Try { $OrgName = (Get-OrganizationConfig).Name }
   Catch  {
      Write-Host "Your PowerShell session is not connected to Exchange Online."
      Write-Host "Please connect to Exchange Online using an administrative account and retry."
      Break }

$StartDate = (Get-Date).AddDays(-30); $EndDate = (Get-Date) 
$HTMLReportFile = "c:\temp\StreamDeletedVideos.html" 
$Records = (Search-UnifiedAuditLog -Operations StreamDeleteVideo -StartDate $StartDate -EndDate $EndDate -ResultSize 2000)
If ($Records.Count -eq 0) {
    Write-Host "No audit records for Stream video uploads found." }
Else {
       Write-Host "Processing" $Records.Count "audit records..."
    $Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report
    # Scan each audit record to extract information
    ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      $DaysElapsed = ($EndDate - (Get-Date($AuditData.CreationTime)))
        $ReportLine = [PSCustomObject] @{
           User                  = $AuditData.UserId
           "Video Name"          = ($AuditData.OperationDetails  | ConvertFrom-Json).Name
           Action                = "Deleted Video"
           "Deleted on"          = Get-Date($AuditData.CreationTime) -format g
           "Days Since Deletion" = $DaysElapsed.Days
           "Days Remaining"      = (30 - $DaysElapsed.Days)          
           "Video Identifier"    = $AuditData.EntityPath  }
      $Report.Add($ReportLine) } }
CLS
Write-Host "Report produced - emailing it to the Administrator account"
# Create HTML Report
$HtmlHeader ="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 28px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 20px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
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
           <p><h1><u>Stream Videos in Recycle Bin</h1></u></p>
           <p><h2>Report Generated: " + $EndDate + "</h2></p>
<p><h3>The following videos are in the Stream recycle bin and will be deleted after 30 days.</h3></p></div>"
# Add information about deleted videos
$HtmlBody = $Report | Sort "Days Remaining" -Descending | ConvertTo-Html -Fragment
$HtmlTail = "<p><b>Please review and make sure that it is OK to allow these videos to be deleted. After Stream removes videos from its Recycle Bin, the videos are no longer recoverable.</b></p>"
$HtmlReport = $HtmlHeader + $HtmlBody + $HtmlTail
$HtmlReport | Out-File $HtmlReportFile -Encoding UTF8
# Set TLS
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
# Import into email body
$HtmlMsg =  Get-Content $HtmlReportFile
$MsgFrom = $O365Cred.UserName ; $SmtpServer = "smtp.office365.com" ; $SmtpPort = '587'
$MsgTo = ("Administrator@domain.com", "Tony.Redmond@domain.com") Change this to add your own set of addressees
# Construct the message parameters and send it off...
    $MsgParam = @{
      To = $MsgTo
       From = $MsgFrom
       Subject = "Deleted Stream Videos in Recycle Bin"
       Body = $HtmlReport
       SmtpServer = $SmtpServer
       Port = $SmtpPort
       Credential = $O365Cred }
     Send-MailMessage @msgParam -UseSSL -BodyAsHTML 

Write-Host "All done. The output file is also available in" $HtmlReportFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
