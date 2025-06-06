# ReportExternalEmailSent.PS1
# Script to demonstrate how to create a report about the external email activity of individuals identified through 
# membership of a distribution group. See https://office365itpros.com/2022/04/13/message-tracing-email-activity/ for details
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportExternalEmailSent.PS1
# Uses cmdlets from the Exchange Online module and the Microsoft Graph PowerShell SDK.

# Updated 4-June-2025 to replace Get-MessageTrace with Get-MessageTraceV2, which is the cmdlet that will be used in the future.

# Check that we have the necessary modules loaded
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) 
   {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break
}

# Connect to the Microsoft Graph PowerShell SDK so that we can send email
Connect-MgGraph -Scope "Mail.Send, Mail.ReadWrite" -NoWelcome

$TenantName = (Get-OrganizationConfig).DisplayName
$StartDate = (Get-Date).AddDays(-10)
$EndDate = Get-Date
$ExternalHTMLFile = "c:\temp\ExternalEmailSent.html"

[array]$Users = Get-DistributionGroupMember -Identity Monitored.Users | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails
# Drop anything else but mailboxes
[array]$Users = $Users | Where-Object {$_.RecipientTypeDetails -eq "UserMailbox"}

Write-Host ("Checking external email for {0} mailboxes" -f $Users.count)
$Report = [System.Collections.Generic.List[Object]]::new() 
ForEach ($User in $Users) {
   Write-Host ("Checking messages sent by {0}" -f $User.DisplayName)
   # Get message information for the last ten days and filter so that we end up with just external addresses
   [string]$SenderAddress = $User.PrimarySmtpAddress.toString()
   # Old code - Get-MessageTrace cmdlet will be deprecated in the future, so we use Get-MessageTraceV2 instead
   # [array]$Messages = Get-MessageTrace -StartDate $StartDate -EndDate $EndDate -SenderAddress $SenderAddress -Status Delivered | Where-Object {$_.RecipientAddress -notlike "*@Office365itpros*"}
   [array]$Messages = Get-MessageTraceV2 -StartDate $StartDate -EndDate $EndDate -SenderAddress $SenderAddress -Status Delivered | Where-Object {$_.RecipientAddress -notlike "*@Office365itpros*"}
   
   ForEach ($M in $Messages) {
      $ReportLine = [PSCustomObject][Ordered]@{
         Date      = Get-Date($M.Received) -format g 
         User      = $M.SenderAddress
         Recipient = $M.RecipientAddress
         Subject   = $M.Subject
         MessageId = $M.MessageId 
      }
      $Report.Add($ReportLine)
   } #End Foreach messages
} # End ForEach Users

# Create HTML content
$Report = $Report | Sort-Object User
$Today = Get-Date -format f
$HtmlStart = 
'<html><head><font face="Segoe UI"><Title>External Email Activity Report</Title></font></p><p><font face="Segoe UI"><h2>Tenant: ' + ($TenantName) + '</h2></p><p><font face="Segoe UI"><h3>Generated: ' + $Today + '</h3></font></p>
 <table align="center" border="1" cellpadding="1" cellspacing="1" style="background-color:#e6e6fa; border-style:hidden">
	<caption>
	<h1><span style="font-size:30px"><strong><span style="color:#3498db">External Email Analysis Report</span></strong></span></h1>
	</caption>
	<thead>
		<tr>
			<th scope="col">Timestamp</th>
			<th scope="col">Semder</th>
			<th scope="col">Recipient</th>
			<th scope="col">Subject</th>
		</tr>
	</thead>
	<tbody>'

# Insert individual message info
ForEach ($R in $Report) {
   $DataLines += "<tr><td>$($R.Date)</td><td>$($R.User)</td><td>$($R.'Recipient')</td><td>$($R.'Subject')</td></tr>"
}

$Htmlend = "</tbody></table></body></html>'"
$Body = $HTMLStart + $DataLines + $HTMLEnd
$Body | Out-File $ExternalCSVFile

# Define who the message comes from (the signed in user for the Graph session) and the recipient (the person who manages the DL)
$MsgFrom = (Get-MgContext).Account
[string]$EmailRecipient = (Get-DistributionGroup -Identity Monitored.Users).ManagedBy
[string]$EmailRecipientAddress = (Get-EXOMailbox -Identity $EmailRecipient).PrimarySmtpAddress
$MsgSubject = "User External Email Report"

# Add the HTML file as an attachment
$EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($ExternalHTMLFile))

$MsgAttachments = @(
    @{
	"@odata.type" = "#microsoft.graph.fileAttachment"
	Name = ($ExternalHTMLFile -split '\\')[-1]
	ContentBytes = $EncodedAttachmentFile
	}
)

# Add the recipient using the mailbox's primary SMTP address
$EmailAddress  = @{address = $EmailRecipientAddress} 
$EmailToRecipient = @{EmailAddress = $EmailAddress}  
Write-Host "Sending report to" $EmailRecipientAddress
$HtmlHeader = "<h2>User External Email Activity Report</h2>"    
$HtmlMsg = "</body></html>" + $htmlheader + $Body + "<p>"
     
# Construct the message body
$MessageBody = @{
   Content = "$($HtmlMsg)"
   ContentType = 'html'
}

# Create a draft message in the signed-in user's mailbox
Try {
   Write-Host "Creating message to send..."
   $NewMessage = New-MgUserMessage -UserId $MsgFrom -Body $MessageBody -ToRecipients $EmailToRecipient -Subject $MsgSubject -Attachments $MsgAttachments -ErrorAction Stop
# Send the message
} Catch {
   Write-Host "Error creating message: $($_.Exception.Message)"
   Break
}
Try {
   Send-MgUserMessage -UserId $MsgFrom -MessageId $NewMessage.Id -ErrorAction Stop
   Write-Host ("Report emailed to {0} and HTML file is available at {1}" -f $EmailRecipientAddress, $ExternalHTMLFile)
} Catch {
   Write-Host "Error sending message: $($_.Exception.Message)"
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.