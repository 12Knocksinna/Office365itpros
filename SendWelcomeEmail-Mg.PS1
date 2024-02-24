# SendWelcomeEmail-Mg.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/SendWelcomeEmail-Mg.PS1
# How to send a welcome message to new mailboxes using the Send-MgUserMessage cmdlet from the Microsoft Graph SDK for PowerShell
# Needs version 1.7.0 or later of the Microsoft.Graph.Mail and Microsoft.Graph.Users.Actions modules

# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select-Object -ExpandProperty Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
# Now connect to the Microsoft Graph SDK for PowerShell
Connect-MgGraph -NoWelcome -Scopes "Mail.Send, Mail.ReadWrite"

$TenantName = (Get-MgOrganization).DisplayName
Clear-Host

# Make sure that we have valid credentials
If (-not $O365Cred) { #Make sure we have credentials
    $O365Cred = (Get-Credential)}
# Message is from the logged in account
$MsgFrom = $O365Cred.UserName

# Define some variables used to construct the HTML content in the message body
#HTML header with styles
$htmlhead="<html>
     <style>
      BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    </style>"

#Content for the message - obviously this is very customizable to reflect what you want to say to new users
$HtmlBody = "<body>
     <h1>Welcome to $($TenantName)</h1>
     <p><strong>Generated:</strong> $(Get-Date -Format g)</p>  
     <h2><u>We're Pleased to Have You Here</u></h2>
     <p><b>Welcome to your new Office 365 account</b></p>
     <p>You can open your account to access your email and documents by clicking <a href=http://www.portal.office.com>here</a> </p>
     <p>Have a great time and be sure to call the help desk if you need assistance. And be sure to read all the great articles about Office 365 published on Practical365.com.</p>"

$MsgSubject = "A warm welcome to $($TenantName)"

# Date to Check for new accounts - we use the last 7 days here, but that's easily changable.
[string]$CheckDate = (Get-Date).AddDays(-7)
# Find all mailboxes created in the target period
[array]$Users = (Get-ExoMailbox -Filter "WhenMailboxCreated -gt '$CheckDate'" -RecipientTypeDetails UserMailbox `
 -ResultSize Unlimited -Properties WhenMailboxCreated | Select-Object WhenMailboxCreated, DisplayName, UserPrincipalName, PrimarySmtpAddress)
If (!($Users)) { 
     Write-Host ("No mailboxes found that were created before {0}. Exiting!" -f $CheckDate) 
}

# Define attachment - I use a document on a web site here, but you could use a local file
$WebAttachmentFile = "https://office365itpros.com/wp-content/uploads/2022/02/WelcomeToOffice365ITPros.docx"
$AttachmentFile = "c:\temp\WelcomeNewEmployeeToOffice365itpros.docx"
Invoke-WebRequest -uri $WebAttachmentFile -OutFile $AttachmentFile
$Attachment = (Get-Item -Path $AttachmentFile).Name
$EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($AttachmentFile))

$MsgAttachment = @(
     @{
	"@odata.type" = "#microsoft.graph.fileAttachment"
	Name = ($AttachmentFile -split '\\')[-1]
	ContentType = "text/plain"
	ContentBytes = $EncodedAttachmentFile
	}
   )

# Create and send welcome email message to each of the new mailboxes
ForEach ($User in $Users) {
      # Add the recipient using the mailbox's primary SMTP address
      $EmailAddress  = @{address = $User.PrimarySmtpAddress} 
      $EmailRecipient = @{EmailAddress = $EmailAddress}  
      Write-Host "Sending welcome email to" $User.DisplayName
      # Customize the message 
      $htmlHeaderUser = "<h2>New User " + $User.DisplayName + "</h2>"    
      # Construct the message body
      $MessageBody = @{
           content = "$($HtmlBody)"
           ContentType = 'html'
           }
     # Create a draft message in the signed-in user's mailbox
     $NewMessage = New-MgUserMessage -UserId $MsgFrom -Body $MessageBody -ToRecipients $EmailRecipient -Subject $MsgSubject -Attachments $MsgAttachment
     # Send the message
     Send-MgUserMessage -UserId $MsgFrom -MessageId $NewMessage.Id  
} # End ForEach User

Write-Host "All done. Messages sent!"
# Clean up
Remove-item $attachmentfile
Remove-item "c:\TempForScriptxxx"

# Example of how to add BCC Recipients by building an array of addresses from some source info (in this case, another array)
# $BCCRecipients = @("Brian.Weakliam@Office365itpros.com", "Kim.Akers@Office365itpros.com")
# $BCC = $BCCRecipients | % {@{emailAddress = @{ address = $_ }}}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
