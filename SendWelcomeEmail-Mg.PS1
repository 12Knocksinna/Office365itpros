# SendWelcomeEmail-Mg.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/SendWelcomeEmail-Mg.PS1
# How to send a welcome message to new mailboxes using the Send-MgUserMessage cmdlet from the Microsoft Graph SDK for PowerShell
# Needs version 1.7.0 or later of the Microsoft.Graph.Mail and Microsoft.Graph.Users.Actions modules

# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
# Now connect to the Microsoft Graph SDK for PowerShell
Connect-MgGraph -TenantId b662313f-14fc-43a2-9a7a-d2e27f4f3478 -Scope "Mail.Send, Mail.ReadWrite"
$Details = Get-MgContext
$Scopes = $Details | Select -ExpandProperty Scopes
$Scopes = $Scopes -Join ", "
$ProfileName = (Get-MgProfile).Name
$TenantName = (Get-MgOrganization).DisplayName
CLS
Write-Host "Microsoft Graph Connection Information"
Write-Host "--------------------------------------"
Write-Host " "
Write-Host ("Connected to Tenant {0} ({1}) as account {2}" -f $Details.TenantId, $TenantName, $Details.Account)
Write-Host "+-------------------------------------------------------------------------------------------------------------------+"
Write-Host ("Profile set as {0}. The following permission scope is defined: {1}" -f $ProfileName, $Scopes)
Write-Host ""
Write-Host "Also successfully connected to the Exchange Online management PowerShell module"

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
#Header for the message
$HtmlBody = "<body>
     <h1>Welcome to $($TenantName)</h1>
     <p><strong>Generated:</strong> $(Get-Date -Format g)</p>  
     <h2><u>We're Pleased to Have You Here</u></h2>"

$htmlline1 = "<p><b>Welcome to Office 365</b></p>"
$htmlline2 = "<p>You can open Office 365 by clicking <a href=http://www.portal.office.com>here</a> </p>"
$htmlline3 = "<p>Have a great time and be sure to call the help desk if you need assistance.</p>"

$MsgSubject = "A warm welcome to $($TenantName)"

# Date to Check for new accounts - we use the last 7 days here, but that's easily changable.
[string]$CheckDate = (Get-Date).AddDays(-7)
# Find all mailboxes created in the target period
[array]$Users = (Get-ExoMailbox -Filter "WhenMailboxCreated -gt '$CheckDate'" -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties WhenMailboxCreated | Select WhenMailboxCreated, DisplayName, UserPrincipalName, PrimarySmtpAddress)
If ($Users -eq $Null) { Write-Host ("No mailboxes found that were created before {0}. Exiting!" -f $CheckDate) }

# Define attachment
$AttachmentFile = "c:\temp\WelcometoOffice365ITPros.docx"
$EncodedAttachmentFile = [convert]::ToBase64String((Get-Content $AttachmentFile -Encoding byte)) 
$MsgAttachment = @{
    "@odata.type"= "#microsoft.graph.fileAttachment"
    name = ($AttachmentFile -split '\\')[-1]
    contentBytes = $EncodedAttachmentFile
}

# Create and send welcome email message to each of the new mailboxes
ForEach ($User in $Users) {
      # Add the recipient using the mailbox's primary SMTP address
      $EmailAddress  = @{address = $User.PrimarySmtpAddress} 
      $EmailRecipient = @{EmailAddress = $EmailAddress}  
      Write-Host "Sending welcome email to" $User.DisplayName
      # Customize the message 
      $htmlHeaderUser = "<h2>New User " + $User.DisplayName + "</h2>"    
      $htmlbody = $htmlheaderUser + $htmlline1 + $htmlline2 + $htmlline3 + "<p>"
      $HtmlMsg = "</body></html>" + $HtmlHead + $HtmlBody
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

# Example of how to add BCC Recipients by building an array of addresses from some source info (in this case, another array)
# $BCCRecipients = @("Brian.Weakliam@Office365itpros.com", "Kim.Akers@Office365itpros.com")
# $BCC = $BCCRecipients | % {@{emailAddress = @{ address = $_ }}}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
