# SendWelcomeMessage-UserMail.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/SendWelcomeMessage-UserMail-Mg.PS1
# How to send a welcome message to new mailboxes using the Send-MgUserMail cmdlet from the Microsoft Graph SDK for PowerShell
# Needs version 1.7.0 or later of the Microsoft.Graph.Mail and Microsoft.Graph.Users.Actions modules

Function Populate-MessageRecipient {
    [cmdletbinding()]
    Param(
        [array]$ListOfAddresses )
    ForEach ($SMTPAddress in $ListOfAddresses) {
        @{
            emailAddress = @{address = $SMTPAddress}
        }    
    }    
}

Function Populate-Attachments {
    [cmdletbinding()]
    Param(
        [array]$ListOfAttachments )
	
    ForEach ($Attachment in $ListOfAttachments) {
     Write-Host "Processing" $Attachment
     $EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($Attachment))
        @{
            "@odata.type"= "#microsoft.graph.fileAttachment"
            name = ($Attachment -split '\\')[-1]
            contentBytes = $EncodedAttachmentFile
            contentType = "text/plain"
          }
        }    
    }   

# Processing
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
# Now connect to the Microsoft Graph SDK for PowerShell
Connect-MgGraph -Scope Mail.Send
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
# Message is from the logged in account - for a background app, you'd use certificate-based authentication or a managed identity
# credential information from 
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

#Content for the message - obviously, this is very customizable and should reflect what you want to say to new users
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
[array]$Users = (Get-ExoMailbox -Filter "WhenMailboxCreated -gt '$CheckDate'" -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties WhenMailboxCreated | Select WhenMailboxCreated, DisplayName, UserPrincipalName, PrimarySmtpAddress)
If (!($Users)) { Write-Host ("No mailboxes found that were created before {0}. Exiting!" -f $CheckDate) }

# Define attachments  we're only using one here, which we fetch from a web site
# if you want to add more files, add the file names to the $AttachmentsList array
$WebAttachmentFile = "https://office365itpros.com/wp-content/uploads/2022/02/WelcomeToOffice365ITPros.docx"
$AttachmentFile = "c:\temp\WelcomeNewEmployeeToOffice365itpros.docx"
Invoke-WebRequest -uri $WebAttachmentFile -OutFile $AttachmentFile

[array]$AttachmentsList = "c:\temp\WelcomeNewEmployeeToOffice365itpros.docx"
[array]$MsgAttachments = Populate-Attachments -ListOfAttachments $AttachmentsList

# Populate CC Recipienmt List - any SMTP address will do
$CcRecipientList = @( 
        'Tenant.Admins@office365itpros.com'
        'Kim.Akers@Office365itpros.com' )
[array]$MsgCcRecipients = Populate-MessageRecipient -ListOfAddresses $CcRecipientList

ForEach ($User in $Users) {
  $ToRecipientList   = @( $User.PrimarySmtpAddress )
  [array]$MsgToRecipients = Populate-MessageRecipient -ListOfAddresses $ToRecipientList
  Write-Host "Sending welcome email to" $User.DisplayName
  # Customize the message 
  $htmlHeaderUser = "<h2>New User " + $User.DisplayName + "</h2>"    
  $HtmlMsg = "</body></html>" + $HtmlHead + $htmlheaderuser + $htmlbody + "<p>"
  # Construct the message body
  $MsgBody = @{
     Content = "$($HtmlMsg)"
     ContentType = 'html'   }

  $Message =  @{subject           = $MsgSubject}
  $Message += @{toRecipients      = $MsgToRecipients}  
  $Message += @{ccRecipients      = $MsgCcRecipients}  
  $Message += @{attachments       = $MsgAttachments}
  $Message += @{body              = $MsgBody}
  $Params   = @{'message'         = $Message}
  $Params  += @{'saveToSentItems' = $True}
  $Params  += @{'isDeliveryReceiptRequested' = $True}

  Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params
}
Write-Host "All done. Messages sent!"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
