# SendWelcomeEmailGraph.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/SendWelcomeEmailGraph.PS1
# How to send a welcome message to new mailboxes using SMTP AUTH
# Updated version of the SendWelcomeEmail.PS1 script using the Graph API call to send mail instead of the Send-MailMessage cmdlet
# Needs an app registered in Azure AD with consent given for the Application Mail.Send permission
# 
# Check we have the right module loaded
$Modules = Get-Module
If ("ExchangeOnlineManagement" -notin  $Modules.Name) {Write-Host "Please connect to Exchange Online Management  before continuing...";break}
# Date to Check for new accounts - we use the last 7 days here, but that's easily changable.
[string]$CheckDate = (Get-Date).AddDays(-7)

# Get Graph access token - change these values for the app you use.
$AppId = "770e01d1-ce75-46ba-a054-4b61c787f682"
$AppSecret = "~_Q2.IPI3kpbmhOG.VtCb1r0_J9G4-D8jG"
$TenantId = "a632313f-14fc-43a2-9a7a-d2e27f4f3478"

# Construct URI and body needed for authentication
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"
}

$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
# Unpack Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
$Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $Token" }

# Define some variables for the message
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
     <h1>Welcome to Our Company</h1>
     <p><strong>Generated:</strong> $(Get-Date -Format g)</p>  
     <h2><u>We're Pleased to Have You Here</u></h2>"

# Other lines
$htmlline1 = "<p><b>Welcome to Office 365</b></p>"
$htmlline2 = "<p>You can open Office Online by clicking <a href=https://www.office.com/?auth=2>here</a> </p>"
$htmlline3 = "<p>Have a great time and be sure to call the help desk if you need assistance.</p>"

# Find all mailboxes created in the target period
[datetime]$DisplayDate = $CheckDate
Write-Host ("Finding User mailboxes created since {0}" -f (Get-Date($DisplayDate) -format g))
$Users = (Get-ExoMailbox -Filter "WhenMailboxCreated -gt '$CheckDate'" -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties WhenMailboxCreated | Select WhenMailboxCreated, DisplayName, UserPrincipalName, PrimarySmtpAddress)
If (!($Users)) {Write-Host "No users found - exiting" ; break} Else { Write-Host $Users.Count "mailboxes found to process"}

# Mail.Send can send from any user. In this example, because we're sending messages to welcome people, we look up the mailboxes on the system to
# find one tagged as the HR admin. You can use whatever method you like as long as the result is a resolveable SMTP address. If you use an application
# access policy to restrict access to the API, make sure this user is included.
$MsgFrom = Get-Recipient -Filter {CustomAttribute1 -eq "HRAdmin"} | Select -ExpandProperty PrimarySmtpAddress

If ($MsgFrom -eq $Null) { Write-Host "Can't find user defined as the sender for the messages - exiting!"; break}

# Use the same approach to define CC recipients for the message
$ccRecipient1 = Get-Recipient -Filter {CustomAttribute1 -eq "HRCC1"} | Select -ExpandProperty PrimarySmtpAddress
$ccRecipient2 = Get-Recipient -Filter {CustomAttribute1 -eq "HRCC2"} | Select -ExpandProperty PrimarySmtpAddress
If (($ccRecipient1 -eq $Null) -or ($ccRecipient2 -eq $Null)) { Write-Host "CC Recipients are not defined. Exiting!" ; break}
Write-Host "Messages will be sent from" $MsgFrom "and CC'd to" $ccRecipient1 "and" $ccRecipient2

# Define attachment to send to new users
$AttachmentFile = "C:\temp\WelcomeToOffice365ITPros.docx"
$ContentBase64 = [convert]::ToBase64String( [system.io.file]::readallbytes($AttachmentFile))
$SentMessages = 0

# Create and send welcome message for each user
ForEach ($User in $Users) {
      $EmailRecipient = $User.PrimarySmtpAddress
      Write-Host "Sending welcome email to" $User.DisplayName
      $MsgSubject = "A Hundred Thousand Welcomes to " + $User.DisplayName
      $htmlHeaderUser = "<h2>New User " + $User.DisplayName + "</h2>"
      $htmlline1 = "<p><b>Welcome to Office 365</b></p>"
      $htmlline2 = "<p>You can open Office Online by clicking <a href=https://www.office.com/?auth=2>here</a> </p>"
      $htmlline3 = "<p>Have a great time and be sure to call the help desk if you need assistance.</p>"
      $htmlbody = $htmlheaderUser + $htmlline1 + $htmlline2 + $htmlline3 + "<p>"
      $HtmlMsg = "</body></html>" + $HtmlHead + $HtmlBody
# Create message body and properties and send
        $MessageParams = @{
          "URI"         = "https://graph.microsoft.com/v1.0/users/$MsgFrom/sendMail"
          "Headers"     = $Headers
          "Method"      = "POST"
          "ContentType" = 'application/json'
          "Body" = (@{
                "message" = @{
                "subject" = $MsgSubject
                "body"    = @{
                    "contentType" = 'HTML' 
                     "content"     = $htmlMsg }
          "attachments" = @(
             @{
              "@odata.type" = "#microsoft.graph.fileAttachment"
              "name" = $AttachmentFile
              "contenttype" = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              "contentBytes" = $ContentBase64 } )  
           "toRecipients" = @(
           @{
             "emailAddress" = @{"address" = $EmailRecipient }
           } ) 
          "ccRecipients" = @(
           @{
             "emailAddress" = @{"address" = $ccRecipient1 }
           } ,
            @{
             "emailAddress" = @{"address" = $ccRecipient2 }
           } )       
         }
      }) | ConvertTo-JSON -Depth 6
   }
   # Send the message
   Invoke-RestMethod @Messageparams
   $SentMessages++
}

Write-Host "All done." $SentMessages "welcome messages sent to happy users"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
