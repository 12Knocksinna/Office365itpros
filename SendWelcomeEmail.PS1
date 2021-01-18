# SendWelcomeEmail.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/SendWelcomeEmail.PS1
# How to send a welcome message to new mailboxes using SMTP AUTH
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
# Date to Check for new accounts - we use the last 7 days here, but that's easily changable.
[string]$CheckDate = (Get-Date).AddDays(-7)
# Make sure that we have valid credentials
If (-not $O365Cred) { #Make sure we have credentials
    $O365Cred = (Get-Credential)}
# Message is from the logged in account
$MsgFrom = $O365Cred.UserName ; $SmtpServer = "smtp.office365.com" ; $SmtpPort = '587'

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

# Find all mailboxes created in the target period
$Users = (Get-ExoMailbox -Filter "WhenMailboxCreated -gt '$CheckDate'" -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties WhenMailboxCreated | Select WhenMailboxCreated, DisplayName, UserPrincipalName, PrimarySmtpAddress)

ForEach ($User in $Users) {
      $EmailRecipient = $User.PrimarySmtpAddress
      Write-Host "Sending welcome email to" $User.DisplayName
      $htmlHeaderUser = "<h2>New User " + $User.DisplayName + "</h2>"
      $htmlline1 = "<p><b>Welcome to Office 365</b></p>"
      $htmlline2 = "<p>You can open Office 365 by clicking <a href=http://www.portal.office.com>here</a> </p>"
      $htmlline3 = "<p>Have a great time and be sure to call the help desk if you need assistance.</p>"
      $htmlbody = $htmlheaderUser + $htmlline1 + $htmlline2 + $htmlline3 + "<p>"
      $HtmlMsg = "</body></html>" + $HtmlHead + $HtmlBody
     # Construct the message parameters and send it off...
        $MsgParam = @{
         To = $EmailRecipient
         From = $MsgFrom
         Subject = "A Hundred Thousand Welcomes"
         Body = $HtmlMsg
         SmtpServer = $SmtpServer
         Port = $SmtpPort
         Credential = $O365Cred }
         Send-MailMessage @msgParam -UseSSL -BodyAsHTML}
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
