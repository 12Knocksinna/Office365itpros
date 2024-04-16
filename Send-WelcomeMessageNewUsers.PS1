# https://github.com/12Knocksinna/Office365itpros/blob/master/Send-WelcomeMessageNewUsers.PS1
# Date to Check for new accounts created in the last 7 days
[string]$CheckDate = (Get-Date).AddDays(-7)
# Make sure that we have valid credentials
If (-not $O365Cred) { #Make sure we have credentials
    $O365Cred = (Get-Credential)}
# Message is from the logged in account
$MsgFrom = $O365Cred.UserName ; $SmtpServer = "smtp.office365.com" ; $SmtpPort = '587'

# Define some variables for the message
#HTML header with styles
$htmlhead="
     <style>
      BODY{font-family: Arial; font-size: 10pt;}
    H1{font-size: 24px;}
    H2{font-size: 18px; padding-top: 10px;}
    H3{font-size: 16px; padding-top: 8px;}
    </style>"
#Header for the message

# Find all mailboxes created in the target period
$Users = (Get-ExoMailbox -Filter "WhenMailboxCreated -gt '$CheckDate'" -RecipientTypeDetails UserMailbox `
    -ResultSize Unlimited -Properties WhenMailboxCreated | Select-Object WhenMailboxCreated, DisplayName, UserPrincipalName, PrimarySmtpAddress)
$Portal = "https://www.microsoft365.com/"

ForEach ($User in $Users) {
      $EmailRecipient = $User.PrimarySmtpAddress
      Write-Host "Sending welcome email to" $User.DisplayName
      $htmlHeaderUser = "<h2>New User " + $User.DisplayName + "</h2>"
      $htmlline1 = "<p><b>Welcome to Office 365</b></p>"
      $htmlline2 = ('<p>Access Microsoft 365 apps by clicking here <a href="{0}" here</a> </p>' -f $portal)
      $htmlline3 = "<p>Have a great time and be sure to call the help desk if you need assistance.</p>"
      $htmlbody = $htmlheaderUser + $htmlline1 + $htmlline2 + $htmlline3 + "<p>"
      $HtmlMsg = "" + $HtmlHead + $HtmlBody
     # Construct the message parameters and send it off...
        $MsgParam = @{
         To = $EmailRecipient
         From = $MsgFrom
         Subject = "A Hundred Thousand Welcomes"
         Body = $HtmlMsg
         SmtpServer = $SmtpServer
         Port = $SmtpPort
         Credential = $O365Cred }
         Send-MailMessage @msgParam -UseSSL -BodyAsHTML
        }