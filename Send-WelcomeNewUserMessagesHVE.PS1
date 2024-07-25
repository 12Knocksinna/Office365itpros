# Send-WelcomeNewUserMessagesHVE.PS1
# The classic send new welcome messages reprised for Exchange Online High-Volume Email
# V1.0 23-July-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Send-WelcomeNewUserMessagesHVE.PS1

# Configure the period to find new mailboxes. The default is 7 days, but you can  use whatever period you like
$CheckDate = (Get-Date).AddDays(-7)
# Connect to Exchange Online
$Modules = Get-Module | Select-Object -ExpandProperty Name
If ('ExchangeOnlineManagement' -notin $Modules) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}
# Find mailboxes matching the search criteria
Write-Host "Looking for recently created mailboxes..."
[array]$Users = (Get-ExoMailbox -Filter "WhenMailboxCreated -gt '$CheckDate'" -RecipientTypeDetails UserMailbox `
    -ResultSize Unlimited -Properties WhenMailboxCreated | Select-Object WhenMailboxCreated, DisplayName, UserPrincipalName, PrimarySmtpAddress)

IF (!$Users) {
    Write-Host "No new mailboxes found"
    Break
}

# Get the credentials for the HVE account we want to use
$HVECredentials = Get-Credential -Message "Enter credentials for the HVE account" -UserName 'HVE1@office365itpros.com'
# BCC address for each message - also used in the message content
$BCCAddress = "Customer.Services@contoso.com"

# Enforce TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

ForEach ($User in $Users) {
    Write-Host ("Processing mailbox {0} created on {1}" -f $User.UserPrincipalName, $User.WhenMailboxCreated)
    # Set up message content    
    $Content = "<h2>Welcome to Microsoft 365</h2><p>" +
    ("<p>Dear {0},</p>" -f $User.displayName) +
    "<p>Welcome to Microsoft 365. Your mailbox is now ready for use. Please sign in to Outlook on the web at " +
    "<a href='https://outlook.office.com'>https://outlook.office.com</a> to access your email.</p>" +
    "<p>If you need help, please check the <a href='https://office365itpros.com/office-365-for-it-pros-faq/'>FAQ</a> or send email to the " +
    ("<a href='mailto:{0}'>End User Support Desk</a>.</p>" -f $BCCAddress)  

    $SendHVEMessageParams = @{
        From = 'HVE02@office365itpros.com'
        To = $User.PrimarySmtpAddress
        Bcc = $BCCAddress
        Subject = "Welcome to Microsoft 365"
        Body = $Content
        Credential = $HVECredentials
        UseSsl = $true
        SmtpServer = 'smtp-hve.office365.com'
        Port = 587
        BodyAsHtml = $True
    }   
    Try {
        Send-MailMessage @SendHVEMessageParams -ErrorAction Stop   
    } Catch {
        Write-Host ("Failed to send email to {0} with error {1}" -f $Recipient, $_.Exception.Message)  
    }
}