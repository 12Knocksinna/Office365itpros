# Disable-JunkEmailOptions.PS1
# An Azure Automation runbook to look for recently added mailboxes and update their junk email configuration if necessary
# V1.0 2-Jan-2024
# https://github.com/12Knocksinna/Office365itpros/blob/master/Disable-JunkEmailOptions.PS1

Function Add-MessageRecipients {
    # Function to build an addressee list to send email   
    [cmdletbinding()]
        Param(
        [array]$ListOfAddresses )
        ForEach ($SMTPAddress in $ListOfAddresses) {
            @{ emailAddress = @{address = $SMTPAddress}}    
        }
} 

# replace practical365.onmicrosoft.com with your Microsoft 365 service domain
Connect-ExchangeOnline -ManagedIdentity -Organization practical365.onmicrosoft.com
Connect-MgGraph -Identity -NoWelcome

# Recipient for the email sent at the end of the script - define the addresses you want to use here
$EmailRecipient = "Tony.Redmond@office365itpros.com"
# When run interactively, email will be sent from the account running the script. This is commented out for use with Azure Automation
# If used with the Mail.Send permission in an Azure Automation runbook, the sender can be any mailbox in the organization
# $MsgFrom = (Get-MgContext).Account
$MsgFrom = "Azure.Management.Account@office365itpros.com"

# Establish how far back we look
[string]$CheckDate = (Get-Date).AddDays(-7)
# Find matching mailboxes
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Filter "WhenCreated -gt '$CheckDate'"

If ($Mbx) {
    $Report = [System.Collections.Generic.List[Object]]::new()
    ForEach ($M in $Mbx) {
        $CurrentJunkMailConfiguration = Get-MailboxJunkEmailConfiguration -Identity $M.ExternalDirectoryObjectId
        If ($CurrentJunkMailConfiguration.Enabled -eq $True) {
            Set-MailboxJunkEmailConfiguration -Identity $M.ExternalDirectoryObjectId -Enabled $False `
            -BlockedSendersAndDomains $null -TrustedSendersAndDomains $null -TrustedRecipientsAndDomains $null 
    
            $DataLine = [PSCustomObject][Ordered]@{
                Mailbox                         = $M.DisplayName
                UPN                             = $M.UserPrincipalName
                'Trusted Senders and Domains'   = $CurrentJunkMailConfiguration.TrustedSendersAndDomains.Count
                'Blocked Senders and Domains'   = $CurrentJunkMailConfiguration.BlockedSendersAndDomains.Count
                'Contacts Trusted'              = $CurrentJunkMailConfiguration.ContactsTrusted
            }
            $Report.Add($DataLine)
        } Else {
            Write-Output ("Mailbox {0} already has the junk email rule disabled" -f $M.displayName)
        }
    }   
} Else {
    Write-Output "No mailboxes found to update..."
}

$ToRecipientList   = @( $EmailRecipient )
[array]$MsgToRecipients = Add-MessageRecipients -ListOfAddresses $ToRecipientList
$MsgSubject = "Mailboxes processed for Junk EMail Settings"
$HtmlHead = "<h2>Mailbox Junk Email Setting Updates</h2><p>The following nailboxes have been disabled for the Outlook Junk Email rule.</p>"
$HtmlBody = $Report | ConvertTo-Html -Fragment 
$HtmlMsg = "</body></html><p>" + $HtmlHead + $Htmlbody + "<p>"

# Construct the message body
$MsgBody = @{
  Content = "$($HtmlMsg)"
  ContentType = 'html'  
}

$Message =  @{subject           = $MsgSubject}
$Message += @{toRecipients      = $MsgToRecipients}  
$Message += @{body              = $MsgBody}
$Params   = @{'message'         = $Message}
$Params  += @{'saveToSentItems' = $True}
$Params  += @{'isDeliveryReceiptRequested' = $True}

# And send the message using the parameters that we've filled in
Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params
Write-Output ("Message containing information about Junk Email rule updates for mailboxes sent to {0}!" -f $EmailRecipient)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.