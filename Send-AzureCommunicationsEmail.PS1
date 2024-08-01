# Send-AzureCommunicationsEmail.PS1
# An example script to show how to send email from Azure Communication Services using PowerShell
# V1.0 7-Jul-2024

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Send-AzureCommunicationsEmail.PS1

function Get-AccessToken {    
    # Get an access token for Azure Communication Services
    $params = @{
        Uri    = "https://login.microsoftonline.com/$($TenantId)/oauth2/v2.0/token"
        Method = "POST"
        Body   = @{
            client_id     = $AppId
            client_secret = $AppSecret
            grant_type    = "client_credentials"
            scope         = "https://communication.azure.com/.default"
        }
    }
    $token = Invoke-RestMethod @params
    return $token.access_token
}

Function Get-MessageRecipients { # Build a list of the recipients for a message
    [cmdletbinding()]
    Param(
        [array]$ListOfAddresses )
    $To = @()

    ForEach ($SMTPAddress in $ListOfAddresses) {
        $Recipient = @{
            Address = $SMTPAddress
            displayName = $SMTPAddress
        }    
        $To += $Recipient
    }   
    Return $To 
}

# Replace the identifiers with appropriate values for your tenant. Application (client) identifier, app secret, and tenant id
$AppId = 'a0789401-fc4f-499a-9bb2-8b0fd349798f'
$AppSecret = 'q7n8Q~xwBCGPz--2sVtDvQGgK4k-CsFtkKQxDb3~'
$TenantId = 'a662313f-14fc-43a2-9a7a-d2e27f4f3478'

# The sender's email address. This must be an address defined as a sender with Email Communication Services
# See https://learn.microsoft.com/azure/communication-services/quickstarts/email/add-multiple-senders for how to
# configure additional email addresses as senders
$SenderAddress = 'DoNotReply@office365itpros.com'

# Endpoint for the Azure Communication Services API
$CommunicationEndPoint = "office365itpros.unitedstates.communication.azure.com"

# Define the headers for the REST API call
$headers = @{
    "Content-Type"  = "application/json"
    "Authorization" = "Bearer $(Get-AccessToken)"
}

# Define the content of the email - handcrafted HTML here. This is an actual message that we sent to 
# subscribers of the Office 365 for IT Pros eBook
$Content = "<p>Dear Subscriber:</p>" +
"<p>As a subscriber to Office 365 for IT Pros (2024 edition), you are entitled to upgrade and extend your " +
"subscription to cover Office 365 for IT Pros (2025 edition). Time is running out for this offer, " +
"which closes on July 21." +
"<p>To avoid problems with messages sent using the Gumroad email system being tagged as spam or ending up in " +
"quarantine, we are sending this message using the Azure Email communication service (ECS). It is a relatively new " +
"service that is designed to handle mass mailings of the type we do when communicating with subscribers. " +
"It has been an interesting project to work out how to process email with ECS.</p>" +
"<p>If you have any questions, please check <a href='https://office365itpros.com/office-365-for-it-pros-faq/'>our FAQ</a> " +
"or send email to o365itprosrenewals@office365itpros.com.</p>" +
"<p>We hope that you like Office 365 for IT Pros (2025 edition).</p>" +
"<p>Best Regards,</p><p>Tony</p><p><p>" +
"<p>------------------------------------------------------------------------------</p>" +
"<p>This email was sent using Azure Email Communication Service" 

# Read in email addresses to process. This script uses a simple CSV file with a single column called Email
[array]$RecipientList = Import-CSV 'c:\temp\MessageRecipients.CSV'
# Remove any duplicates found in the set imported from the CSV
$RecipientList = $RecipientList | Sort-Object Email -Unique

# ECS has a limit of 100 recipients per message. If the list is longer, we'll only process the first 100
If ($RecipientList.count -gt 100) {
    Write-Host "Too many recipients to process in one go. Limiting the set to the first 100"
    $RecipientList = $RecipientList | Select-Object -First 100
}

$Report = [System.Collections.Generic.List[Object]]::new()
$MessageIds = [System.Collections.Generic.List[Object]]::new()

[int]$i = 0
Write-Host "Processing messages... "
ForEach ($Recipient in $RecipientList.Email) {
    # Construct the TO addresses for the message
    [array]$ToRecipientAddress = Get-MessageRecipients -ListOfAddresses $Recipient
    $i++
    Write-Host ("Sending email to {0} ({1}/{2})" -f $Recipient, $i, $RecipientList.count)

    # Build a hash table containing the settings for the message
    $Email = @{
         # The sender's email address
        senderAddress = $senderAddress
        # Create a unique identifier for this message
        headers = @{
            id = ("{0}-{1}" -f (Get-Date -format s), $ToRecipientAddress.address)
        }
       
        # The content of the email, including the subject and HTML body
        content = @{
            subject   = "Office 365 for IT Pros Subscription Update"
            html      = $Content
        }
        # The recipients of the email
        recipients = @{
            to = $ToRecipientAddress
            bcc = @(
               @{
                   address     = "o365itprosrenewals@office365itpros.com"
                   displayname = "Office 365 for IT Pros Support"
                }
            )
        }
        # The reply-to addresses for the email - doesn't have to be the same as the sender address
        ReplyTo = @(
            @{
                address     = "o365itprosrenewals@office365itpros.com"
                displayName = "Office 365 for IT Pros Support"
            }
        )
        userEngagementTrackingDisabled = $false
    }

    # Convert the email settings structure to JSON
    $EmailSettings = $Email | ConvertTo-Json -Depth 10
    $MailStatus = $null
    # Define the URI to post to when sending a message with ECS. 
    # The same URI is used for all messages. The body of the message dictates who receives the email
    $Uri = ("https://{0}/emails:send?api-version=2023-03-31" -f $CommunicationEndpoint)
    # Submit the message to the Email Communication service
    try {
        $MailStatus = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $EmailSettings -UseBasicParsing
        $ReportLine = [PSCustomObject][Ordered]@{   
            Id          = $MailStatus.id
            Timestamp   = (Get-Date -format s) 
            Recipient   = $Recipient
        }
        $MessageIds.Add($ReportLine)
    }
    catch {
        Write-Host ("Failed to send email to {0}" -f $Recipient)
        $ReportLine = [PSCustomObject][Ordered]@{   
            Timestamp   = (Get-Date -format s) 
            Recipient   = $Recipient
            Status      = $Mailstatus.status
            Id          = $Mailstatus.id
        }
        $Report.Add($ReportLine)
    }
    Start-Sleep -Seconds 2
    $Recipient = $null
}

Write-Host "Checking status of messages sent..."
# All messages are sent. Let's find out what happened
ForEach ($Message in $MessageIds) {
    $Uri = ("https://{0}/emails/operations/{1}?api-version=2023-03-31" -f $CommunicationEndpoint, $Message.id)
    Try {
        $MailStatus = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
        Write-Host ("Message sent to {0} returned status of {1}" -f $Message.Recipient, $MailStatus.status)
        $ReportLine = [PSCustomObject][Ordered]@{  
            Timestamp   = $Message.Timestamp
            Recipient   = $Message.Recipient
            Status      = $MailStatus.status
        }
        $Report.Add($ReportLine)
    } Catch {
        Write-Error $_.Exception.Message
    }
}
Clear-Host
Write-Host "Summary of email transmission session"
Write-Host "-------------------------------------"
Write-Host ""
Write-Host ("Messages sent via ECS endpoint {0}" -f $CommunicationEndPoint)
Write-Host ""
Write-Host ("Number of messages sent:              {0}" -f $Report.Count)   
Write-Host ("Number of messages sent successfully: {0}" -f ($Report | Where-Object { $_.Status -eq "Succeeded" }).Count)
Write-Host ("Number of messages failed:            {0}" -f ($Report | Where-Object { $_.Status -eq "Failed" }).Count)
Write-Host ("Number of messages pending:           {0}" -f ($Report | Where-Object { $_.Status -eq "Running" }).Count)
Write-Host ""
Write-Host "Message status"
Write-Host ""
$Report


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.