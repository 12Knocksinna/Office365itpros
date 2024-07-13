# Send-AzureCommunicationsEmail.PS1
# An example script to show how to send email from Azure Communication Services using PowerShell
# V1.0 7-Jul-2024

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Send-AzureCommunicationsEmail.PS1

function Get-AccessToken {    
    # Get an access token for the tenant using the defined credentials
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

    # Call the REST API and get the access token
    $token = Invoke-RestMethod @params
    return $token.access_token
}

Function Get-MessageToRecipients { # Build a list of the To recipients for a message
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
$AppId = 'e0789401-fc4f-499a-9bb2-8b0fd349798f'
$AppSecret = 'q7n8Q~xwBCGPz--2sVtDvQGgK4k-CsFtkKQxDb3~'
$TenantId = 'b662313f-14fc-43a2-9a7a-d2e27f4f3478'

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

# Define the content of the email - handcrafted HTML here
$Content = "<p>Dear Subscriber,</p>" +
"<p>You bought a full-price copy of the Office 365 for IT Pros eBook (2024 edition) during June 2024." +
"We promised to make a free subscription to the 2025 edition to you, but you have not yet accessed " +
"Gumroad.com to claim your free copy of <b><i>Office 365 for IT Pros (2025 edition)</i></b>, including " +
"the new <b><i>Automating Microsoft 365 with PowerShell</i></b> eBook.</span></p>" +
"
<p><span lang=EN-US>To claim your free subscription, please " +
'<a href="https://o365itpros.gumroad.com/l/O365IT/June2024Buyer">click this link</a>. ' +
"After receiving the emailed receipt, use the View Content button in the receipt to access the EPUB " +
"and PDF files for the two books. If you do not get the receipt, please check your Junk Email folder.</span></p>" +

"<p><span lang=EN-US>If you have any other problems, please contact us at o365itprosrenewals@office365itpros.com " +
"and we will do our best to resolve the issue.</span></p>" +
"<p><span lang=EN-US>Tony</span></p>"

[array]$RecipientList = Import-CSV 'c:\temp\UserData.CSV'
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Recipient in $RecipientList.Email) {
    # Construct the TO addresses for the message
    $RecipientAddresses = $Recipient
    [array]$ToRecipientAddress = Get-MessageToRecipients -ListOfAddresses $RecipientAddresses
    Write-Host ("Sending email to {0}" -f $Recipient)

    # Build a hash table containing the settings for the message
    $Email = @{
         # The sender's email address
        senderAddress = $senderAddress
        # Email headers, including a unqiue ID.
        headers = @{
            id = ("{0}-{1}" -f (Get-Date -format s), $RecipientAddress.displayname)
        }
       
        # The content of the email, including the subject and HTML body
        content = @{
            subject   = "Office 365 for IT Pros Email Subscription Update"
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
    $Mailcheck = $null
    # URI to send a message... The same URI is used for all messages. The body of the message dictates who receives the email
    $Uri = ("https://{0}/emails:send?api-version=2023-03-31" -f $CommunicationEndpoint)
    # Submit the message to the Email Communication service
    try {
        $Mailstatus = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $EmailSettings -UseBasicParsing
    }
    catch {
        Write-Host ("Failed to send email to {0}" -f $Recipient)
        $ReportLine = [PSCustomObject][Ordered]@{   
            Timestamp   = (Get-Date -format s) 
            Recipient   = $Recipient
            Status      = $Mailcheck.status
        }
        $Report.Add($ReportLine)
    }
    # Check the status of a sent message.
    Start-Sleep -Seconds 4
    $Uri = ("https://{0}/emails/operations/{1}?api-version=2023-03-31" -f $CommunicationEndpoint, $Mailstatus.id)
    Try {
        $MailCheck = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
        Write-Host ("Message sent to {0} returned status of {1}" -f $Recipient, $Mailcheck.status)
        $ReportLine = [PSCustomObject][Ordered]@{  
            Timestamp   = (Get-Date -format s) 
            Recipient   = $Recipient
            Status      = $Mailcheck.status
        }
        $Report.Add($ReportLine)
    } Catch {
        Write-Error $_.Exception.Message
    }
    $Recipient = $null
}

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