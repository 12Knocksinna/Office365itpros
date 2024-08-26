# Send-EmailExchangeHVE.PS1
# Example scriot to show how to send email using the Exchange High Volume Email (HVE) service in Office 365
# V1.0 22-Jul-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Send-EmailExchangeHVE.PS1

# Define the tenant identifier, subscription identifier, and the account used to send the email
$SubscriptionId = '35429342-a1a5-4427-9e2d-551840f2ad25'
$TenantId = 'b662313f-14fc-43a2-9a7a-d2e27f4f3478'
$HVEAccount = 'HVE1@office365itpros.com'

# Connect to Azure to access Key Vault
$AzConnection = Connect-AzAccount -Tenant $TenantId -Subscription $SubscriptionId
If (!($AzConnection)) {
    Write-Host "Failed to connect to Azure to retrieve HVE password"
    Break
}

# Check if we have a connection to Exchange Online and connect if necessary to fetch list of accepted domains
$Modules = Get-Module | Select-Object -ExpandProperty Name
If ('ExchangeOnlineManagement' -notin $Modules) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

[array]$Domains = (Get-AcceptedDomain).DomainName

# Retrieve the HVE password from Key Vault
$HVE01Password = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -name "HVE01Password" -AsPlainText
[securestring]$SecurePassword = ConvertTo-SecureString $HVE01Password -AsPlainText -Force
[pscredential]$HVECredentials = New-Object System.Management.Automation.PSCredential ($HVEAccount, $SecurePassword)

# Read in email addresses from a file
[array]$EmailAddresses = Import-CSV -Path "C:\Temp\StillToBuy.csv"
$EmailAddresses = $EmailAddresses | Sort-Object Email -Unique
If (!($EmailAddresses)) {
    Write-Host "Failed to read email addresses from file for HVE to process"
    Break
}       

# Build the HTML content for the email
$Content = "<p>Dear Subscriber:</p>" +
"<p>Over the past few weeks, we have sent out reminders that a discount is available for you to upgrade " +
"your subscription to cover Office 365 for IT Pros (2025 edition). It's now down to the wire and we will not offer a special discount to subscribers to the 2024 edition after September 1, 2024. " +
"The Office 365 for IT Pros (2025 edition) bundle includes the new 250-page <b>Automating Microsoft 365 with PowerShell eBook</b>. " +
"You do not have to buy the PowerShell book separately if you upgrade your subscription. Like the main book, we plan to update the PowerShell book monthly.</p>" + 
"<p>The current price to extend a subscription is `$29.95. This price will expire at midnight Pacific Time on September 1, 2024. " +
"After that, we will have a single `$39.95 discount available for any previous subscriber, no matter what edition they last bought. <p>"+
"<p>Please use the <a href='https://o365itpros.gumroad.com/l/O365IT/Subscriber2025'>link</a> to secure your upgrade </p>" +
"<p>I apologize if you receive this email after already upgrading. We have had a few problems with the Gumroad email system and were not " +
"confident that everyone received the upgrade link. Every year we get complaints from people who don't receive a notification " +
"about upgrade offers, so we tend to over-communicate now. If you have already upgraded, you might have received this email because " +
"you used a different email address to purchase your subscription. To avoid any future confusion, please send a note to support@gumroad.com to ask them " +
"to combine your purchases under your preferred email address.</p>" +
"<p>If you have any questions about subscriptions, please check <a href='https://office365itpros.com/office-365-for-it-pros-faq/'>our FAQ</a> " +
"or send email to <a href='mailto:o365itprosrenewals@office365itpros.com'>Customer Services</a>.</p>" +
"<p>Thank you for supporting Office 365 for IT Pros (2024 edition). We hope that you like Office 365 for IT Pros " +
"(2025 edition).</p>" +
"<p>Best Regards,</p><p>Tony</p><p><p>" +
"<p>P.S.: Please don't share the upgrade link with others. The link is for your use only. We have had a few instances where people shared"+
"the link with colleagues who used the link to subscribe. It is embarrassing all round when we have to reverse transactions and close " +
"accounts, but keeping the link secure is the only way to enable us to offer a low-cost subscription extension to our subscribers.</p>" +
"<p>------------------------------------------------------------------------------</p>" +
"<p>This email was sent using the Exchange Online High-Volume Email (HVE) service" 

[int]$i = 0
$Report = [System.Collections.Generic.List[Object]]::new()
[datetime]$StartTime = Get-Date

ForEach ($Recipient in $EmailAddresses.Email) {
    $i++
    Write-Host ("Sending HVE email to {0} ({1}/{2})" -f $Recipient, $i, $EmailAddresses.Count)
    $SendHVEMessageParams = @{
        From = $HVEAccount
        To = $Recipient
        Bcc = 'Customer.Services@office365itpros.com'
        Subject = "Offer to extend Office 365 for IT Pros subscription expires on September 1, 2024"
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
        $ErrorFlag = $True
    }
    If ($ErrorFlag) {
        $ReportLine = [PSCustomObject][Ordered]@{   
            Action      = 'Message Failed'
            Timestamp   = (Get-Date -format s) 
            Recipient   = $Recipient
        }
    } Else {
        $ReportLine = [PSCustomObject][Ordered]@{   
            Action      = 'Message Sent'
            Timestamp   = (Get-Date -format s) 
            Recipient   = $Recipient
        }
    }

    $Report.Add($ReportLine)  
    $ErrorFlag = $False

    If (($Recipient.Split('@')[1]) -notin $Domains) {
    # Pause only needed if sending email to external recipients
        Write-Host "Pausing for 5 seconds to avoid throttling after sending to external recipient..."
        Start-Sleep -Seconds 5
    }
}
[datetime]$EndTime = Get-Date

$SuccessfulSends = $Report | Where-Object {$_.Action -eq 'Message Sent'}
$Duration = $EndTime - $StartTime
Write-Host ("Started at {0} and finished at {1}" -f $StartTime, $EndTime)   
Write-Host ("{0} messages sent in {1} seconds" -f $SuccessfulSends.Count, $Duration.TotalSeconds)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.