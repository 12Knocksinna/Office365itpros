# Remove-UtilityMessagesFromMailboxes.PS1
# Example of how to use the Microsoft Graph PowerShell SDK cmdlets to remove "utility" messages from user mailboxes.
# Utility messages are those sent by workloads like Teams and SharePoint Online to mailboxes to inform users of events.
# V1.0 26-Aug-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Remove-UtilityMessagesFromMailboxes.PS1

# Connect to the Microsoft Graph
$AppThumbprint = "F79286DB88C21491110109A0222348FACF694CBD"
$AppId = "d58578ac-7cb4-4b5a-a296-f19218a03f11"
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3478"

# The app needs the following application permissions:
# User.Read.All: Read user account details
# Mail.ReadWrite: Read and write mail in all mailboxes
# Mail.Send: Send mail as a user

Connect-MgGraph -NoWelcome -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $AppThumbprint

# List the email addresses of the notification messages to search for and remove
$UM0 = "notifications@communityhub.microsoft.com"
$UM1 = "messages-noreply@linkedin.com"
$UM2 = "noreply@microsoft.com"
$UM3 = "notifications-noreply@linkedin.com"
$UM4 = "noreply@yammer.com"
$UM5 = "microsoft@powerapps.com"
$UM6 = "o365mc@microsoft.com"
$UM7 = "jobs-listings@linkedin.com"
$UM8 = "notifications@yammer.com"
$UM9 = "noreply@amazon.com"
$UM10 = "notify@mstechcommunity.microsoft.com"
$UM11 = "notifications@owler.com"

# Set the start and end dates for the search - starting five years ago and ending 30 days ago
[datetime]$StartDate = (Get-Date).AddDays(-1825)
[string]$StartDate = Get-Date $StartDate -Format "yyyy-MM-ddTHH:mm:ssZ"
[datetime]$EndDate = (Get-Date).AddDays(-30)
[string]$EndDate = Get-Date $EndDate -Format "yyyy-MM-ddTHH:mm:ssZ"

# Define the service plan IDs for Exchange Online (Plan 1) and Exchange Online (Plan 2)
$ExoServicePlan1 = "9aaf7827-d63c-4b61-89c3-182f06f82e5c"
$ExoServicePlan2 = "efb87545-963c-4e0d-99df-69c6916d9eb0" 

# Find users assigned a license that includes the Exchange Online (Plan 1) or Exchange Online (Plan 2) service plans.
# The check also looks for users with the Exchange Online service plan enabled.
Write-Host "Finding mailboxes..." -ForegroundColor Cyan
[array]$Users = Get-MgUser -Filter "assignedPlans/any(c:c/servicePlanId eq $ExoServicePlan1 and capabilityStatus eq 'Enabled') `
  or assignedPlans/any(c:c/servicePlanId eq $ExoServicePlan2 and capabilityStatus eq 'Enabled')" `
    -ConsistencyLevel eventual -CountVariable Test -All -PageSize 999 -Sort ('displayname') `
    -Property Id, displayName, userprincipalName, assignedLicenses, assignedPlans, department, country, CreatedDateTime

If ($Users) {
    Write-Output ("{0} users licensed for Exchange Online found" -f $Users.Count)
} Else {
    Write-Output "No mailboxes found"
    Break
}

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    Write-Host ("Processing mailbox {0}" -f $User.DisplayName) -ForegroundColor Yellow
    [int]$i = 0

    [array]$Messages = Get-MgUserMailFolderMessage -UserId $User.Id -MailFolderId 'Inbox' `
    -Filter "(ReceivedDateTime ge $StartDate and ReceivedDateTime le $EndDate) `
        and (sender/emailAddress/address eq '$UM0' or sender/emailAddress/address eq '$UM1' or sender/emailAddress/address eq '$UM2' `
        or sender/emailAddress/address eq '$UM3' or sender/emailAddress/address eq '$UM4' or sender/emailAddress/address eq '$UM5' `
        or sender/emailAddress/address eq '$UM6' or sender/emailAddress/address eq '$UM7' or sender/emailAddress/address eq '$UM8' `
        or sender/emailAddress/address eq '$UM9' or sender/emailAddress/address eq '$UM10' or sender/emailAddress/address eq '$UM11')" `
        -Property Id, Subject, Sender, SentDateTime -All -PageSize 999

    If ($Messages) {
        # If messages are found, delete them!
        ForEach ($Message in $Messages) {
            Try {
                Remove-MgUserMessage -UserId $User.Id -MessageId $Message.Id -ErrorAction Continue
                $i++
                $ReportLine = [PSCustomObject][Ordered]@{ 
                    Timestamp   = (Get-Date -format s)   
                    Action      = 'Message Deleted'
                    Mailbox     = $User.DisplayName
                    SMTPAddress = $User.UserPrincipalName
                    Subject     = $Message.Subject
                    SenderEmail = $Message.Sender.EmailAddress.Address
                    SenderName  = $Message.Sender.EmailAddress.Name
                    Sent        = $Message.SentDateTime
                }
            } Catch {
                Write-Host ("Failed to remove message {0} from mailbox {1} with error {2}" -f $Message.Id, $User.DisplayName, $_.Exception.Message) -ForegroundColor Red
                $ReportLine = [PSCustomObject][Ordered]@{  
                    Timestamp   = (Get-Date -format s)  
                    Action      = 'Message Deletion Failed'
                    Mailbox     = $User.DisplayName
                    SMTPAddress = $User.UserPrincipalName
                    Subject     = $Message.Subject
                    SenderEmail = $Message.Sender.EmailAddress.Address
                    SenderName  = $Message.Sender.EmailAddress.Name
                    Sent        = $Message.SentDateTime
                }
            }
            $Report.Add($ReportLine)
        }
    } Else {
        Write-Host ("No notification messages found in mailbox {0}" -f $User.DisplayName) -ForegroundColor Green
    }
    If ($i -gt 0) {
        Write-Host ("{0} notification messages removed from mailbox {1}" -f $i, $User.DisplayName) -ForegroundColor Green
    }   
}

$Report | Out-GridView -Title "Utility Messages Removed from Mailboxes" 

# Send email to the mailbox owners
[array]$IndividualMailboxes = $Report | Sort-Object SMTPAddress -Unique | Select-Object -ExpandProperty SMTPAddress
$MsgSubject = "Notification Messages Removed from Your Mailbox"
$MsgFrom = 'Azure.Management.Account@office365itpros.com'

$HTMLHeader = "<body><h1>Messages removed by the ever-loyal Utility Message Removal Service</h1>
     <p><strong>Generated:</strong> $(Get-Date -Format g)</p>  
     <h2><u>Messages removed from your Inbox</u></h2></p>
     <p>The background agent removes messages from selected senders from your mailbox to avoid a build-up of messages with little value.</p>
     <p>If necessary, you can recover these messages by going to the Deleted Items folder and using the <i>Recover items recently removed from this folder</i> option.</p>"
$HTMLFooter = 
    "<p>------------------------------------------------------------------------------</p>" +
    "<p>This email was sent by the Utility Message Removal service" +
    "<p>------------------------------------------------------------------------------</p></body>"

ForEach ($Address in $IndividualMailboxes) {
  
    # Create a HTML fragment containing details of the messages removed from the mailbox
    $HTMLReport = $Report | Where-Object {$_.SMTPAddress -eq $Address -and $_.Action -eq 'Message Deleted'} | `
        Select-Object Action, Subject, SenderEmail, SenderName, Sent | Sort-Object {$_.Sent -as [datetime]} -Descending
    Write-Host ("Sending notification about {0} removed messages to {1}" -f $HTMLReport.count, $Address)
    $HTMLReport = $HTMLReport | ConvertTo-Html -Fragment
    # Assemble the pieces into a complete HTML message
    $HTMLBody = $HTMLHeader + $HTMLReport + $HTMLFooter
    # Add the recipient using the mailbox's primary SMTP address
    $EmailAddress  = @{address = $Address} 
    $EmailRecipient = @{EmailAddress = $EmailAddress}  
    # Construct the message body
    $MessageBody = @{
        content = "$($HtmlBody)"
        ContentType = 'html'
    }
    # Create a draft message in the sending mailbox
    $NewMessage = New-MgUserMessage -UserId $MsgFrom -Body $MessageBody -ToRecipients $EmailRecipient -Subject $MsgSubject 
    # Send the message
    Send-MgUserMessage -UserId $MsgFrom -MessageId $NewMessage.Id  

}

Write-Host ("Utility message cleanup complete. Details of {0} notification messages sent to {1} users" -f $Report.Count, $IndividualMailboxes.Count) -ForegroundColor Cyan

# Generate the report in either Excel worksheet or CSV format, depending on if the ImportExcel module is available
If (Get-Module ImportExcel -ListAvailable) {
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Removed Notification Messages Report.xlsx"
    $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Deleted Notification Messages" -Title ("Removed Notification Messages {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "Microsoft365LicensingReport" 
    Write-Host ("Report saved to {0}" -f $ExcelOutputFile) -ForegroundColor Green
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Removed Notification Messages Report.CSV"
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
    Write-Host ("Report saved to {0}" -f $CSVOutputFile) -ForegroundColor Green
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.