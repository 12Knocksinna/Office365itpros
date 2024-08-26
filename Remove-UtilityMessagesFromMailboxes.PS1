# Remove-UtilityMessagesFromMailboxes.PS1
# Example of how to use the Microsoft Graph PowerShell SDK cmdlets to remove "utility" messages from user mailboxes.
# Utility messages are those sent by workloads like Teams and SharePoint Online to mailboxes to inform users of events.
# V1.0 26-Aug-2024
# GitHub link:

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ('ExchangeOnlineManagement' -notin $Modules) {
    Write-Output "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}   

Connect-MgGraph -NoWelcome -Scopes User.Read, Mail.ReadWrite, Mail.Send
# List the utility email addresses to search for and remove
$UM0 = "notifications@communityhub.microsoft.com"
$UM1 = "messages-noreply@linkedin.com"
$UM2 = "noreply@microsoft.com"
$UM3 = "mssecurity-noreply@microsoft.com"
$UM4 = "noreply@yammer.com"
$UM5 = "microsoft@powerapps.com"
$UM6 = "o365mc@microsoft.com"
$UM7 = "jobs-listings@linkedin.com"

# Set the start and end dates for the search
[datetime]$StartDate = (Get-Date).AddDays(-730)
[string]$StartDate = Get-Date $StartDate -Format "yyyy-MM-ddTHH:mm:ssZ"
[string]$EndDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"

# Find all mailboxes in the tenant
Write-Host "Finding mailboxes..." -ForegroundColor Cyan
[array]$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Sort-Object DisplayName
If ($Mbx) {
    Write-Output ("{0} mailboxes found" -f $Mbx.Count)
} Else {
    Write-Output "No mailboxes found"
    Break
}

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
    Write-Host ("Processing mailbox {0}" -f $M.DisplayName) -ForegroundColor Yellow

    [array]$Messages = Get-MgUserMailFolderMessage -UserId $M.externaldirectoryobjectid -MailFolderId 'Inbox' `
    -Filter "(ReceivedDateTime ge $StartDate and ReceivedDateTime le $EndDate) `
            and (sender/emailAddress/address eq '$UM0' or sender/emailAddress/address eq '$UM1' or sender/emailAddress/address eq '$UM2' `
            or sender/emailAddress/address eq '$UM3' or sender/emailAddress/address eq '$UM4' or sender/emailAddress/address eq '$UM5' `
            or sender/emailAddress/address eq '$UM6' or sender/emailAddress/address eq '$UM7')" `
            -Property Id, Subject, Sender, SentDateTime -All -PageSize 500

    If ($Messages) {
        ForEach ($Message in $Messages) {
            Try {
                Remove-MgUserMessage -UserId $M.ExternalDirectoryObjectId -MessageId $Message.Id -ErrorAction Continue
                $ReportLine = [PSCustomObject][Ordered]@{   
                    Action      = 'Message Deleted'
                    Mailbox     = $M.DisplayName
                    Timestamp   = (Get-Date -format s) 
                    Subject     = $Message.Subject
                    SenderEmail = $Message.Sender.EmailAddress.Address
                    SenderName  = $Message.Sender.EmailAddress.Name
                    Sent        = $Message.SentDateTime
                }
            } Catch {
                Write-Host ("Failed to remove message {0} from mailbox {1} with error {2}" -f $Message.Id, $M.DisplayName, $_.Exception.Message) -ForegroundColor Red
                $ReportLine = [PSCustomObject][Ordered]@{   
                    Action      = 'Message Deletion Failed'
                    Mailbox     = $M.DisplayName
                    Timestamp   = (Get-Date -format s) 
                    Subject     = $Message.Subject
                    SenderEmail = $Message.Sender.EmailAddress.Address
                    SenderName  = $Message.Sender.EmailAddress.Name
                    Sent        = $Message.SentDateTime
                }
            }
            $Report.Add($ReportLine)
        }
    } Else {
        Write-Host ("No utility messages found in mailbox {0}" -f $M.DisplayName) -ForegroundColor Green
    }

}

$Report | Out-GridView -Title "Utility Messages Removed from Mailboxes" 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.

[array]$Messages = Get-MgUserMailFolderMessage -UserId $M.externaldirectoryobjectid -MailFolderId 'Inbox' `
-Filter "(ReceivedDateTime ge $StartDate and ReceivedDateTime le $EndDate) `
        and (sender/emailAddress/address eq '$UM0' or sender/emailAddress/address eq '$UM1' or sender/emailAddress/address eq '$UM2' `
        or sender/emailAddress/address eq '$UM3' or sender/emailAddress/address eq '$UM4' or sender/emailAddress/address eq '$UM5' `
        or sender/emailAddress/address eq '$UM6' or sender/emailAddress/address eq '$UM7')" `
        -Property Id, Subject, Sender, SentDateTime -All -PageSize 500
