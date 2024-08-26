# Remove-UtilityMessagesFromMailboxes.PS1
# Example of how to use the Microsoft Graph PowerShell SDK cmdlets to remove "utility" messages from user mailboxes.
# Utility messages are those sent by workloads like Teams and SharePoint Online to mailboxes to inform users of events.
# V1.0 26-Aug-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Remove-UtilityMessagesFromMailboxes.PS1

# Connect to the Microsoft Graph
$AppThumbprint = "F79286DB88C21491110109A0222348FACF694CBD"
$AppId = "d58578ac-7cb4-4b5a-a296-f19218a03f11"
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3478"

Connect-MgGraph -NoWelcome -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $AppThumbprint
# List the utility email addresses to search for and remove
$UM0 = "notifications@communityhub.microsoft.com"
$UM1 = "messages-noreply@linkedin.com"
$UM2 = "noreply@microsoft.com"
$UM3 = "mssecurity-noreply@microsoft.com"
$UM4 = "noreply@yammer.com"
$UM5 = "microsoft@powerapps.com"
$UM6 = "o365mc@microsoft.com"
$UM7 = "jobs-listings@linkedin.com"
$UM8 = "notifications@yammer.com"

# Set the start and end dates for the search
[datetime]$StartDate = (Get-Date).AddDays(-730)
[string]$StartDate = Get-Date $StartDate -Format "yyyy-MM-ddTHH:mm:ssZ"
[string]$EndDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"

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

    [array]$Messages = Get-MgUserMailFolderMessage -UserId $User.Id -MailFolderId 'Inbox' `
    -Filter "(ReceivedDateTime ge $StartDate and ReceivedDateTime le $EndDate) `
            and (sender/emailAddress/address eq '$UM0' or sender/emailAddress/address eq '$UM1' or sender/emailAddress/address eq '$UM2' `
            or sender/emailAddress/address eq '$UM3' or sender/emailAddress/address eq '$UM4' or sender/emailAddress/address eq '$UM5' `
            or sender/emailAddress/address eq '$UM6' or sender/emailAddress/address eq '$UM7' or sender/emailAddress/address eq '$UM8')" `
            -Property Id, Subject, Sender, SentDateTime -All -PageSize 999

    If ($Messages) {
        ForEach ($Message in $Messages) {
            Try {
               # Remove-MgUserMessage -UserId $User.Id -MessageId $Message.Id -ErrorAction Continue
                $ReportLine = [PSCustomObject][Ordered]@{   
                    Action      = 'Message Deleted'
                    Mailbox     = $User.DisplayName
                    Timestamp   = (Get-Date -format s) 
                    Subject     = $Message.Subject
                    SenderEmail = $Message.Sender.EmailAddress.Address
                    SenderName  = $Message.Sender.EmailAddress.Name
                    Sent        = $Message.SentDateTime
                }
            } Catch {
                Write-Host ("Failed to remove message {0} from mailbox {1} with error {2}" -f $Message.Id, $User.DisplayName, $_.Exception.Message) -ForegroundColor Red
                $ReportLine = [PSCustomObject][Ordered]@{   
                    Action      = 'Message Deletion Failed'
                    Mailbox     = $User.DisplayName
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
        Write-Host ("No utility messages found in mailbox {0}" -f $User.DisplayName) -ForegroundColor Green
    }

}

$Report | Out-GridView -Title "Utility Messages Removed from Mailboxes" 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.