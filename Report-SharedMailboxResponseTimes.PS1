# Report-SharedMailboxResponseTimes.PS1
# This script demonstrates the principles behind generation of a report of response times for shared mailboxes in an Exchange Online environment.

# V1.0 18-Feb-2026
# GitHub link:

# Uses the delegated Mail.Read.Shared permission
Connect-MgGraph -Scopes "Mail.Read.Shared" -NoWelcome

# Shared mailboxes to monitor - the signed in user must have full access to these mailboxes
[array]$SharedMailboxes = "contact@office365itpros.com"
# Email addresses that might generate responses to inbound messages - this includes the shared mailboxes themselves and any other addresses that might be used to respond to messages in the shared mailbox
[array]$EmailAddressesForResponses = $SharedMailboxes, "Tony.Redmond@redmondassociates.org", "tony.redmond@office365itpros.com" | ForEach-Object { $_.ToLower() }   

# Go back 3 months for this check. Adjust as necessary
[datetime]$StartDate = (Get-Date).AddDays(-180)
[string]$StartDate = Get-Date $StartDate -Format "yyyy-MM-ddTHH:mm:ssZ"

ForEach ($Mailbox in $SharedMailboxes) {
    Write-Host "Processing mailbox: $Mailbox" -ForegroundColor Cyan
    Try {
        [array]$Items = Get-MgUserMailFolderMessage -MailFolderId 'Inbox' -UserId $Mailbox -Filter "receivedDateTime ge $StartDate" -All -PageSize 500 `
            -Select ConversationId,Id,ReceivedDateTime,Subject,From -ErrorAction Stop
    } Catch {
        Write-Host ("Error accessing mailbox {0}" -f $Mailbox) -ForegroundColor Red
        Continue
    }
    # Find unique conversations in the Inbox
    [array]$Conversations = $Items | Sort-Object ConversationId -Unique | Sort-Object {$_.ReceivedDateTime -as [datetime]} -Descending
    Write-Host ("Found {0} unique conversations in mailbox {1}" -f $Conversations.Count, $Mailbox) -ForegroundColor Green

    $Report = [System.Collections.Generic.List[Object]]::new()
    ForEach ($Conversation in $Conversations) {
        # Get all messages in the conversation
        $ConversationId = $Conversation.ConversationId
        [array]$ThreadItems = Get-MgUserMessage -UserId $Mailbox -Filter "conversationID eq '$ConversationId'" | Sort-Object {$_.ReceivedDateTime -as [datetime]}
        # Process the thread if it has more than one message (indicating at least one response)
        If ($ThreadItems.Count -gt 1) {
            $FirstResponseMessage = $ThreadItems | Where-Object { $_.From.EmailAddress.Address.ToLower() -in $EmailAddressesForResponses } | Select-Object -First 1
            $OriginalMessage = $ThreadItems[0]
            $FinalMessage = $ThreadItems[-1]
            $FirstResponseTimeMinutes = ($FirstResponseMessage.ReceivedDateTime - $OriginalMessage.ReceivedDateTime).TotalMinutes
            $FinalResponseTimeMinutes = ($FinalMessage.ReceivedDateTime - $OriginalMessage.ReceivedDateTime).TotalMinutes
            # Format response time as days, hours, and minutes
            $Days = [math]::Floor($FirstResponseTimeMinutes / 1440)
            $Hours = [math]::Floor(($FirstResponseTimeMinutes % 1440) / 60)
            $Minutes = [math]::Floor($FirstResponseTimeMinutes % 60)
            $ResponseTimeFormatted = "{0}d {1}h {2}m" -f $Days, $Hours, $Minutes
            $FinalDays = [math]::Floor($FinalResponseTimeMinutes / 1440)
            $FinalHours = [math]::Floor(($FinalResponseTimeMinutes % 1440) / 60)
            $FinalMinutes = [math]::Floor($FinalResponseTimeMinutes % 60)
            $ResponseTimeToFinalFormatted = "{0}d {1}h {2}m" -f $FinalDays, $FinalHours, $FinalMinutes
            # Add to report
            $ReportLine = [PSCustomObject][Ordered]@{ 
                Mailbox                             = $Mailbox
                Subject                             = $OriginalMessage.Subject
                Sender                              = $OriginalMessage.From.EmailAddress.Address
                'Initial message received at'       = Get-Date $OriginalMessage.ReceivedDateTime -format 'dd-MMM-yyyy HH:mm'
                'First response at'                 = Get-Date $FirstResponseMessage.ReceivedDateTime -format 'dd-MMM-yyyy HH:mm'
                'Final interaction'                 = Get-Date $FinalMessage.ReceivedDateTime -format 'dd-MMM-yyyy HH:mm'
                'Time to initial response'          = $ResponseTimeFormatted
                'Time to final response'            = $ResponseTimeToFinalFormatted
                ConversationId                      = $Conversation.ConversationId
                InitialResponseMinutes              = $FirstResponseTimeMinutes
                FinalResponseMinutes                = $FinalResponseTimeMinutes
            }
            $Report.Add($ReportLine)
        }
    }
}

Write-Host ""
Write-Host "Report generated. Displaying results:" -ForegroundColor Cyan
$Report | Select-Object 'Initial message received at', Sender, Subject | Format-Table -Wrap -AutoSize
Write-Host ""
Write-Host "Total messages processed: $($Report.Count)"

# Calculate and format average response times
$AvgInitialMinutes = ($Report | Measure-Object InitialResponseMinutes -Average).Average
$AvgFinalMinutes = ($Report | Measure-Object FinalResponseMinutes -Average).Average

$AvgInitialDays = [math]::Floor($AvgInitialMinutes / 1440)
$AvgInitialHours = [math]::Floor(($AvgInitialMinutes % 1440) / 60)
$AvgInitialSecs = [math]::Floor($AvgInitialMinutes % 60)
$AvgInitialFormatted = "{0}d {1}h {2}m" -f $AvgInitialDays, $AvgInitialHours, $AvgInitialSecs

$AvgFinalDays = [math]::Floor($AvgFinalMinutes / 1440)
$AvgFinalHours = [math]::Floor(($AvgFinalMinutes % 1440) / 60)
$AvgFinalSecs = [math]::Floor($AvgFinalMinutes % 60)
$AvgFinalFormatted = "{0}d {1}h {2}m" -f $AvgFinalDays, $AvgFinalHours, $AvgFinalSecs

Write-Host ("Average time to initial response: {0}" -f $AvgInitialFormatted)
Write-Host ("Average time to final response: {0}" -f $AvgFinalFormatted)


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.