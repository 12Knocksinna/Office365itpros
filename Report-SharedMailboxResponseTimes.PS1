# Report-SharedMailboxResponseTimes.PS1
# This script demonstrates the principles behind generation of a report of response times for shared mailboxes in an Exchange Online environment.

# V1.0 18-Feb-2026
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-SharedMailboxResponseTimes.PS1

# Example use .\Report-SharedMailboxResponseTimes.PS1 -SharedMailboxes $SharedMailboxesToProcess -EmailAddressesForAgents $EmailAddressesForAgents 
# $SharedMailboxesToProcess is an array of shared mailbox email addresses to process and $EmailAddressesForAgents is an array of email addresses 
# for agents who respond to messages in the shared mailboxes. If you do not provide values for these parameters, the script use default values 
# defined in the code.

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$false)]
    [array]$SharedMailboxes,
    
    [Parameter(Mandatory=$false)]
    [array]$EmailAddressesForAgents 
    # These email addresses are used to identify agents who respond to messages. They include the email addresses for the
    # shared mailboxes being processed (as agents may respond from the shared mailbox address) and the email addresses of the agents themselves
    # if the agents use the Send On Behalf Of permission to send email for the mailbox. Adjust as necessary for your environment.
)

# Uses the delegated Mail.Read.Shared permission
Connect-MgGraph -Scopes "Mail.Read.Shared" -NoWelcome

If (-not $SharedMailboxes) {
    # If no mailboxes were provided, use a default list of shared mailboxes to process. Adjust as necessary for your environment.
    $SharedMailboxes = @("contact@office365itpros.com")
}

# Go back 6 months for this check. Adjust as necessary.
[datetime]$StartDate = (Get-Date).AddDays(-180)
[string]$StartDate = Get-Date $StartDate -Format "yyyy-MM-ddTHH:mm:ssZ"

Write-Host ("Processing shared mailboxes: {0}" -f ($SharedMailboxes -join ', ')) -ForegroundColor Green
Write-Host ("Identifying agents by email addresses: {0}" -f ($EmailAddressesForAgents -join ', ')) -ForegroundColor Green
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Mailbox in $SharedMailboxes) {
    Write-Host "Processing mailbox: $Mailbox" -ForegroundColor Cyan
    # Make sure that the current mailbox is included in the list of email addresses for agents (as agents may respond from the shared mailbox address) and use lower cased addresses
    $EmailAddressesForAgents = $EmailAddressesForAgents + $Mailbox
    $EmailAddressesForAgents = $EmailAddressesForAgents | ForEach-Object { $_.ToLower() } | Sort-Object -Unique
    Try {
        [array]$Items = Get-MgUserMailFolderMessage -MailFolderId 'Inbox' -UserId $Mailbox -Filter "receivedDateTime ge $StartDate" -All -PageSize 500 `
            -Select ConversationId,Id,ReceivedDateTime,Subject,From -ErrorAction Stop
    } Catch {
        Write-Host ("Error accessing mailbox {0}" -f $Mailbox) -ForegroundColor Red
        Continue
    }
    # Find unique conversations in the Inbox
    [array]$Conversations = $Items | Sort-Object ConversationId -Unique | Sort-Object {$_.ReceivedDateTime -as [datetime]} -Descending
    Write-Host ("Found {0} unique conversations in mailbox {1). Checking each conversation to extract response data" -f $Conversations.Count, $Mailbox) -ForegroundColor Green
   
    ForEach ($Conversation in $Conversations) {
        # Get all messages in the conversation
        $ConversationId = $Conversation.ConversationId
        [array]$ThreadItems = Get-MgUserMessage -UserId $Mailbox -Filter "conversationID eq '$ConversationId'" | Sort-Object {$_.ReceivedDateTime -as [datetime]}
        # Process the thread if it has more than one message (indicating at least one response)
        If ($ThreadItems.Count -gt 1) {
            $FirstResponseMessage = $ThreadItems | Where-Object { $_.From.EmailAddress.Address.ToLower() -in $EmailAddressesForAgents } | Select-Object -First 1
            If ($null -eq $FirstResponseMessage) {
                # No response from an agent in this thread, skip to the next conversation
                Continue
            }
            $OriginalMessage = $ThreadItems | Where-Object {$_.Subject.Substring(0,3) -ne "RE:" -or $_.Subject.Substring(0,3) -ne "Re:"} | Select-Object -First 1
            If ($null -eq $OriginalMessage) {
                # No original message found so we can't measure the response time for the conversation. Skip to the next conversation
                Continue
            }
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

            # Find the agents who participated in the conversation by reading the set of from addresses used in messages in the thread 
            # and dropping the original sender's address (which is not an agent). Group by address to get a count of messages from each agent and a list of names associated with each address.
            [array]$AgentsInConversation = $Threaditems.sender.emailaddress | Where-Object {$_.Address -ne $OriginalMessage.From.EmailAddress.Address} |
                Group-Object -Property Address |
                ForEach-Object {
                    [pscustomobject]@{
                        Address = $_.Name
                        Name    = ($_.Group.Name | Sort-Object -Unique) -join '; '
                        Count   = $_.Count
                    }
                } |
            Sort-Object Address
            # For debugging...
            # Write-Host "Conversation Id: $($Conversation.ConversationId)" -ForegroundColor Yellow
            # Add to report
            $ReportLine = [PSCustomObject][Ordered]@{ 
                Mailbox                             = $Mailbox
                Subject                             = $OriginalMessage.Subject
                Sender                              = $OriginalMessage.From.EmailAddress.Address
                Agents                              = ($AgentsInConversation.Name -join ', ')
                'Initial message received at'       = Get-Date $OriginalMessage.ReceivedDateTime -format 'dd-MMM-yyyy HH:mm'
                'First response at'                 = Get-Date $FirstResponseMessage.ReceivedDateTime -format 'dd-MMM-yyyy HH:mm'
                'Final interaction'                 = Get-Date $FinalMessage.ReceivedDateTime -format 'dd-MMM-yyyy HH:mm'
                'Time to initial response'          = $ResponseTimeFormatted
                'Time to final response'            = $ResponseTimeToFinalFormatted
                ConversationId                      = $Conversation.ConversationId
                InitialResponseMinutes              = $FirstResponseTimeMinutes
                FinalResponseMinutes                = $FinalResponseTimeMinutes
                AgentsInvolved                      = $AgentsInConversation
            }
            $Report.Add($ReportLine)
        }
    }
}

Write-Host ""
Write-Host "Report generated. Displaying results" -ForegroundColor Cyan
Write-Host ""
ForEach ($Mailbox in $SharedMailboxes) {
    Write-Host ("Mailbox: {0}" -f $Mailbox) -ForegroundColor Green
    $ReportForMailbox = $Report | Where-Object { $_.Mailbox -eq $Mailbox }
    $ReportForMailbox | Select-Object 'Initial message received at', Sender, Subject | Format-Table -Wrap -AutoSize
    Write-Host ""
    Write-Host "Total messages processed: $($ReportForMailbox.Count)"

    # Calculate and format average response times
    $AvgInitialMinutes = ($ReportForMailbox | Measure-Object InitialResponseMinutes -Average).Average
    $AvgFinalMinutes = ($ReportForMailbox | Measure-Object FinalResponseMinutes -Average).Average

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
    Write-Host ""
}

Write-Host "All done. Reporting information..."

# Generate the report in either Excel worksheet or CSV format, depending on if the ImportExcel module is available
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Shared Mailboxes Responses Report.xlsx"
    $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Shared Mailboxes Responses" -Title ("Shared Mailboxes Responses Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "SharedMailboxesResponsesReport" -AutoSize -AutoFilter 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Shared Mailboxes Responses Report.CSV"
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}

If ($ExcelGenerated -eq $true) {
    Write-Host ("Shared Mailboxes Responses report available in Excel workbook {0}" -f $ExcelOutputFile)
} Else {
    Write-Host ("Shared Mailboxes Responses report available in CSV file {0}" -f $CSVOutputFile)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.