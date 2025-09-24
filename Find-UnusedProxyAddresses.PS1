
# Check if we can run an Exchange Online cmdlet. If we can, go on, else connect to Exchange Online
If ($Null -eq (Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$False -ErrorAction Stop
}

# Message trace date is kept for a maximum of 10 days
$StartDate = (Get-Date).AddDays(-10)
$EndDate = (Get-Date)

Write-Host ("Message trace data will be analyzed between {0} and {1}" -f (Get-Date $StartDate -format 'dd-MMM-yyyy HH:mm'), (Get-Date $EndDate -format 'dd-MMM-yyyy HH:mm'))
Write-Host "Fetching message trace data to analyze"

[array]$Messages = $Null
[int]$BatchSizeForMessages = 2000
    # original code [array]$MessagePage = Get-MessageTrace -StartDate $StartDate -EndDate $EndDate -PageSize 1000 -Page $i -Status "Delivered"
Try {
    # The warning action is suppressed here because we don't want to see warnings when more data is available
    [array]$MessagePage = Get-MessageTraceV2 -StartDate $StartDate -EndDate $EndDate `
        	-ResultSize $BatchSizeForMessages -Status "Delivered" -ErrorAction Stop -WarningAction SilentlyContinue
    $Messages += $MessagePage
} Catch {
    Write-Host ("Error fetching message trace data: {0}" -f $_.Exception.Message)
    Break
}
If ($MessagePage.count -eq $BatchSizeForMessages) {
    Do {
        Write-Host ("Fetched {0} messages so far" -f $Messages.count)
        $LastMessageFetched = $MessagePage[-1]
        $LastMessageFetchedDate = $LastMessageFetched.Received.ToString("O")
        $LastMessageFetchedRecipient = $LastMessageFetched.RecipientAddress
        # Fetch the next page of messages
        [array]$MessagePage = Get-MessageTraceV2 -StartDate $StartDate -EndDate $LastMessageFetchedDate `
            -StartingRecipientAddress $LastMessageFetchedRecipient -ResultSize $BatchSizeForMessages -Status "Delivered" -ErrorAction Stop -WarningAction SilentlyContinue
        If ($MessagePage) {
            $Messages += $MessagePage
        }
    } While ($MessagePage.count -eq $BatchSizeForMessages)
}
# Remove Exchange Online public folder hierarchy synchronization messages
$Messages = $Messages | Where-Object {$_.Subject -NotLike "*HierarchySync*"}
# Now, do we have any messsages to process?
If ($Messages.count -eq 0) {
    Write-Host "No messages found for analysis"
    Break
} Else {
    Write-Host ("After excluding system messages, there are {0} messages for analysis" -f $Messages.count)
}

[array]$Domains = Get-AcceptedDomain | Select-Object -ExpandProperty DomainName

$Report = [System.Collections.Generic.List[Object]]::new() 

ForEach ($M in $Messages) {
    $Direction = "Inbound"
    $SenderDomain = $M.SenderAddress.Split("@")[1]
    $RecipientDomain = $M.RecipientAddress.Split("@")[1]
    If ($SenderDomain -in $Domains) {
        $Direction = "Outbound" 
    }
    $ReportLine = [PSCustomObject]@{
        TimeStamp       = $M.Received
        Sender          = $M.SenderAddress
        Recipient       = $M.RecipientAddress
        Subject         = $M.Subject
        Status          = $M.Status
        Direction       = $Direction
        SenderDomain    = $SenderDomain
        RecipientDomain = $RecipientDomain
    }
    $Report.Add($ReportLine)

}

# Build an array of email addresses from this tenant used to send or receive messages
[array]$Senders = $Report | Where-Object {$_.SenderDomain -in $Domains} | Sort-Object Sender -Unique | Select-Object -ExpandProperty Sender
[array]$Recipients = $Report | Where-Object {$_.RecipientDomain -in $Domains} | Sort-Object Recipient -Unique | Select-Object -ExpandProperty Recipient
[array]$AllAddresses = $Senders + $Recipients | Sort-Object -Unique

$ReportAddresses = [System.Collections.Generic.List[Object]]::new()

# Find mailboxes
[array]$Mailboxes = Get-ExoMailbox -ResultSize Unlimited -Properties EmailAddresses -RecipientTypeDetails UserMailbox, SharedMailbox 

ForEach ($Mbx in $Mailboxes) {
    $MbxAddresses = $Mbx.EmailAddresses | Where-Object {$_ -like "smtp:*"} 
    ForEach ($Address in $MbxAddresses) {
        $SmtpAddress = $Address.Substring(5)  # Remove the smtp: prefix
        If ($SmtpAddress -notin $AllAddresses) {
            $ReportLine = [PSCustomObject]@{
                DisplayName = $Mbx.DisplayName
                PrimarySmtpAddress = $Mbx.PrimarySmtpAddress
                Alias = $Mbx.Alias
                MailboxType = $Mbx.RecipientTypeDetails
                UnusedProxyAddress = $SmtpAddress
            }
            $ReportAddresses.Add($ReportLine)
        }
    }   
}

