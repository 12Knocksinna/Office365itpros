# Analyze-Mailtraffic.PS1
# An example of using the Exchange Online message trace log to analyze inbound and outbound traffic
# for a Microsoft 365 tenant. See https://office365itpros.com/2023/08/23/message-trace-analysis/ for more information about the script.
# https://github.com/12Knocksinna/Office365itpros/blob/master/Analyze-MailTraffic.PS1

# Updated 4-June-2025 to replace Get-MessageTrace with Get-MessageTraceV2, which is the cmdlet that will be used in the future.

# Check if we can run an Exchange Online cmdlet. If we can, go on, else connect to Exchange Online
If ($Null -eq (Get-ConnectionInformation)) {
    Connect-ExchangeOnline
}

# Message trace date is kept for a maximum of 10 days
$StartDate = (Get-Date).AddDays(-10)
$EndDate = (Get-Date)

Write-Host ("Message trace data will be analyzed between {0} and {1}" -f $StartDate, $EndDate)
Write-Host "Fetching message trace data to analyze"
$MoreMessagesAvailable = $true
[array]$Messages = $Null
[int]$BatchSizeForMessages = 2000
    # original code [array]$MessagePage = Get-MessageTraceV2 -StartDate $StartDate -EndDate $EndDate -PageSize 1000 -Page $i -Status "Delivered"
Try {
    [array]$MessagePage = Get-MessageTraceV2 -StartDate $StartDate -EndDate $EndDate -ResultSize $BatchSizeForMessages -Status "Delivered" -ErrorAction Stop
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
            -StartingRecipientAddress $LastMessageFetchedRecipient -ResultSize $BatchSizeForMessages -Status "Delivered" -ErrorAction Stop
        If ($MessagePage) {
            $Messages += $MessagePage
        }
    } While ($MessagePage.count -eq $BatchSizeForMessages)
} Else {
    Write-Host "No more messages available."
    $MoreMessagesAVailable = $false
}
   
# Remove Exchange Online public folder hierarchy synchronization messages
$Messages = $Messages | Where-Object {$_.Subject -NotLike "*HierarchySync*"}

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
# Extract the inbound and outbound messages
[array]$OutboundMessages = $Report | Where-Object {$_.Direction -eq "Outbound"}
[array]$InboundMessages = $Report | Where-Object {$_.Direction -eq "Inbound"}

Write-Host ""
# Report the top 10 domains for outbound messages
Write-Host "Top 10 domains for outbound messages"
Write-Host "------------------------------------"
$OutboundMessages | Group-Object RecipientDomain -NoElement | Sort-Object Count -Descending | Select-Object -First 10 | Format-Table Name, Count -AutoSize
Write-Host ""
# And the same for inbound messages
Write-Host "Top 10 domains for inbound messages"
Write-Host "-----------------------------------"
$InboundMessages | Group-Object SenderDomain -NoElement | Sort-Object Count -Descending | Select-Object -First 10 | Format-Table Name, Count -AutoSize

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
