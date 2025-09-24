# Find-UnusedProxyAddresses.PS1

# A script to check user and shared mailboxes to discover if any of the SMTP proxy addresses assigned to the mailbox    
# have not been used to send or receive email in the last 10 days. The script uses message trace data fetched from Exchange Online.

# Base on the code written for https://office365itpros.com/2023/08/23/message-trace-analysis/ to analyze mail traffic. See
# https://github.com/12Knocksinna/Office365itpros/blob/master/Analyze-MailTraffic.PS1

# Version 1.0 24-September-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-UnusedProxyAddresses-10days.PS1


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
[array]$Mailboxes = Get-ExoMailbox -ResultSize Unlimited -Properties EmailAddresses -RecipientTypeDetails UserMailbox, SharedMailbox | Sort-Object DisplayName

# This code to check against the array of all senders and recipients in the last 10 days
ForEach ($Mbx in $Mailboxes) {
    $MbxAddresses = $Mbx.EmailAddresses | Where-Object {$_ -like "smtp:*"} 
    ForEach ($Address in $MbxAddresses) {
        $SmtpAddress = $Address.Substring(5)  # Remove the smtp: prefix
        If ($SmtpAddress -notin $AllAddresses) {
            $ReportLine = [PSCustomObject]@{
                DisplayName         = $Mbx.DisplayName
                PrimarySmtpAddress  = $Mbx.PrimarySmtpAddress
                Alias               = $Mbx.Alias
                MailboxType         = $Mbx.RecipientTypeDetails
                UnusedProxyAddress  = $SmtpAddress
                CheckType           = "Sender"
            }
            $ReportAddresses.Add($ReportLine)
        }   
    }
}

$ReportAddresses | Out-GridView

generate an output file
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\UnusedProxyAddresses.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    } 
    $ReportAddresses | Export-Excel -Path $ExcelOutputFile -WorksheetName "Unused Proxy Addresses" -Title ("Unused Proxy Addresses Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "UnusedProxyAddresses" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\UnusedProxyAddresses.CSV"
    $ReportAddresses | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}

If ($ExcelGenerated) {
    Write-Host ("Excel worksheet output written to {0}" -f $ExcelOutputFile)
} Else {
    Write-Host ("CSV output file written to {0}" -f $CSVOutputFile)
}   

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.