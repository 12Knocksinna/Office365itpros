# Analyze-Mailtraffic.PS1
# An example of using the Exchange Online message trace log to analyze inbound and outbound traffic
# for a Microsoft 365 tenant. See https://office365itpros.com/2023/08/23/message-trace-analysis/ for more information about the script.
# https://github.com/12Knocksinna/Office365itpros/blob/master/Analyze-MailTraffic.PS1

# Updated 4-June-2025 to replace Get-MessageTrace with Get-MessageTraceV2, which is the cmdlet that will be used in the future.
# Updated 15-Jul-2025 to generate a HTML report of the results.

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
# Extract the inbound and outbound messages
[array]$OutboundMessages = $Report | Where-Object {$_.Direction -eq "Outbound"}
[array]$InboundMessages = $Report | Where-Object {$_.Direction -eq "Inbound"}

Write-Host ""
# Report the top 10 domains for outbound messages
Write-Host ("Top 10 domains for outbound messages between {0} and {1}" -f (Get-Date $StartDate -format 'dd-MMM-yyyy HH:mm'), (Get-Date $EndDate -format 'dd-MMM-yyyy HH:mm'))
Write-Host "------------------------------------------------------------------------------------"
$OutboundMessages | Group-Object RecipientDomain -NoElement | Sort-Object Count -Descending | Select-Object -First 10 | Format-Table Name, Count -AutoSize
Write-Host ""
# And the same for inbound messages
Write-Host "Top 10 domains for inbound messages"
Write-Host "-----------------------------------"
$InboundMessages | Group-Object SenderDomain -NoElement | Sort-Object Count -Descending | Select-Object -First 10 | Format-Table Name, Count -AutoSize


# Generate sortable HTML table with type-aware sorting - use number as the type for numeric values, date for dates, and string for text
$HtmlHeader = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Detailed Message Statistics</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; background: #f4f6f8; color: #222; }
h1 { background: #0078d4; color: #fff; padding: 16px; border-radius: 6px 6px 0 0; margin-bottom: 20px; }
table { width: 100%; background: #fff; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-collapse: collapse; }
th, td { padding: 12px; text-align: left; }
th { background: #e5eaf1; cursor: pointer; position: relative; }
th:hover { background: #d0e7fa; }
th::after { content: '↕'; position: absolute; right: 8px; opacity: 0.5; }
tr:nth-child(even) { background: #f0f4fa; }
tr:hover { background: #d0e7fa; }
</style>
<script>
function parseValue(val, type) {
    if(type === 'number') return parseFloat(val.replace(/,/g,'')) || 0;
    if(type === 'date') return new Date(val);
    return val.toLowerCase();
}
function sortTable(n, type) {
    var table = document.getElementById('msgstats');
    var rows = Array.from(table.rows).slice(1);
    var dir = table.getAttribute('data-sortdir'+n) === 'asc' ? 'desc' : 'asc';
    rows.sort(function(a, b) {
        var x = parseValue(a.cells[n].innerText, type);
        var y = parseValue(b.cells[n].innerText, type);
        if(x < y) return dir === 'asc' ? -1 : 1;
        if(x > y) return dir === 'asc' ? 1 : -1;
        return 0;
    });
    rows.forEach(function(row) { table.tBodies[0].appendChild(row); });
    table.setAttribute('data-sortdir'+n, dir);
}
</script>
</head>
<body>
<h1>Detailed Message Statistics</h1>
<table id="msgstats">
<thead>
<tr>
<th onclick="sortTable(0,'date')">Timestamp</th>
<th onclick="sortTable(1,'string')">Sender</th>
<th onclick="sortTable(2,'string')">Recipient</th>
<th onclick="sortTable(3,'string')">Subject</th>
<th onclick="sortTable(4,'string')">Direction</th>
<th onclick="sortTable(5,'string')">SenderDomain</th>
<th onclick="sortTable(6,'string')">RecipientDomain</th>
</tr>
</thead>
<tbody>
"@

$Report = $Report | Sort-Object {$_.Timestamp -as [datetime]}, {$_.Recipient} -Descending
$HtmlRows = foreach ($Row in $Report ) {
    "<tr><td>$($row.Timestamp)</td><td>$($row.Sender)</td><td>$($row.Recipient)</td><td>$($row.Subject)</td><td>$($row.Direction)</td><td>$($row.SenderDomain)</td><td>$($row.RecipientDomain)</td></tr>"
}

$HtmlFooter = @"
</tbody>
</table>
</body>
</html>
"@

$ReportFile = "c:\temp\MessageStats.html"
#Generate the full HTML content and save it to a file
$HtmlFile = $HtmlHeader + ($HtmlRows -join "`n") + $HtmlFooter
$HtmlFile | Out-File -FilePath $ReportFile -Encoding utf8

Write-Host "Detailed message statistics saved to $ReportFile" -ForegroundColor Green

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
