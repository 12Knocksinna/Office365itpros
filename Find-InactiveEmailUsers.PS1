# Find-InactiveEmailUsers.PS1
# Find inactive email users based on the message trace information, which means we can only go back ten days...

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-InactiveEmailUsers.PS1

# V1.0 26-Nov-2024
# V2.0 08-June-2026 Updated to use the Get-MessageTraceV2 cmdlet. Also, the script can now fetch message trace
#                   data for the entire 90 days it is available online, instead of just the last 10 days.
#                   The script now also generates a report in Excel format if the ImportExcel module is available, and some other miscellaneous updates.

function Get-DateRangeData {
    param (
        [datetime]$StartDate,
        [datetime]$EndDate
    )
    
    # Function to fetch message trace data for expanded events over a 10-day period.
    [array]$Messages = $null
    Write-Host ("Collecting message trace data between {0} and {1}..." -f $StartDate, $EndDate)
    [int]$BatchSizeForMessages = 2000
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
            Write-Host ("Fetched {0} messages so far for this chunk..." -f $Messages.count)
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
    return $Messages
}

# -- Start of processing ---

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ($Modules -notcontains "ExchangeOnlineManagement") {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$False
}

# It's possible to go back 90 days, but the message trace data must be retrieved in 10-day chunnks.
# This code creates an array of 1-day ranges to fetch the message trace data. We use 1-day because Get-MessageTraceV2
# can fetch up to 5000 messages.
$StartDate = (Get-Date).AddDays(-90).ToUniversalTime()
$EndDate = (Get-Date).ToUniversalTime()

$iterations = 90
[int]$i = 0

[array]$DateRanges = $null

For ($i=0; $i -lt $iterations; $i++) {
    $RangeStart = $StartDate.AddDays($i + 1)
    $RangeEnd = $StartDate.AddDays(($i + 1) +1 )
    If ($RangeEnd -gt $EndDate) {
        $RangeEnd = $EndDate
    }
    $DateRanges += @{Start=$RangeStart; End=$RangeEnd}
}

$Report = [System.Collections.Generic.List[Object]]::new()

[array]$Data = $null
[array]$Messages = $null

# Fetch each of the 1-day chunks of message trace data and add it to the Messages array
ForEach ($DateRange in $DateRanges) {
    $Data = Get-DateRangeData -StartDate $DateRange.Start -EndDate $DateRange.End
    # Process the data for this date range, for example by adding it to a report or exporting it to a file
    Write-Host ("Fetched {0} messages between {1} and {2}" -f $Data.count, $DateRange.Start, $DateRange.End)
    $Messages+= $Data
}

# Make sure that the messages are sorted in descending order by received date, so that we can easily find the last message sent by a user.
$Messages = $Messages | Sort-Object {$_.Received -as [datetime]} -Descending

# Create an array of unique sender addresses from the message trace data, which will be used to check against the distribution lists in the tenant
[array]$MessageTable = ($Messages | Sort-Object SenderAddress -Unique)

Write-Host "Looking for mailboxes..."
[array]$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails 'UserMailbox' | Sort-Object DisplayName
Write-Host ("Processing {0} mailboxes to check activity..." -f $Mbx.count)
$Report = [System.Collections.Generic.List[Object]]::new()
[int]$ActiveStatusCount = 0
ForEach ($M in $Mbx) {
    $LastActiveDate = $null
    If ($MessageTable.SenderAddress -Match $M.PrimarySMTPAddress) {
        $ActiveStatus = "Active"; $ActiveStatusCount++
        $LastActiveDate = $Messages | Where-Object {$_.SenderAddress -eq $M.PrimarySMTPAddress} | Select-Object -First 1 | Select-Object -ExpandProperty Received
        If ($null -ne $LastActiveDate) {
            $LastActiveDate = Get-Date ($LastActiveDate) -format 'dd-MMM-yyyy HH:mm:ss'
        } Else {
            $LastActiveDate = "N/A"
        }
        Write-Host ("{0} is active - message found on {1}" -f $M.DisplayName, $LastActiveDate) -Foregroundcolor Yellow 
    } Else {
        $ActiveStatus = "Inactive"
        $LastActiveDate = $null
        Write-Host ("{0} is inactive" -f $M.DisplayName) -Foregroundcolor Red 
    }

    $Reportline = [pscustomobject]@{
        Name                    = $M.DisplayName 
        UPN                     = $M.UserPrincipalName
        PrimarySMTPAddress      = $M.PrimarySMTPAddress
        Active                  = $ActiveStatus 
        'Last sent message'     = $LastActiveDate
    }
    $Report.Add($ReportLine)
    $Text = ("Mailbox state checked on {0} and determined as {1}. Last message sent on {2}" `
         -f (Get-Date -format 'dd-MMM-yyyy HH:mm'), $ActiveStatus, $LastActiveDate )
      # This line updates the mailbox with details of the assessment. Comment it out if you don't want to do this
    Set-Mailbox -Identity $M.Alias -CustomAttribute15 $Text
}

Write-Host ""
Write-Host ("Total mail users checked: {0}" -f $Mbx.count)
Write-Host ("Active mail users:        {0}" -f $ActiveStatusCount)
Write-Host ("Inactive mail users:      {0}" -f ($Mbx.count - $ActiveStatusCount))
Write-Host ""

# Generate reports
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $OutputXLSXFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveMailUsers.xlsx"
    Remove-Item -Path $OutputXLSXFile -ErrorAction SilentlyContinue
    $Report | Export-Excel -Path $OutputXLSXFile -WorksheetName "Inactive Mail Users Report" -Title ("Inactive Mail Users Report{0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "InactiveMailUsers" 
} Else {
    $OutputCSVFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveMailUsers.csv"
    Remove-Item -Path $OutputCSVFile -ErrorAction SilentlyContinue
    $Report | Export-Csv -Path $OutputCSVFile -NoTypeInformation -Encoding Utf8
}
  
If ($ExcelGenerated) {
  Write-Host ("An Excel report is available in {0}" -f $OutputXLSXFile)
} Else {
  Write-Host ("A CSV report is available in {0}" -f $OutputCSVFile)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# The Automating Microsoft 365 with PowerShell eBook is available from https://o365itpros.gumroad.com/l/M365PS in EPUB and PDF format 
# or in paperback at https://amzn.to/3L4IBmU

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.