# FindInactiveDls
# Find inactive distribution lists based on the message trace information, which means we can only go back ten days...
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindInactiveDLs.PS1

# V1.1 Updated 27-Aug-2025  Replaced old message trace cmdlet with Get-MessageTraceV2
# V2.0 Updated 10-June-2026 Fetch message data for DL expansion events for the entire 90 days it is available online.
#                           Also added export to Excel worksheet if the ImportExcel module is available and some other miscellaneous updates

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
        	-ResultSize $BatchSizeForMessages -Status "Expanded" -ErrorAction Stop -WarningAction SilentlyContinue
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
                -StartingRecipientAddress $LastMessageFetchedRecipient -ResultSize $BatchSizeForMessages -Status "Expanded" -ErrorAction Stop -WarningAction SilentlyContinue
            If ($MessagePage) {
                $Messages += $MessagePage
            }
        } While ($MessagePage.count -eq $BatchSizeForMessages)
}

    # Remove Exchange Online public folder hierarchy synchronization messages
    $Messages = $Messages | Where-Object {$_.Subject -NotLike "*HierarchySync*"}
    return $Messages
}
# Start of processing

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ($Modules -notcontains "ExchangeOnlineManagement") {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$False
}

# It's possible to go back 90 days, but the message trace data must be retrieved in 10-day chunnks.
# This code creates an array of 10-day ranges to fetch the message trace data.
$StartDate = (Get-Date).AddDays(-90).ToUniversalTime()
$EndDate = (Get-Date).ToUniversalTime().AddMinutes(10)

$iterations = 9
[int]$i = 0

[array]$DateRanges = $null

For ($i=0; $i -lt $iterations; $i++) {
    $RangeStart = $StartDate.AddDays($i * 10)
    $RangeEnd = $StartDate.AddDays(($i + 1) * 10)
    If ($RangeEnd -gt $EndDate) {
        $RangeEnd = $EndDate
    }
    $DateRanges += @{Start=$RangeStart; End=$RangeEnd}
}

[array]$Data = $null
[array]$Messages = $null

# Fetch each of the 10-day chunks of message trace data and add it to the Messages array
ForEach ($DateRange in $DateRanges) {
    $Data = Get-DateRangeData -StartDate $DateRange.Start -EndDate $DateRange.End
    # Process the data for this date range, for example by adding it to a report or exporting it to a file
    Write-Host ("Fetched {0} messages between {1} and {2}" -f $Data.count, $DateRange.Start, $DateRange.End)
    $Messages+= $Data
}

# Sort the messages by received date in descending order so that we can easily find the most recent message for each distribution list
$Messages = $Messages | Sort-Object {$_.Received -as [datetime]} -Descending

# Create an array of unique recipient addresses from the message trace data, which will be used to check against the distribution lists in the tenant
[array]$MessageTable = ($Messages | Sort-Object RecipientAddress -Unique)
[array]$DLs = Get-DistributionGroup -ResultSize Unlimited -RecipientTypeDetails 'MailUniversalDistributionGroup'
Write-Host ("Processing {0} distribution lists..." -f $DLs.count)
$Report = [System.Collections.Generic.List[Object]]::new()
[int]$ActiveStatusCount = 0
ForEach ($DL in $DLs) {
    $LastActiveDate = "Unknown"; $NumberOfExpandedEvents = 0
    If ($MessageTable.recipientaddress -Match $DL.PrimarySMTPAddress) {
        $ActiveStatus = "Active"; $ActiveStatusCount++
        $LastActiveDate = $MessageTable | Where-Object {$_.RecipientAddress -eq $DL.PrimarySMTPAddress} | Select-Object -ExpandProperty Received
        $LastActiveDate = Get-Date ($LastActiveDate) -format 'dd-MMM-yyyy HH:mm:ss'
        $NumberOfExpandedEvents = ($Messages | Where-Object {$_.RecipientAddress -eq $DL.PrimarySMTPAddress}).Count
        Write-Host ("{0} is active - most recent message was on {1} with a total of {2} messages found" -f $DL.DisplayName, $LastActiveDate, $NumberOfExpandedEvents) -Foregroundcolor Yellow 
    } Else {
        $ActiveStatus = "Inactive"
        Write-Host ("{0} is inactive" -f $DL.DisplayName) -Foregroundcolor Red 
    }
    $Reportline = [pscustomobject]@{
        Name = $DL.DisplayName 
        Active = $ActiveStatus 
        'Last active date' = $LastActiveDate
        'Number of messages sent to DL' = $NumberOfExpandedEvents
    }
    $Report.Add($ReportLine)
    $Text = ("DL state checked on {0} and determined as {1}. Last message addressed on {2}" `
         -f (Get-Date -format g), $ActiveStatus, $LastActiveDate )
      # This line updates the DL with details of the assessment. Comment it out if you don't want to do this
    Set-DistributionGroup -Identity $DL.ExternalDirectoryObjectId -CustomAttribute15 $Text
}

# Set up to use the ImportExcel module to generate Excel worksheets if it is available. If not, the report will be generated in CSV format instead.
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
}

# Generate the attachment in either Excel worksheet or CSV format, depending on if the ImportExcel module is available. If interactive, create
# the file in the user's Downloads folder, otherwise create it in the current directory.
If ($ExcelGenerated) {
    $OutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveDLs.xlsx"
    Remove-Item -Path $OutputFile -ErrorAction SilentlyContinue
    $Report | Export-Excel -Path $OutputFile -WorksheetName "Inactive DLs" -Title ("Inactive DLs {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "InactiveDLs" -AutoSize -AutoFilter -FreezeTopRow   
} Else {
    $OutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveDLs.csv"
    Remove-Item -Path $OutputFile -ErrorAction SilentlyContinue
    $Report | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding Utf8
}

Write-Host "Report generated in: $OutputFile" -ForegroundColor Green

Write-Host ""
Write-Host ("Total distribution lists checked: {0} using messages sent between {1} and {2}" -f $DLs.count, (Get-Date $StartDate -format 'dd-MMM-yyyy HH:mm'), (Get-Date $EndDate -format 'dd-MMM-yyyy HH:mm'))
Write-Host ("Active distribution lists:        {0} ({1:P2})" -f $ActiveStatusCount, ($ActiveStatusCount / $DLs.count))
Write-Host ("Inactive distribution lists:      {0} ({1:P2})" -f ($DLs.count - $ActiveStatusCount), (($DLs.count - $ActiveStatusCount) / $DLs.count))
Write-Host ""   

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# The Automating Microsoft 365 with PowerShell eBook is available from https://o365itpros.gumroad.com/l/M365PS in EPUB and PDF format 
# or in paperback at https://amzn.to/3L4IBmU

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.