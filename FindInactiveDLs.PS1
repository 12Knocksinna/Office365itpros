# FindInactiveDls
# Find inactive distribution lists based on the message trace information, which means we can only go back ten days...
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindInactiveDLs.PS1

# V1.1 Updated 27-Aug-2025  Replaced old message trace cmdlet with Get-MessageTraceV2
# V2.0 Updated 7-June-2026  Fetch message data for DL expansion events for the entire 90 days it is available online.
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
    # Call to API to get data for the specified date range goes here, for example:
    # $Data = Get-MgAuditLog -StartDateTime $StartDate -EndDateTime $EndDate
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
$EndDate = (Get-Date).ToUniversalTime()

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

$Report = [System.Collections.Generic.List[Object]]::new()

[array]$Data = $null
[array]$Messages = $null

# Fetch each of the 10-day chunks of message trace data and add it to the Messages array
ForEach ($DateRange in $DateRanges) {
    $Data = Get-DateRangeData -StartDate $DateRange.Start -EndDate $DateRange.End
    # Process the data for this date range, for example by adding it to a report or exporting it to a file
    Write-Host ("Fetched {0} messages between {1} and {2}" -f $Data.count, $DateRange.Start, $DateRange.End)
    $Messages+= $Data
}

# Create an array of unique recipient addresses from the message trace data, which will be used to check against the distribution lists in the tenant
[array]$MessageTable = ($Messages | Sort-Object RecipientAddress -Unique)
[array]$DLs = Get-DistributionGroup -ResultSize Unlimited -RecipientTypeDetails 'MailUniversalDistributionGroup'
Write-Host ("Processing {0} distribution lists..." -f $DLs.count)
[int]$ActiveStatusCount = 0
ForEach ($DL in $DLs) {
   $LastActiveDate = "Unknown"
   If ($MessageTable.recipientaddress -Match $DL.PrimarySMTPAddress) {
      $ActiveStatus = "Active"; $ActiveStatusCount++
      $LastActiveDate = $MessageTable | Where-Object {$_.RecipientAddress -eq $DL.PrimarySMTPAddress} | Select-Object -ExpandProperty Received
      $LastActiveDate = Get-Date ($LastActiveDate) -format 'dd-MMM-yyyy HH:mm:ss'
      Write-Host ("{0} is active - message found on {1}" -f $DL.DisplayName, $LastActiveDate) -Foregroundcolor Yellow 
   } Else {
      $ActiveStatus = "Inactive"
      Write-Host ("{0} is inactive" -f $DL.DisplayName) -Foregroundcolor Red 
   }
   $Reportline = [pscustomobject]@{
      Name = $DL.DisplayName 
      Active = $ActiveStatus 
      LastActiveDate = $LastActiveDate
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
    $Report | Export-Excel -Path $OutputFile -WorksheetName "Inactive DLs" -Title ("Inactive DLs {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "InactiveDLs" -AutoSize -AutoFilter -FreezeTopRow   
} Else {
    $OutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveDLs.csv"
    $Report | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding Utf8
}

Write-Host "Report generated in: $OutputFile" -ForegroundColor Green

Write-Host ""
Write-Host ("Total distribution lists checked: {0}" -f $DLs.count)
Write-Host ("Active distribution lists:        {0}" -f $ActiveStatusCount)
Write-Host ("Inactive distribution lists:      {0}" -f ($DLs.count - $ActiveStatusCount))
Write-Host ""

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.