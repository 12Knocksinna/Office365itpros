# FindInactiveDls
# Find inactive distribution lists based on the message trace information, which means we can only go back ten days...
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindInactiveDLs.PS1

# Updated 27-Aug-2025 for the new Message trace

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ($Modules -notcontains "ExchangeOnlineManagement") {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$False
}

$EndDate = Get-Date
$StartDate = $EndDate.AddDays(-10)
$CSVFile = "c:\Temp\ListofDLs.csv"
[array]$Messages = $Null 
[int]$BatchSizeForMessages = 2000

$Report = [System.Collections.Generic.List[Object]]::new()

Write-Host "Collecting message trace data for the last 10 days"
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