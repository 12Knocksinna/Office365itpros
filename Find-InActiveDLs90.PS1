# Find-InactiveDls90.PS1
# Find inactive distribution lists based on hostorical message trace information, which allows us to go back 90 days.
# https://github.com/12Knocksinna/Office365itpros/blob/master/Find-InActiveDLs90.PS1

# Updated 29-Nov-2023

Connect-ExchangeOnline

$Report = [System.Collections.Generic.List[Object]]::new()
$DataFolder = "c:\temp\MtData\"
$CSVFile = "c:\temp\HistoricalDLMessageTrace.CSV"

[array]$DataFiles = Get-ChildItem -Path $DataFolder | Select-Object -ExpandProperty Name
If (!($DataFiles)) {
    Write-Host "No historical message tracking logs to analyze - exiting"
    Break
}

Write-Host ("Preparing to process {0} historical message trace data files..." -f $DataFiles.count)
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report

# Create a hash table of distribution list SMTP addresses and display names that we can use to
# check if a recipient in a message trace file is a DL
[array]$DLs = Get-DistributionGroup -ResultSize Unlimited
$DLSMTPAddresses = @{}
ForEach ($DL in $DLs) {
   $DLSMTPAddresses.Add([string]$DL.PrimarySMTPAddress,[string]$DL.DisplayName)
}

# Loop through the historical trace files to find message trace data related to messages sent to
# distribution lists in the tenant
ForEach ($File in $DataFiles) {
   $MtDataFile = $DataFolder + $File
   [array]$MtData = Import-CSV -Path $MtDataFile -Encoding unicode
   ForEach ($Line in $MtData) {
      If (!([string]::IsNullOrEmpty($Line.origin_timestamp_utc))) {
         [array]$RecipientStatus = $Line.Recipient_Status.split(";")
         # array of individual recipients for a message
         ForEach ($RecipientDetail in $RecipientStatus) {
            $DLName = $Null
            $Recipient = $RecipientDetail.Split("##")[0]
            $DLName = $DLSMTPAddresses[$Recipient]
            If ($DLName) { # DL recipient found
               $SenderDomain  = $Line.Sender_address.Split("@")[1]
               $ReportLine = [PSCustomObject]@{ 
                  Timestamp        = $Line.origin_timestamp_utc
                  Sender           = $Line.sender_address
                  Subject          = $Line.message_subject
                  DLSMTP           = $Recipient
                  DLName           = $DLName
                  Bytes            = $Line.total_bytes
                  Message_id       = $Line.message_id
                  Sender_Domain    = $SenderDomain
                  Client_IP        = $Line.original_client_ip
                  Direction        = $Line.directionality
               }
               $Report.Add($ReportLine) 
            }
         }
      }
   }
}  

$OutputReport = [System.Collections.Generic.List[Object]]::new()
[int]$TotalDLOK = 0
# Check each DL to see if we can find a record
ForEach ($DL in $DLs) {
   [array]$DLFound = $Report | Where-Object {$_.DLSMTP -eq $DL.PrimarySMTPAddress} | Sort-Object -Descending {$_.TimeStamp -as [datetime]} | `
      Select-Object -First 1
   If ($DLFound) {
      $DateLastMessage = (Get-Date $DLFound.TimeStamp -format g)
      Write-Host ("Found message for Distribution list {0} at {1}" -f $DL.DisplayName, $DateLastMessage) -Foregroundcolor Red
      $Text = ("DL state checked on {0} and determined as active. Last message addressed on {1}" `
         -f (Get-Date -format g), $DateLastMessage )
      Set-DistributionGroup -Identity $DL.Alias -CustomAttribute15 $Text
      $TotalDLOK++
      $ReportLine = [PSCustomObject]@{ 
         Timestamp = $DLFound.Timestamp
         Sender    = $DLFound.Sender
         DLName    = $DLFound.DLName
         DLSMTP    = $DLFound.DLSMTP
         Subject   = $DLFound.Subject
      }
      $OutputReport.Add($ReportLine)
   } Else {
      Write-Host ("No messages found for distribution list {0}" -f $DL.DisplayName) -ForegroundColor Yellow
      $ReportLine = [PSCustomObject]@{ 
         Timestamp = "N/A"
         Sender    = "N/A"
         DLName    = $DL.DisplayName
         DLSMTP    = $DL.PrimarySMTPAddress
         Subject   = "No messages found"
      }
      $OutputReport.Add($ReportLine)
   }
}

$OutputReport | Sort-Object TimeStamp | Out-GridView
$OutputReport | Export-CSV -NoTypeInformation $CSVFile -Encoding utf8
$PercentOK = ($TotalDLOK/$DLs.Count).ToString("P")

Write-Host ""
Write-Host ("Total distribution lists checked:     {0}" -f $DLs.count)
Write-Host ("Active distribution lists:            {0}" -f $TotalDLOK)
Write-Host ("Percentage active distribution lists: {0}" -f $PercentOK)
Write-Host ("Inactive distribution lists:          {0}" -f ($DLs.count - $TotalDLOK))
Write-Host ""
Write-Host ("Report file available in:             {0}" -f $CSVFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.