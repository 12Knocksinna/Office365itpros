# AuditRecordsTeamsRecordingsUploads.PS1
# A script to show how to use audit records in the Office 365 audit log to track uploads of Teams meeting recordings

$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}

# Start and end date for the audit scan. By default, we look for 14 days, but you can choose \
# any value you like up to 365 (assuming Office 365 E5)
$StartDate = (Get-Date).AddDays(-14); $EndDate = (Get-Date) # Set your own date span here!
$OutputCSVFile = "C:\temp\AuditEventsTeamsRecordings.csv"

# Find the audit records
[array]$Records = (Search-UnifiedAuditLog -Operations FileUploaded, FileModified -StartDate $StartDate -EndDate $EndDate -Formatted -ResultSize 5000)

If (!($Records)) {Write-Host "No audit records found - exiting!"; break}

$TaggedRecordings = [System.Collections.Generic.List[Object]]::new() 	
ForEach ($Rec in $Records) {
   $AuditData = $Rec.AuditData | ConvertFrom-Json
   If (($AuditData.SourceFileExtension -eq "mp4") -and ($AuditData.SourceRelativeUrl -like "*/Recordings")) { 
      $RecordingFileName = $AuditData.SourceFileName
      $DateLoc = $RecordingFileName.IndexOf("-202")
      If ($DateLoc -eq -1) {$Topic = $RecordingFileName} Else {$Topic = $RecordingFileName.SubString(0,$DateLoc)}
      $DataLine = [PSCustomObject] @{
         Workload            = $AuditData.Workload
         Date                = $Rec.CreationDate
         User                = $Rec.UserIds
         Recording           = $RecordingFileName
         "Meeting title"     = $Topic
         Site                = $AuditData.SiteURL
         FullURL             = $AuditData.ObjectId
         Folder              = $AuditData.SourceRelativeURL
         Operation           = $Rec.Operations }
    $TaggedRecordings.Add($DataLine) 

   } #End If
} #End For

# Generate the list of audit events for Teams recordings
$UploadedMeetings = $TaggedRecordings | ? {$_.Recording -notlike "*~tmp*"}
$UploadedMeetings | Out-GridView
$UploadedMeetings | Export-CSV -NoTypeInformation $OutputCSVFile


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
