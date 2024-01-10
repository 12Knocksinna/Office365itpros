# Report-StreamAuditEvents.PS1
# Report the audit events collected in the audit event log for Stream video creation and editing
# V1.0 10-Jan-2023
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-StreamAuditEvents.PS1

$Status = Get-ConnectionInformation -ErrorAction SilentlyContinue
If (!($Status)) {
  Connect-ExchangeOnline -SkipLoadingCmdletHelp 
}

$EndDate = (Get-Date).AddDays(1); $StartDate = (Get-Date).AddDays(-30)
[array]$Operations = "FileUploaded", "FileModified", "FileAccessed"
Write-Host "Looking for Stream audit events..."
[array]$Records = (Search-UnifiedAuditLog -Operations $Operations -StartDate $StartDate -EndDate $EndDate `
  -Formatted -ResultSize 5000 -SessionCommand ReturnLargeSet)
If (!($Records)) {
    Write-Host "No audit records found - exiting!"; break
}

Write-Host "Processing events..."
$StreamEvents = [System.Collections.Generic.List[Object]]::new() 
$Records = $Records | Sort-Object CreatedDateTime
ForEach ($Rec in $Records) {
   $AuditData = $Rec.AuditData | ConvertFrom-Json
   If (($AuditData.SourceFileExtension -in "mp4", "webm")) { 
      $RecordingFileName = $AuditData.SourceFileName
      $DateLoc = $RecordingFileName.IndexOf("-202")
      $Topic = $null
      If ($DateLoc -eq -1) {$Topic = $RecordingFileName} Else 
         {$Topic = $RecordingFileName.SubString(0,$DateLoc)}
      $DataLine = [PSCustomObject] @{
         Workload            = $AuditData.Workload
         Timestamp           = $Rec.CreationDate
         User                = $Rec.UserIds
         Recording           = $RecordingFileName
         Topic               = $Topic
         Site                = $AuditData.SiteURL
         FullURL             = $AuditData.ObjectId
         Folder              = $AuditData.SourceRelativeURL
         Operation           = $Rec.Operations }
    $StreamEvents.Add($DataLine) 
   } #End If
} #End For

Write-Host ""
$StreamFileOperations = $StreamEvents | Sort-Object {$_.Timestamp -as [datetime]} | Where-Object { $_.User -ne "app@sharepoint" }
$StreamFileOperations | Group-Object User | Sort-Object Count -Descending | Format-Table Name, Count
$StreamFileOperations | Out-GridView


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
