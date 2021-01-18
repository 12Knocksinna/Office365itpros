# ReportRetentionLabelsAuditEvents.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportRetentionLabelAuditEvents.PS1
# Example used in Chapter 20 of how to find and report retention labels assigned to documents.
$Records = (Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date).AddDays(+1) -Operations "TagApplied" -Formatted -ResultSize 2000)
If ($Records.Count -eq 0) {
    Write-Host "No retention label assignment records found." }
Else {
    Write-Host "Processing" $Records.Count "audit records..."
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
  $AuditData = ConvertFrom-Json $Rec.Auditdata
  If ($AuditData.UserType -ne "Regular") { # auto-applied Label
      $ReportLine = [PSCustomObject]@{
       TimeStamp = $Rec.CreationDate
       User      = "Auto-Applied Label"
       Action    = $AuditData.Operation
       Label     = $Rec.UserIds
       Type      = "File"
       File      = $AuditData.SourceFileName
       Library   = $AuditData.SourceRelativeUrl
       Site      = $AuditData.SiteURL }
    $Report.Add($ReportLine) }
  Else { # Label applied by a user
      $ReportLine = [PSCustomObject]@{
       TimeStamp = $Rec.CreationDate
       User      = $Rec.UserIds
       Action    = $AuditData.Operation
       Label     = $AuditData.DestinationLabel
       Type      = $AuditData.ItemType
       File      = $AuditData.SourceFileName
       Library   = $AuditData.SourceRelativeUrl
       Site      = $AuditData.SiteURL }
       $Report.Add($ReportLine) }
}}
Write-Host "All done!"
$Report | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
