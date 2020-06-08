# ReportRetentionLabelsAuditEvents.PS1
# Example used in Chapter 20 of how to find and report retention labels assigned to documents.
$Records = (Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date).AddDays(+1) -Operations "TagApplied" -Formatted -ResultSize 2000)
If ($Records.Count -eq 0) {
    Write-Host "No retention label assignment records found." }
Else {
    Write-Host "Processing" $Records.Count "audit records..."
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
  $AuditData = ConvertFrom-Json $Rec.Auditdata
  If ($AuditData.UserType -ne "Regular") { # Library-applied Label
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
