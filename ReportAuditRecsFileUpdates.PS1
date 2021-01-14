# ReportAuditRecsFileUpdates.PS1
# Example of using audit records to find out who changed a file
$FileName = (Read-Host "Enter file name to search")
$Records = (Search-UnifiedAuditLog -Operations FileModified, FileAccessed, FileUploaded -ObjectIds $FileName -ResultSize 1000 -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date).AddDays(+1))
If ($Records.Count -eq 0) {
   Write-Host "No audit records found for file names beginning with" $FileName }
 Else {
   Write-Host "Processing" $Records.Count "audit records..."
   $Report = [System.Collections.Generic.List[Object]]::new()
   ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      $ReportLine = [PSCustomObject]@{
           TimeStamp   = $Rec.CreationDate
           User        = $AuditData.UserId
           Action      = $AuditData.Operation
           SiteUrl     = $AuditData.SiteUrl
           Site        = $AuditData.SourceRelativeUrl
           File        = $AuditData.SourceFileName
           IpAddress   = $AuditData.ClientIP
           App         = $AuditData.UserAgent  }
      $Report.Add($ReportLine) }}
Write-Host “All done”
$Report | Out-GridView
