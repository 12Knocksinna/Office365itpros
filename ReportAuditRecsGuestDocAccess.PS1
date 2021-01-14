# ReportAuditRecsGuestDocAccess.PS1
# Example for Chapter 21 of Office 365 for IT Pros
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportAuditRecsGuestDocAccess.PS1
# Note: There might well be far more than 5000 FileAccessed audit records to be retrieved. You can either go back and fetch more or shorten the timeframe that you look.
$Records = (Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date).AddDays(+1) -Operations FileAccessed -ResultSize 5000 -Formatted | ? {$_.UserIds -Like "*#EXT#*" })
If ($Records.Count -eq 0) {
   Write-Host "No SharePoint file access records found for guest users." }
 Else {
   Write-Host "Processing" $Records.Count" SharePoint file access audit records..."
   $Report = [System.Collections.Generic.List[Object]]::new()
   ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      If ($AuditData.SourceFileName -NotLike "*aspx*" -And $AuditData.SourceFileName -NotLike "*jpg*" ) {
         $ReportLine = [PSCustomObject]@{
           TimeStamp   = $Rec.CreationDate
           User        = $Rec.UserIds
           Action      = $AuditData.Operation
           Workload    = $AuditData.Workload
           URL         = $AuditData.SiteUrl
           Document    = $AuditData.SourceFileName }      
        $Report.Add($ReportLine)}
      }}
$Report | Out-GridView
