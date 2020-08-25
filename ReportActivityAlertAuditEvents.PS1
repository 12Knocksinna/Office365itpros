# ReportActivityAlertAuditEvents.PS1
# Used as an example in Chapter 20 
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportActivityAlertAuditEvents.PS1
$Records = (Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date).AddDays(+1) -RecordType SecurityComplianceAlerts -Formatted -ResultSize 3000)
If ($Records.Count -eq 0) {
   Write-Host "No alert audit records found." }
 Else {
   Write-Host "Processing" $Records.Count "audit records..."
   $Report = [System.Collections.Generic.List[Object]]::new()
   ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      $Data = ConvertFrom-Json $Auditdata.data
      If ($Rec.Operations -eq "AlertTriggered") {
      $ReportLine = [PSCustomObject]@{
           TimeStamp   = $Rec.CreationDate
           User        = $Data.f3u
           Action      = $Data.an
           Status      = $AuditData.ResultStatus
           Severity    = $AuditData.Severity
           Workload    = $AuditData.Source
           Operation   = $Rec.Operations
           Category    = $AuditData.Category }
      $Report.Add($ReportLine) }
    Else {
      $ReportLine = [PSCustomObject]@{
           TimeStamp   = $Rec.CreationDate
           User        = $Data.eid
           Action      = $Data.lon
           Status      = $AuditData.ResultStatus
           Severity    = $AuditData.Severity
           Workload    = $AuditData.Source
           Operation   = $Rec.Operations
           Category    = $AuditData.Category }
        $Report.Add($ReportLine) }
}} 

$Report | Select Timestamp, Operation, User, Action
