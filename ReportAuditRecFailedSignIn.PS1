# ReportAuditRecFailedSignIn.PS1
# Example used in Chapter 21 of Office 365 for IT Pros
$Records = (Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date).AddDays(+1)  -Operations UserLoginFailed  -SessionCommand  ReturnLargeSet     -ResultSize 5000 -Formatted)
If ($Records.Count -eq 0) {
   Write-Host "No audit records found for failed logins." }
 Else {
   Write-Host "Processing" $Records.Count "audit records..."
   $Report = [System.Collections.Generic.List[Object]]::new()
   ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      $ReportLine = [PSCustomObject]@{
           TimeStamp   = $Rec.CreationDate
           User        = $AuditData.UserId
           Action      = $AuditData.Operation
           Status      = $AuditData.ResultStatus
           IpAddress   = $AuditData.ActorIpAddress
           Error       = $AuditData.LogonError
           UserAgent   = $AuditData.ExtendedProperties.value[0] }
      $Report.Add($ReportLine) }
}
$Report | Sort User, Timestamp | Select Timestamp, User, IpAddress, UserAgent
