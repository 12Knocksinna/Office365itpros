# Example from Chapter 21 of Office 365 for IT Pros
# Find audit records for user sign ins, including connections to Teams
$Records = (Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date).AddDays(+1)   -Operations  UserLoggedIn,  TeamsSessionStarted  -SessionCommand ReturnLargeSet -ResultSize 5000 -Formatted)
If ($Records.Count -eq 0) {
   Write-Host "No Audit records found for user logons." }
 Else {
   Write-Host "Processing" $Records.Count "audit records..."
   $Report = [System.Collections.Generic.List[Object]]::new()
   ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      If ($AuditData.Workload -eq "MicrosoftTeams")
         { $Client =  $Auditdata.ObjectId 
           $Status = "OK"}
      Else
         { $Client = $AuditData.ExtendedProperties[0].Value
           $Status = $AuditData.ExtendedProperties[3].Value  }
      $ReportLine = [PSCustomObject]@{
           TimeStamp   = $Rec.CreationDate
           User        = $AuditData.UserId
           Action      = $AuditData.Operation
           Client      = $Client
           IpAddress   = $AuditData.ActorIpAddress
           Status      = $Status
           Workload    = $AuditData.Workload  }
      $Report.Add($ReportLine) }}
$Report | Out-GridView
