# Report-ActivityAlertAuditEvents.PS1
# Used as an example in Chapter 20 
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ActivityAlertAuditEvents.PS1

# Updated 15-Feb-2025

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Connect-ExchangeOnline -ShowBanner:$false
}

Write-Host "Searching for audit records"
[array]$Records = (Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-180) -EndDate (Get-Date).AddDays(+1) -RecordType SecurityComplianceAlerts -Formatted -ResultSize 3000)
If ($Records.Count -eq 0) {
   Write-Host "No alert audit records found." 
   Break
}

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
        Category    = $AuditData.Category 
      }
      $Report.Add($ReportLine) 
  } Else {
      $ReportLine = [PSCustomObject]@{
        TimeStamp   = $Rec.CreationDate
        User        = $Data.eid
        Action      = $Data.lon
        Status      = $AuditData.ResultStatus
        Severity    = $AuditData.Severity
        Workload    = $AuditData.Source
        Operation   = $Rec.Operations
        Category    = $AuditData.Category 
      }
      $Report.Add($ReportLine)
  }
}

$Report | Select-Object Timestamp, Action, Operation, User  | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
