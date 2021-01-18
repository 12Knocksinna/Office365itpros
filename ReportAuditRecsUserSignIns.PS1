# Example from Chapter 21 of Office 365 for IT Pros
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportAuditRecsUserSignIns.PS1
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
           $Status = "Succeeded"}
      Else
         { $Client = $AuditData.ExtendedProperties[0].Value
           $Status = $AuditData.ResultStatus  }
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

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
