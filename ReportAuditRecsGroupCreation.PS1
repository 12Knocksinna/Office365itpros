# Example for Chapter 21 of Office 365 for IT Pros
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportAuditRecsGroupCreation.PS1
[array]$Records = (Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-180) -EndDate (Get-Date).AddDays(+1) -Operations "Add Group" -ResultSize 5000)
If ($Records.Count -eq 0) {
   Write-Host "No group creation records found." }
 Else {
   Write-Host "Processing" $Records.Count "audit records..."
$Report = [System.Collections.Generic.List[Object]]::new()
   ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      If ($AuditData.Actor[2].Id.SubString(0,5) -eq "User_") { # Dynamic group created from Azure AD portal
         $GroupId = $AuditData.Target.id[0]         
         $GroupName = $AuditData.Target.id[1] }
      Else {
         $GroupId = $AuditData.Target.Id[1]
         $GroupName = $AuditData.Target.Id[3] }
      $ReportLine = [PSCustomObject]@{
           TimeStamp   = Get-Date $AuditData.CreationTime -format g
           User        = $AuditData.UserId
           Action      = $AuditData.Operation
           Status      = $AuditData.ResultStatus
           Workload    = $AuditData.Actor[2].Id
           GroupId     = $GroupId
           GroupName   = $GroupName }
      $Report.Add($ReportLine) }}
$Report | Select Timestamp, Workload, User, GroupName -Unique

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
