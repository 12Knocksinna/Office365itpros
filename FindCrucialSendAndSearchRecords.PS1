# FindCrucialSendAndSearchRecords.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindCrucialSendAndSearchRecords.PS1
# Examples used in Chapter 21 of Office 365 for IT Pros.
$StartDate = (Get-Date).AddDays(-90); $EndDate = (Get-Date) 
$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 -Operations Send
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
If ($Records.count -gt 0) {
   ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.AuditData 
      $ReportLine = [PSCustomObject] @{
        TimeStamp   = Get-Date($AuditData.CreationTime) -format g
        User        = $AuditData.MailboxOwnerUPN
        Operation   = $AuditData.Operation
        Subject     = $AuditData.Item.Subject
        MessageId   = $AuditData.Item.InternetMessageId }
     $Report.Add($ReportLine) }
} # End if

$Operations = "SearchQueryInitiatedSharePoint", "SearchQueryInitiatedExchange"
$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 -Operations $Operations 
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
If ($Records.count -gt 0) {
   ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.AuditData 
     Switch ($AuditData.Operation) {
      "SearchQueryInitiatedSharePoint" { # SharePoint search
       $ReportLine = [PSCustomObject] @{
         TimeStamp   = Get-Date($AuditData.CreationTime) -format g
         User        = $AuditData.UserId
         Client      = $AuditData.QuerySource
         Search      = $AuditData.QueryText 
         Scenario    = $AuditData.ScenarioName }
       $Report.Add($ReportLine) }
      "SearchQueryInitiatedExchange" { # Exchange search event
        $ReportLine = [PSCustomObject] @{
         TimeStamp   = Get-Date($AuditData.CreationTime) -format g
         User        = $AuditData.UserId
         Client      = $AuditData.QuerySource
         Search      = $AuditData.QueryText
         Scenario    = $AuditData.ScenarioName }
       $Report.Add($ReportLine) }
    } # End Switch
   } # End For
} # End if

$Report | Format-Table TimeStamp, Client, Search, User

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
