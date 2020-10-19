# FindCrucialSendAndSearchRecords.PS1
# Examples used in Chapter 21 of Office 365 for IT Pros.

$Records = Search-UnifiedAuditLog -StartDate "18-Oct-2020 12:30" -EndDate "20-Oct-2020" -ResultSize 5000 -Operations Send
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
$Records = Search-UnifiedAuditLog -StartDate "1-Oct-2020 12:30" -EndDate "20-Oct-2020" -ResultSize 5000 -Operations $Operations 


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
