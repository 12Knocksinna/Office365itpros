# FindDeletedEmailAuditRecords.PS1
# Example of how to search the Office 365 audit log to find who deleted messages. The output file can be refined to focus on specific
# mailboxes
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindDeletedEmailAuditRecords.PS1

If ($Null -eq (Get-ConnectionInformation)) {
    Connect-ExchangeOnline
}
$StartDate = (Get-Date).AddDays(-90); $EndDate = (Get-Date) 
$Records = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "HardDelete", "SoftDelete" -ResultSize 5000) 
If (!($Records)) { Write-Host "No deletion records found."; break } 
Else { 
 Write-Host "Processing" $Records.Count "audit records..." 
 $Report = [System.Collections.Generic.List[Object]]::new() # Create output file  
 ForEach ($Rec in $Records) { 
    $AuditData = ConvertFrom-Json $Rec.Auditdata 
    If ($AuditData.ResultStatus -eq "PartiallySucceeded") {
        $EMailSubjects = "*** Not deleted by" + $AuditData.ClientInfoString + " ***" }
    Else {
        $EmailSubjects = $AuditData.AffectedItems.Subject -join ", " }
    $ReportLine = [PSCustomObject] @{ 
      TimeStamp          = Get-Date($AuditData.CreationTime) -format g 
      User               = $AuditData.UserId 
      Action             = $AuditData.Operation 
      Status             = $AuditData.ResultStatus 
      Mailbox            = $AuditData.MailboxOwnerUPN 
      "Message Subjects" = $EmailSubjects
      Folder             = $AuditData.Folder.Path.Split("\")[1] 
      Client             = $AuditData.ClientInfoString } 
    $Report.Add($ReportLine) }
  } 
$Report | Sort-Object Mailbox | Select Timestamp, Action, User, Mailbox, "Message Subjects", Folder | Out-GridView
