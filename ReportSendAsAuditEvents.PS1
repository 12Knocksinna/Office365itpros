# ReportSendAsAuditEvents.PS1
# First populate the Recipients Hash Table with user mailboxes, group mailboxes, and shared mailboxes
CLS
Write-Host "Populating Recipients Table..."
$RecipientsTable = @{}
Try {
    $Recipients = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox}
Catch {
    Write-Host "Can't find recipients" ; break}
# Now Populate hash table with label data  
$Recipients.ForEach( {
       $RecipientsTable.Add([String]$_.PrimarySmtpAddress, $_.RecipientTypeDetails) } )
# And include group mailboxes
$GroupMailboxes = Get-Mailbox -ResultSize Unlimited -GroupMailbox
$GroupMailboxes.ForEach( {
       $RecipientsTable.Add([String]$_.PrimarySmtpAddress, $_.RecipientTypeDetails) } )
Write-Host "Finding audit records for Send As operations..."
$Records = (Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date).AddDays(+1) -Operations "SendAs" -ResultSize 1000)
If ($Records.Count -eq 0) {
    Write-Host "No audit records for Send As found." }
Else {
    Write-Host "Processing" $Records.Count "Send As audit records..."
    $Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
    # Scan each audit record to extract information
    ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      $MailboxType = $RecipientsTable.Item($AuditData.MailboxOwnerUPN) # Look up hash table
      If ($MailboxType -eq "GroupMailbox") {$Reason = "Group Mailbox Send"} Else {$Reason = "Delegate Send As"}
      If ($AuditData.UserId -eq "S-1-5-18") {$UserId = "Service Account"} Else {$UserId = $AuditData.UserId}
      $ReportLine = [PSCustomObject] @{
           TimeStamp   = Get-Date($AuditData.CreationTime) -format g
           SentBy      = $AuditData.MailboxOwnerUPN
           SentAs      = $AuditData.SendAsUserSmtp
           Subject     = $AuditData.Item.Subject
           User        = $AuditData.UserId
           Action      = $AuditData.Operation
           Reason      = $Reason
           UserType    = $AuditData.UserType
           LogonType   = $AuditData.LogonType
           ClientIP    = $AuditData.ClientIP
           MailboxType = $MailboxType
           ClientInfo  = $AuditData.ClientInfoString
           Status      = $AuditData.ResultStatus }        
      $Report.Add($ReportLine) }
}
$Report | ? {$_.MailboxType -eq "UserMailbox"}"} | Out-GridView
