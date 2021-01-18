# A script to find SendAs records in the Office 365 audit log and identify events belonging to user/shared mailboxes
# and those belonging to group mailboxes/Teams
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindSendAsAuditRecords.ps1
# V1.0 10 Feb 2020
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
# You might need to increase the number of retrieved records if your tenant generates lots of SendAs events
$Records = (Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date).AddDays(+1) -Operations "SendAs" -ResultSize 2000)
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
$Report | ? {$_.MailboxType -eq "UserMailbox"} | Out-GridView
$Report | Export-Csv -NoTypeInformation -Path c:\temp\SendASAuditRecords.csv
Write-Host "Report File saved in" c:\temp\SendASAuditRecords.csv

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
