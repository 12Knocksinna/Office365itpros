# AnalyzeSendAsAuditData.ps1
# Script to analyze the audit records extracted for SendAs events against the permissions assigned in a tenant
# and report the permissions used and unused.
# V1.0 24-Mar-2020
# https://github.com/12Knocksinna/Office365itpros/blob/master/AnalyzeSendAsAuditData.ps1

If (!(Test-Path -Path c:\temp\MailboxAccessPermissions.csv)) {
   Write-Host "Error: c:\temp\MailboxAccessPermissions.csv not found. Please run script to generate report of mailbox SendAs permissions" ; break}
If (!(Test-Path -Path c:\temp\SendAsAuditRecords.CSV)) {
   Write-Host "Error: c:\temp\SendAsAuditRecords.CSV not found. Please run script to extract audit records for SendAs events" ; break}

# Import information gathered about audit records and permissions
$UserSendAsRecords = Import-CSV c:\temp\SendAsAuditRecords.CSV  | ? {$_.MailboxType -eq "UserMailbox"}
$SendAsData = Import-CSV c:\temp\MailboxAccessPermissions.csv | ? {$_.Permission -eq "SendAs"}
If ($SendAsData.Count -eq 0) { Write-Host "No SendAs permissions found"; break}

$PermissionsUsed = 0; $PermissionsNotUsed = 0
$PermissionsUsedReport = [System.Collections.Generic.List[Object]]::new()

# Check each assigned permission to establish if we can find an audit record to prove its use (or not)
ForEach ($S in $SendAsData) {
  $AuditCheck = $UserSendAsRecords | Where-Object {$_.SentBy -eq $S.AssignedTo -and $_.SentAs -eq $S.UPN} | Select -ExpandProperty TimeStamp
  If ($AuditCheck -eq $Null) {
       $PermissionsNotUsed++
       $ReportLine  = [PSCustomObject] @{
         Mailbox    = $S.Mailbox
         UPN        = $S.UPN
         Assignedto = $S.AssignedTo
         Status     = "SendAs permission not used"}
       $PermissionsUsedReport.Add($ReportLine) }
    Else {
       $LastUsedDate  = $AuditCheck | Sort-Object {$_.TimeStamp -as [datetime]} -Descending | Select -Last 1  # Grab latest SendAs
       $PermissionsUsed++
       $ReportLine  = [PSCustomObject] @{
         Mailbox    = $S.Mailbox
         UPN        = $S.UPN
         Assignedto = $S.AssignedTo
         Status     = [String]$AuditCheck.Count + " SendAs permissions used. Last use on " + $LastUsedDate }
       $PermissionsUsedReport.Add($ReportLine) }
}
    
$PermissionsUsedReport | Sort AssignedTo | Out-GridView
$PermissionsUsedReport | Export-CSV -NoTypeInformation c:\Temp\SendAsPermissionsUsageReport.CSV

CLS
Write-Host "SendAs Analysis for the last 90 days"
Write-Host "------------------------------------"
Write-Host "Total SendAs Permissions   :" $SendAsData.Count
Write-Host "SendAs permissions used    :" $PermissionsUsed
Write-Host "SendAs permissions not used:" $PermissionsNotUsed
Write-Host ""
Write-Host "Output report available in c:\Temp\SendAsPermissionsUsageReport.CSV"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization.  Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.

