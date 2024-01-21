# UpdateMailboxAuditing.PS1
# A script to update Office 365 E3 user and shared mailboxes and make sure that they are enabled for mailbox auditing
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateMailboxAuditing.PS1
# Check if we can run an Exchange Online cmdlet. If we can, go on, else connect to Exchange Online
If ($Null -eq (Get-ConnectionInformation)) {
  Connect-ExchangeOnline
}
# Connect to the Graph SDK
Connect-MgGraph -Scopes User.Read.All -NoWelcome
# GUID for Office 365 E3
$Office365E3 = "6fd2c87f-b296-42f0-b197-1e91e994b900"

$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
$ProgressDelta = 100/($Mbx.count); $PercentComplete = 0; $MbxNumber = 0; $SharedMailboxNumber = 0; $MbxUpdated = 0; $SharedMbxUpdated
Clear-Host
Write-Host "Finding accounts (mailboxes) with Office 365 E3 licenses..."
# Process mailboxes - Check Azure Active Directory to find accounts with Office 365 E3 licenses
[array]$Mbx = Get-MgUser -filter "assignedLicenses/any(s:s/skuId eq $Office365E3)" -All
# Loop through accounts, find if they have not been enabled by checking CustomAttribute6, and enable if needed
ForEach ($M in $Mbx) {
    $MbxNumber++
    $MbxStatus = $M.DisplayName + " ["+ $MbxNumber +"/" + $Mbx.Count + "]"
    Write-Progress -Activity "Checking mailbox" -Status $MbxStatus -PercentComplete $PercentComplete
    $PercentComplete += $ProgressDelta
    $MbxProps = (Get-ExoMailbox -Identity $M.UserPrincipalName -Properties CustomAttribute6, RecipientTypeDetails)
    If ($MbxProps.CustomAttribute6 -ne "Mailbox Auditing Enabled") {
       Set-Mailbox -Identity $M.UserPrincipalName -AuditEnabled $false
       Set-Mailbox -Identity $M.UserPrincipalName -AuditEnabled $True -AuditOwner @{Add="MailItemsAccessed"} -CustomAttribute6 "Mailbox Auditing Enabled"
       $MbxUpdated++
       $ReportLine  = [PSCustomObject] @{
           Mailbox         = $M.DisplayName
           UPN             = $M.UserPrincipalName
           Department      = $M.Department
           Country         = $M.Country
           AuditingEnabled = "Y"
           MailboxType     = $MbxProps.RecipientTypeDetails} 
         $Report.Add($ReportLine) }
}
# Now process shared mailboxes. These don't have a license, so we fetch them from Exchange Online and check
[array]$SharedMbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox -Properties CustomAttribute6 -Filter {CustomAttribute6 -eq $Null}
$ProgressDelta = 100/($SharedMbx.count); $PercentComplete = 0; $MbxNumber = 0
ForEach ($M in $SharedMbx) {
    $SharedMailboxNumber++
    $MbxStatus = $M.DisplayName + " ["+ $SharedMailboxNumber +"/" + $SharedMbx.Count + "]"
    Write-Progress -Activity "Checking shared mailbox" -Status $MbxStatus -PercentComplete $PercentComplete
    $PercentComplete += $ProgressDelta
    If ($M.CustomAttribute6 -ne "Mailbox Auditing Enabled") {
       Set-Mailbox -Identity $M.UserPrincipalName -AuditEnabled $True -CustomAttribute6 "Mailbox Auditing Enabled"
       $SharedMbxUpdated++
       $ReportLine  = [PSCustomObject] @{
           Mailbox         = $M.DisplayName
           UPN             = $M.UserPrincipalName
           Department      = "Shared Mailbox"
           AuditingEnabled = "Y"
           MailboxType     = $M.RecipientTypeDetails} 
         $Report.Add($ReportLine) }
}
Write-Host "All done!"
Write-Host "---------"
Write-Host ""
Write-Host "Mailbox auditing enabled for Office 365 E3 mailboxes:" $MbxUpdated
Write-Host "Mailbox auditing enabled for shared mailboxes       :" $SharedMbxUpdated
$Report | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
