# ReportMailboxSendPermissionsMailboxes.PS1
# Quick and simple script to generate a report of non-standard permissions applied to Exchange Online user and shared mailboxes
# Needs to be connected to Exchange Online PowerShell with an administrative account to run
# V1.0 16-Mar-2020
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportMailboxSendPermissionsMailboxes.PS1
CLS
Write-Host "Fetching mailboxes"
$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -ResultSize Unlimited -PropertySet All | Select DisplayName, UserPrincipalName, RecipientTypeDetails, GrantSendOnBehalfTo
If ($Mbx.Count -eq 0) { Write-Error "No mailboxes found. Script exiting..." -ErrorAction Stop } 
CLS
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
$ProgressDelta = 100/($Mbx.count); $PercentComplete = 0; $MbxNumber = 0
ForEach ($M in $Mbx) {
    $MbxNumber++
    $MbxStatus = $M.DisplayName + " ["+ $MbxNumber +"/" + $Mbx.Count + "]"
    Write-Progress -Activity "Checkiing Send permissions for mailbox" -Status $MbxStatus -PercentComplete $PercentComplete
    $PercentComplete += $ProgressDelta
    $Permissions = Get-ExoRecipientPermission -Identity $M.UserPrincipalName | ? {$_.Trustee -ne "NT AUTHORITY\SELF"}
    If ($Null -ne $Permissions) {
    # Grab each permission and output it into the report
       ForEach ($Permission in $Permissions) {
       $ReportLine  = [PSCustomObject] @{
           Mailbox     = $M.DisplayName
           UPN         = $M.UserPrincipalName
           Permission  = $Permission | Select -ExpandProperty AccessRights
           AssignedTo  = $Permission.Trustee
           MailboxType = $M.RecipientTypeDetails } 
         $Report.Add($ReportLine) }}

    # Check if this mailbox has granted Send on Behalf of permission to anyone
    If (![string]::IsNullOrEmpty($M.GrantSendOnBehalfTo)) {
       ForEach ($Permission in $M.GrantSendOnBehalfTo) {
       $ReportLine  = [PSCustomObject] @{
           Mailbox     = $M.DisplayName
           UPN         = $M.UserPrincipalName
           Permission  = "Send on Behalf Of"
           AssignedTo  = (Get-ExoRecipient -Identity $Permission).PrimarySmtpAddress
           MailboxType = $M.RecipientTypeDetails } 
         $Report.Add($ReportLine) }}
}

$Report | Sort -Property @{Expression = {$_.MailboxType}; Ascending= $False}, Mailbox | Export-CSV c:\temp\MailboxSendPermissions.csv -NoTypeInformation
Write-Host "All done." $Mbx.Count "mailboxes scanned. Report of send permissions available in c:\temp\MailboxSendPermissions.csv"
