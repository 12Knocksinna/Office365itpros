# ReportMailboxPermissionsMailboxes.PS1
# Quick and simple script to generate a report of non-standard permissions applied to Exchange Online user and shared mailboxes
# Needs to be connected to Exchange Online PowerShell with an administrative account to run
# V1.0 16-Mar-2020
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportMailboxPermissionsMailboxes.PS1
CLS
Write-Host "Fetching mailboxes"
$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -ResultSize Unlimited -PropertySet Delivery -Properties RecipientTypeDetails, DisplayName | Select DisplayName, UserPrincipalName, RecipientTypeDetails, GrantSendOnBehalfTo
If ($Mbx.Count -eq 0) { Write-Error "No mailboxes found. Script exiting..." -ErrorAction Stop } 
CLS
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
$ProgressDelta = 100/($Mbx.count); $PercentComplete = 0; $MbxNumber = 0
ForEach ($M in $Mbx) {
    $MbxNumber++
    $MbxStatus = $M.DisplayName + " ["+ $MbxNumber +"/" + $Mbx.Count + "]"
    Write-Progress -Activity "Checking permissions for mailbox" -Status $MbxStatus -PercentComplete $PercentComplete
    $PercentComplete += $ProgressDelta
    $Permissions = Get-ExoRecipientPermission -Identity $M.UserPrincipalName | ? {$_.Trustee -ne "NT AUTHORITY\SELF"}
    If ($Null -ne $Permissions) {
    # Grab information about SendAs permission and output it into the report
       ForEach ($Permission in $Permissions) {
       $ReportLine  = [PSCustomObject] @{
           Mailbox     = $M.DisplayName
           UPN         = $M.UserPrincipalName
           Permission  = $Permission | Select -ExpandProperty AccessRights
           AssignedTo  = $Permission.Trustee
           MailboxType = $M.RecipientTypeDetails } 
         $Report.Add($ReportLine) }}

    # Grab information about FullAccess permissions
    $Permissions = Get-ExoMailboxPermission -Identity $M.UserPrincipalName | ? {$_.User -Like "*@*" }    
    If ($Null -ne $Permissions) {
       # Grab each permission and output it into the report
       ForEach ($Permission in $Permissions) {
         $ReportLine  = [PSCustomObject] @{
           Mailbox     = $M.DisplayName
           UPN         = $M.UserPrincipalName
           Permission  = $Permission | Select -ExpandProperty AccessRights
           AssignedTo  = $Permission.User
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

$Report | Sort -Property @{Expression = {$_.MailboxType}; Ascending= $False}, Mailbox | Export-CSV c:\temp\MailboxAccessPermissions.csv -NoTypeInformation
Write-Host "All done." $Mbx.Count "mailboxes scanned. Report of send permissions available in c:\temp\MailboxAccessPermissions.csv"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
