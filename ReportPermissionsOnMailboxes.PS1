# ReportPermissionsOnMailboxes.PS1
# https://github.com/12Knocksinna/Office365itpros/edit/master/ReportPermissionsOnMailboxes.PS1
# Quick and simple script to generate a report of non-standard permissions applied to Exchange Online user and shared mailboxes
# Needs to be connected to Exchange Online PowerShell with an administrative account to run
# V1.0 25-Feb-2020
CLS
Write-Host "Fetching mailboxes"
$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -Properties RecipientTypeDetails -ResultSize Unlimited

If ($Mbx.Count -eq 0) { Write-Error "No mailboxes found. Script exiting..." -ErrorAction Stop } 
# We have some mailboxes, so we can process them...
CLS
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
$ProgressDelta = 100/($Mbx.count); $PercentComplete = 0; $MbxNumber = 0
ForEach ($M in $Mbx) {
    $MbxNumber++
    $MbxStatus = $M.DisplayName + " ["+ $MbxNumber +"/" + $Mbx.Count + "]"
    Write-Progress -Activity "Processing mailbox" -Status $MbxStatus -PercentComplete $PercentComplete
    $PercentComplete += $ProgressDelta
    # REST cmdlet equivalent
    # $Permissions = Get-ExoMailboxPermission -Identity $M.UserPrincipalName | ?  {$_.User -Like "*@*" }   
    $Permissions = Get-MailboxPermission -Identity $M.UserPrincipalName | ? {$_.User -Like "*@*" }    
    If ($Null -ne $Permissions) {
       # Grab each permission and output it into the report
       ForEach ($Permission in $Permissions) {
         $ReportLine  = [PSCustomObject] @{
           Mailbox    = $M.DisplayName
           UPN        = $M.UserPrincipalName
           Permission = $Permission | Select -ExpandProperty AccessRights
           AssignedTo = $Permission.User
           MailboxType = $M.RecipientTypeDetails } 
         $Report.Add($ReportLine) }
     } 
}     
$Report | Sort -Property @{Expression = {$_.MailboxType}; Ascending= $False}, Mailbox | Export-CSV c:\temp\MailboxPermissions.csv -NoTypeInformation
Write-Host "All done." $Mbx.Count "mailboxes scanned. Report of non-standard permissions available in c:\temp\MailboxPermissions.csv"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
