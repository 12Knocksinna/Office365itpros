# Report-MRMTagsOnFolders.PS1
# A script to scan an Exchange Online mailbox and report the retention setting for each mailbox
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-MRMTagsOnFolders.PS1

Connect-ExchangeOnline

$User = Read-Host "Enter name of user mailbox to examine"
$User = Get-ExoMailbox -Identity $User -ErrorAction SilentlyContinue -Properties RetentionPolicy
If (!($User)) { Write-Host ("Can't find mailbox for {0}" -f $User) ; break }

Write-Host ("Checking mailbox folders for {0}" -f $User.DisplayName)
[array]$MailboxFolders = Get-ExoMailboxFolderStatistics -Identity $User.UserPrincipalName | Where-Object {$_.FolderType -eq 'User created' -or $_.FolderType -eq 'Inbox' `
  -or $_.FolderType -eq 'SentItems' -or $_FolderType -eq 'DeletedItems' -or $_.FolderType -eq 'JunkEMail' `
  -or $_.FolderType -eq 'Contacts'} | Sort-Object Name

# Remove some of the unintreesting folders from the set found in the mailbox. There's no point
# processing these because users can't assign retention tags to the folders.
$MailboxFolders = $MailboxFolders | Where-Object {$_.Name -ne 'Social Activity Notifications'}
$MailboxFolders = $MailboxFolders | Where-Object {$_.Name -ne 'Clutter'}
$MailboxFolders = $MailboxFolders | Where-Object {$_.Name -ne 'Quick Step Settings'}
$MailboxFolders = $MailboxFolders | Where-Object {$_.Name -ne 'Suggested Contacts'}
$MailboxFolders = $MailboxFolders | Where-Object {$_.Name -ne 'SearchDiscoveryHoldsFolder'}
$MailboxFolders = $MailboxFolders | Where-Object {$_.Name -ne 'DiscoveryHolds'}
$MailboxFolders = $MailboxFolders | Where-Object {$_.Name -ne 'Conversation History'}
$MailboxFolders = $MailboxFolders | Where-Object {$_.Name -ne 'SearchDiscoveryHoldsUnindexedItemFolder'}

$MailboxFolders = $MailboxFolders | Where-Object {$_.ContainerClass -ne 'IPF.Note.SocialConnector.FeedItems'}

[array]$Tags = Get-RetentionPolicy $User.RetentionPolicy |Select-Object -ExpandProperty RetentionPolicyTagLinks

[array]$DefaultTags = $Null
ForEach ($Tag in $Tags) {
    If ((Get-RetentionPolicyTag -Identity $Tag | Select-Object -ExpandProperty Type) -eq 'All') {
    $DefaultTags += $Tag }
}

# Make a string of the default tags
$DefaultTagsNames = $DefaultTags -Join ", "
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Folder in $MailboxFolders) {
  If ($Folder.DeletePolicy.length -eq 0) {
     $DeletePolicy = $DefaultTagsNames 
  } Else {
     $DeletePolicy = $Folder.DeletePolicy
  }
  $ReportLine = [PSCustomObject][Ordered]@{
     Folder               = $Folder.Name
     Type                 = $Folder.FolderType
     Path                 = $Folder.FolderPath
     'MRM Policy'         = $DeletePolicy }
               
  $Report.Add($ReportLine) 
}
# And output...
$Report

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
