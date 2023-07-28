# Report-MRMTagsOnFolders.PS1
# A script to scan an Exchange Online mailbox and report the retention setting for each mailbox

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

$DefaultTags = $Null
ForEach ($Tag in $Tags) {
    If ((Get-RetentionPolicyTag -Identity $Tag | Select-Object -ExpandProperty Type) -eq 'All') {
    $DefaultTags += $Tag }
}

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