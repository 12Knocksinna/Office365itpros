# Find-TargetedCollectionFolders.PS1
# Generate folder identifiers for use in a content search for specific folders in Recoverable Items in both the primary
# and archive mailbox
# https://github.com/12Knocksinna/Office365itpros/blob/master/Find-TargetedCollectionFolders.PS1

$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) { Write-Host "Please connect to the Exchange Online management module and restart." ; break }

$Target = Read-Host "Enter the user principal name of the mailbox to check"
[array]$Mbx = Get-ExoMailbox -Identity $Target -ErrorAction SilentlyContinue
If ($Mbx.count -ne 1) { Write-Host ("Sorry - can't find a mailbox for {0}" -f $Target) ; break }

$UserAccount = $Mbx.UserPrincipalName
$FolderQueries = @()
$ArchiveQueries = @()
$Encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
$Nibbler = $Encoding.GetBytes("0123456789ABCDEF")

Write-Host ("Checking primary mailbox for {0}" -f $UserAccount)
[array]$Folders = Get-ExoMailboxFolderStatistics -Identity $UserAccount -FolderScope RecoverableItems
If (!($Folders)) { Write-Host ("Unable to retrieve mailbox folder statistics for {0} - exiting" -f $UserAccount) ; break }

ForEach ($Folder in $Folders)  {
  $FolderPath = $Folder.FolderPath;
  If (($FolderPath -eq "/Versions") -or ($FolderPath -eq "/Deletions") -or ($FolderPath -eq "/Purges") -or ($FolderPath -eq "/DiscoveryHolds") -or ($FolderPath -eq "/SubstrateHolds")) {
   $FolderId = $Folder.FolderId      
   $FolderIdBytes = [Convert]::FromBase64String($folderId)
   $IndexIdBytes = New-Object byte[] 48
   $IndexIdIdx=0
   $FolderIdBytes | Select-object -skip 23 -First 24 | %{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
   $FolderQuery = "folderid:$($encoding.GetString($indexIdBytes))";
   $FolderDetails = New-Object PSObject
      Add-Member -InputObject $FolderDetails -MemberType NoteProperty -Name FolderPath -Value $FolderPath
      Add-Member -InputObject $FolderDetails -MemberType NoteProperty -Name FolderQuery -Value $folderQuery
      $FolderQueries += $FolderDetails
  } # End if
} # End Foreach

# Content search will always process an archive mailbox if one is available, so we need to fetch the identifiers for the target folders in the archive too
$ArchiveDatabase = Get-ExoMailbox -Identity $UserAccount -PropertySet Archive | Select-Object -ExpandProperty ArchiveDatabase
If ($ArchiveDatabase) { # We need to process archive folders too
   Write-Host ("Checking archive mailbox for {0}" -f $UserAccount)
   [array]$Folders = Get-ExoMailboxFolderStatistics -Identity $UserAccount -FolderScope RecoverableItems -Archive
   If (!($Folders)) { Write-Host ("Unable to retrieve archive mailbox folder statistics for {0}" -f $UserAccount) }
   ForEach ($Folder in $Folders)  {
    $FolderPath = $Folder.FolderPath;
    If (($FolderPath -eq "/Versions") -or ($FolderPath -eq "/Deletions") -or ($FolderPath -eq "/Purges") -or ($FolderPath -eq "/DiscoveryHolds") -or ($FolderPath -eq "/SubstrateHolds")) {
    $FolderId = $Folder.FolderId      
    $FolderIdBytes = [Convert]::FromBase64String($folderId)
    $IndexIdBytes = New-Object byte[] 48
    $IndexIdIdx=0
    $FolderIdBytes | Select-object -skip 23 -First 24 | %{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
    $FolderQuery = "folderid:$($encoding.GetString($indexIdBytes))";
    $FolderDetails = New-Object PSObject
      Add-Member -InputObject $FolderDetails -MemberType NoteProperty -Name FolderPath -Value $FolderPath
      Add-Member -InputObject $FolderDetails -MemberType NoteProperty -Name FolderQuery -Value $folderQuery
      $ArchiveQueries += $FolderDetails
  } # End if Folders
 } # End Foreach folders
} # End if Archive

Write-Host ""
Write-Host "Folder identifiers to use for content search"
Write-Host "--------------------------------------------"
Write-Host ""
Write-Host "Primary mailbox"
$FolderQueries | Format-Table
$KQLQuery = $FolderQueries.FolderQuery -join " OR "
Write-Host ""

If ($ArchiveDatabase)  {
  Write-Host "Archive mailbox"
  $ArchiveQueries | Format-Table
  $KQLQuery2 = $ArchiveQueries.FolderQuery -join " OR "
  $KQLQuery = $KQLQuery + " OR " + $KQLQuery2
}

Write-Host ("And here's the KQL query to insert into the content search: {0}" -f $KQLQuery)

Connect-IPPSSession
Write-Host "Creating the compliance search..."
$SearchName = "Focused Mailbox Search"
Remove-ComplianceSearch -Identity $SearchName -Confirm:$False -ErrorAction SilentlyContinue
New-ComplianceSearch -Name $SearchName -ContentMatchQuery $KQLQuery -Description ("Focused folder search for mailbox {0}" -f $UserAccount) -ExchangeLocation $UserAccount
Write-Host "Starting search"
Start-ComplianceSearch -Identity $SearchName
 Do {
       Write-Host ("Waiting for search {0} to comlete..." -f $SearchName)
       Start-Sleep -Seconds 5
       $ComplianceSearch = Get-ComplianceSearch -Identity $SearchName
   } While ($ComplianceSearch.Status -ne 'Completed')

Write-Host ("Search found {0} items in mailbox {1}" -f $ComplianceSearch.Items, $UserAccount)
Write-Host ""

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
