# ReportPermissionsFolderLevel.PS1
# Uses Exchange Online Management module (REST)
# V1.0 26-Feb-2020  Office 365 for IT Pros https://gum.co/O365IT/
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportPermissionsFolderLevel.PS1

# Check that we're connected to the Exchange Online Management module...
$Modules = Get-Module | ? { $_.Name -eq "ExchangeOnlineManagement" }  
If ($Null -eq $Modules) {
    Write-Host "Please connect to the Exchange Online Management module before running this script!"; break }

# Find mailboxes
$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Select DisplayName, ExternalDirectoryObjectId, UserPrincipalName, Name
$Report = [System.Collections.Generic.List[Object]]::new() 
$ProgressDelta = 100/($Mbx.count); $PercentComplete = 0; $MbxNumber = 0

# Loop through mailboxes to fetch folders and procvess them
ForEach ($M in $Mbx) {
  $MbxNumber++
  $MbxStatus = $M.DisplayName + " ["+ $MbxNumber +"/" + $Mbx.Count + "]"
  Write-Progress -Activity "Processing mailbox" -Status $MbxStatus -PercentComplete $PercentComplete
  $PercentComplete += $ProgressDelta
  # Get Folders to check - not interested in system folders and focusing on folders that are usually delegated
  # We also don't check empty folders...
  $Folders = Get-ExoMailboxFolderStatistics -Identity $M.ExternalDirectoryObjectId | ? {$_.FolderType -eq "User Created" -or $_.FolderType -eq "Inbox" -or $_.FolderType -eq "SentItems" -or $_.FolderType -eq "Contacts" -or $_.FolderType -eq "Calendar" -and $_.ItemsInFolder -gt 0 } | Select Name, itemsinfolder
  # Loop through folder to extract pernmissions for each one  
  ForEach ($Folder in $Folders) {
    $FolderName = $M.Name + ":\" + $Folder.Name
    $Permissions = (Get-ExoMailboxFolderPermission -Identity $FolderName -ErrorAction SilentlyContinue) 
      ForEach ($Permission in $Permissions) { # Check each permission, ignoring ones we don't have any interest in
         $User = $Permission.User
         If (($User -ne "Default" -and $User -ne "Anonymous") -and ($User -ne $M.DisplayName -and $Permission.AccessRights -ne "None")) {
            $i = 1
            ForEach ($Access in $Permission.AccessRights) { # Break up the array of access rights into a formatted string for the report
               If ($i -eq 1) { $AccessRights = $Access; $i++ }  Else {$AccessRights = $AccessRights + "; " + $Access }}
            $ReportLine = [PSCustomObject]@{
                Mailbox    = $M.DisplayName
                Folder     = $Folder.Name
                Permission = $AccessRights
                Assignedto = $User }
            $Report.Add($ReportLine) }}
          # Debug line -uncomment if you want to see these messages
          # Write-Host "Mailbox:" $M.Name "Permission" $Permission.AccessRights "on" $Folder.Name "for" $User
}}

$Report | Export-CSV -NoTypeInformation c:\temp\FolderDelegatedPermissions.CSV
CLS
Write-Host "Finished processing" $Mbx.Count "mailboxes. The report file is in c:\temp\FolderDelegatedPermissions.CSV"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
