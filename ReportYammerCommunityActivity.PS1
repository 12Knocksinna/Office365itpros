# ReportYammerCommunityActivity.PS1
# Report the activity of Yammer communities (for networks configured in Microsoft 365 mode) using the compliance records stored in Exchange Online
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportYammerCommunityActivity.PS1

$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}

[array]$YammerGroups = Get-UnifiedGroup -ResultSize Unlimited |?{$_.GroupSku -eq "Yammer"}

If (!($YammerGroups)) {Write-Host "No Microsoft 365 Groups found for Yammer -exiting"; break}

$YammerData = [System.Collections.Generic.List[Object]]::new()
ForEach ($Group in $YammerGroups) {
  Write-Host "Processing" $Group.DisplayName
  $Folder = (Get-ExoMailboxFolderStatistics -Identity $Group.ExternalDirectoryObjectId -Folderscope NonIPMRoot -IncludeOldestAndNewestItems | ?{$_.FolderType -eq "Yammer"})
  If ($Folder.NewestItemReceivedDate) {
    $TimeSincePost = New-TimeSpan ($Folder.NewestItemReceivedDate)
    $FormattedTime = "{0:dd}d:{0:hh}h:{0:mm}m" -f $TimeSincePost }
  Else { 
      $FormattedTime = "N/A" }
  $ReportLine = [PSCustomObject][Ordered]@{  
     DisplayName = $Group.DisplayName
     Items       = $Folder.ItemsInFolder
     NewestItem  = $Folder.NewestItemReceivedDate
     TimeSincePost = $FormattedTime }
  $YammerData.Add($ReportLine)
} # End For
     
$YammerData | Sort Items -Descending | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
