# Script to report active Teams
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportActiveTeams.ps1
# For Chapter 13 of Office 365 for IT Pros
# Updated because Microsoft changed the location of the Teams compliance records
Write-Host "Fetching list of teams..."
$Teams = (Get-Team)
Write-Host "Starting to process" $Teams.Count "teams"
$Count = 0
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($T in $Teams | Sort DisplayName) {
  $ActiveStatus = "Inactive"
  $G = Get-UnifiedGroup -Identity $T.GroupId | Select Alias, ManagedBy, WhenCreated, GroupMemberCount, DisplayName
  $TeamsData = (Get-ExoMailboxFolderStatistics -Identity $G.Alias -FolderScope NonIpmRoot -IncludeOldestAndNewestItems | ? {$_.FolderType -eq "TeamsMessagesData"})
  If ($TeamsData.ItemsInFolder) {
       Write-Host "Processing" $G.DisplayName
       $TimeSinceCreation = (Get-Date) - $TeamsData.CreationTime
       $Count++  
       $ChatCount = $TeamsData.ItemsInFolder
       $NewestChat = $TeamsData.NewestItemReceivedDate
       $ChatsPerDay = $ChatCount/$TimeSinceCreation.Days
       $ChatsPerDay = [math]::round($ChatsPerday,2)
  } #End if
  If ($TeamsData.ItemsInFolder -eq 0) {
     Write-Host "No Teams compliance records found for” $T.DisplayName -foregroundcolor Red
     $ChatsPerDay = 0 
     $NewestChat = "N/A"
     $ChatCount = 0 } 
  If ($ChatsPerDay -gt 0 -and $ChatsPerDay -le 2) { $ActiveStatus = "Moderate" }
  Elseif ($ChatsPerDay -gt 2 -and $ChatsPerDay -le 5) { $ActiveStatus = "Reasonable"}
  Elseif ($ChatPerDay -gt 5) { $ActiveStatus = "Heavy" } 
 
  $ReportLine = [PSCustomObject]@{
     Alias        = $G.Alias
     Name         = $T.DisplayName
     Owner        = $G.ManagedBy
     Members      = $G.GroupMemberCount
     WhenCreated  = $G.WhenCreated
     ChatCount    = $ChatCount
     LastChat     = $NewestChat
     DaysOld      = $TimeSinceCreation.Days 
     ChatsPerDay  = $ChatsPerDay
     ActiveStatus = $ActiveStatus}
  $Report.Add($ReportLine)
}
Write-Host $Count "of" $Teams.Count "have some Teams activity"
$Report | Group ActiveStatus | Sort Count -Descending | Format-Table Name, count
$Report | Export-CSV c:\temp\TeamsReport.csv -NoTypeInformation

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
