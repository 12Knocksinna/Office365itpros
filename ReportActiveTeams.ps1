# Script to report active Teams
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportActiveTeams.ps1
# For Chapter 13 of Office 365 for IT Pros
Write-Host "Fetching list of teams..."
$Teams = (Get-Team)
Write-Host "Starting to process" $Teams.Count "teams"
$Count = 0
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($T in $Teams) {
  $ActiveStatus = "Inactive"
  $G = Get-UnifiedGroup -Identity $T.GroupId | Select Alias, ManagedBy, WhenCreated, GroupMemberCount, DisplayName
  $TeamsData = (Get-MailboxFolderStatistics -Identity $G.Alias -FolderScope ConversationHistory -IncludeOldestAndNewestItems)
  ForEach ($Folder in $TeamsData) { # We might have one or two subfolders in Conversation History; find the one for Teams      
  If ($Folder.FolderType -eq "TeamChat" -and $Folder.ItemsInFolder -gt 0) {
       Write-Host "Processing" $G.DisplayName
       $TimeSinceCreation = (Get-Date) - $Folder.CreationTime
       $Count++  
       $ChatCount = $Folder.ItemsInFolder
       $NewestChat = $Folder.NewestItemReceivedDate
       $ChatsPerDay = $ChatCount/$TimeSinceCreation.Days
       $ChatsPerDay = [math]::round($ChatsPerday,2)}
  If ($Folder.ItemsInFolder -eq 0 -and $Folder.FolderType -eq "TeamChat" ) {
     Write-Host "No Teams compliance records found for” $T.DisplayName -foregroundcolor Red
     $ChatsPerDay = 0 
     $NewestChat = "N/A"
     $ChatCount = 0 } }

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
