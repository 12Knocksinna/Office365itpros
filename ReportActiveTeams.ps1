# Script to report active Teams
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportActiveTeams.ps1
# For Chapter 13 of Office 365 for IT Pros
# Updated because Microsoft changed the location of the Teams compliance records
# This version uses the Get-UnifiedGroup cmdlet only instead of Get-Team because it avoids loading a second module
# Check that the right modules are loaded
$Modules = Get-Module
If ("ExchangeOnlineManagement" -notin  $Modules.Name) {Write-Host "Please connect to Exchange Online Management  before continuing...";break}

Write-Host "Fetching list of teams..."
# $Teams = (Get-Team | Sort DisplayName)
$Teams = Get-UnifiedGroup -Filter {ResourceProvisioningOptions -eq "Team"} -ResultSize Unlimited | `
   Select-Object Alias, ManagedBy, WhenCreated, GroupMemberCount, DisplayName, ExternalDirectoryObjectId | `
   Sort-Object DisplayName
Write-Host "Starting to process" $Teams.Count "teams"
$Count = 0; Clear-Host ; $GroupNumber = 0
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($T in $Teams) {
  $ActiveStatus = "Inactive"
  $GroupNumber++
  $ProgressBar = "Processing team " + $T.DisplayName + " (" + $GroupNumber + " of " + $Teams.Count + ")" 
  Write-Progress -Activity "Checking Teams for activity" -Status $ProgressBar -PercentComplete ($GroupNumber/$Teams.Count*100)
  # $G = Get-UnifiedGroup -Identity $T.GroupId | Select-Object Alias, ManagedBy, WhenCreated, GroupMemberCount, DisplayName, ExternalDirectoryObjectId
  $TeamsData = (Get-ExoMailboxFolderStatistics -Identity $T.ExternalDirectoryObjectId -FolderScope NonIpmRoot -IncludeOldestAndNewestItems | `
   Where-Object {$_.FolderType -eq "TeamsMessagesData"})
  If ($TeamsData.ItemsInFolder) {
       # Write-Host "Processing" $T.DisplayName
       If ($null -ne $TeamsData.OldestItemReceivedDate ) {
           $TimeSinceCreation = (New-TimeSpan -Start $TeamsData.OldestItemReceivedDate -End (Get-Date)).Days }
       Else {
           $TimeSinceCreation = "No compliance records found" }
       $Count++  
       $ChatCount = $TeamsData.ItemsInFolder
       $NewestChat = $TeamsData.NewestItemReceivedDate
  } #End if
# Calculate chats per day (only since Microsoft moved the compliance record location in October 2020)
  If ($ChatCount -eq 0) {
     Write-Host "No Teams compliance records found for” $T.DisplayName -foregroundcolor Red
     $ChatsPerDay = 0 
     $NewestChat = "N/A"
     $ChatCount = 0 } 
  Else {
     $ChatsPerDay = $ChatCount/$TimeSinceCreation
     $ChatsPerDay = [math]::round($ChatsPerday,2) }

  If ($ChatsPerDay -gt 0 -and $ChatsPerDay -le 2) { $ActiveStatus = "Moderate" }
  Elseif ($ChatsPerDay -gt 2 -and $ChatsPerDay -le 5) { $ActiveStatus = "Reasonable"}
  Elseif ($ChatPerDay -gt 5) { $ActiveStatus = "Heavy" }   
 
  $ReportLine = [PSCustomObject]@{
     Alias        = $T.Alias
     Name         = $T.DisplayName
     Owner        = $T.ManagedBy
     Members      = $T.GroupMemberCount
     "Date of team creation"  = $T.WhenCreated
     ChatCount    = $ChatCount
     LastChat     = $NewestChat
     "Days for compliance records"   = $TimeSinceCreation 
     ChatsPerDay  = $ChatsPerDay
     ActiveStatus = $ActiveStatus}
  $Report.Add($ReportLine)
}
Clear-Host
Write-Host $Count "of" $Teams.Count "have some Teams activity"
$Report | Group-Object ActiveStatus | Sort-Object Count -Descending | Format-Table Name, count
$Report | Export-CSV c:\temp\TeamsReport.csv -NoTypeInformation
# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
