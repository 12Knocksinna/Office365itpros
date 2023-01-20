# FindTeamChannelsWithWiki.PS1
# Quick and dirty script to find wiki tabs in Teams channels
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindTeamChannelsWithWiki.PS1

Connect-MgGraph -Scopes Directory.Read.All, Team.ReadBasic.All, TeamSettings.Read.All, ChannelSettings.Read.All
Select-MgProfile beta  # Important because Get-MgTeam only returns all teams when connected to the beta endpoint

Write-Host "Finding Teams"
[array]$Teams = Get-MgTeam -All
If (!($Teams)) {Write-Host "For some reason, I can't find any Teams... exiting..."; break }
$Teams = $Teams | Sort-Object DisplayName
CLS
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report; $ReportLine = $Null
$i = 0
# Loop through each team to examine its channels, tabs, and apps
ForEach ($Team in $Teams) {
   $i++
   $ProgressBar = "Processing Team " + $Team.DisplayName + " (" + $i + " of " + $Teams.Count + ")"
    If ($TeamDetails.IsArchived -ne $True) { # Team is not archived, so we can fetch information
    Write-Progress -Activity "Checking Teams Information" -Status $ProgressBar -PercentComplete ($i/$Teams.Count*100)
      [array]$TeamChannels = Get-MgTeamChannel -TeamId $Team.Id
      ForEach ($Channel in $TeamChannels) {
        [array]$Tabs = Get-MgTeamChannelTab -ChannelId $Channel.Id -TeamId $Team.Id -ExpandProperty TeamsApp -ErrorAction SilentlyContinue
        If ($Tabs) {
         $TabNumber = 0
         # Debug line - uncomment it to see the channeld details as they are processed
         # Write-Host ("Processing tab {0} in channel {1} of team {2}" -f $Tab.DisplayName, $Channel.DisplayName, $Team.DisplayName) 
         ForEach ($Tab in $Tabs) {
            $ReportLine = [PSCustomObject][Ordered]@{
              Team           = $Team.DisplayName
              TeamId         = $Team.Id
              Channel        = $Channel.DisplayName
              "Channel tyoe" = $Channel.MembershipType
              "Tab name"     = $Tab.DisplayName
              AppId          = $Tab.TeamsApp.Id
              AppName        = $Tab.TeamsApp.DisplayName }
         $Report.Add($ReportLine) } #End ForEach Tabs
         Sleep -Milliseconds 50 # Brief pause before next channel
       } #End if Tabs
      } #End ForEach channel
     } #End If (archived check)
    Else { Write-Host "The" $Team.DisplayName "team is archived - no check done" }
} #End ForEach Team

$WikiReport = $Report | Where-Object {$_.AppName -eq "Wiki"}
$WikiReport | Out-GridView
Write-Host ("All done. {0} Teams processed. The Wiki tab was found in {1} channels" -f $Teams.count, $WikiReport.count)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
