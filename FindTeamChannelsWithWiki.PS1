# FindTeamChannelsWithWiki.PS1
# Quick and dirty script to find wiki tabs in Teams channels
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindTeamChannelsWithWiki.PS1
# V2.0 - Include check for content from wiki configuration

Connect-MgGraph -NoWelcome `
  -Scopes Directory.Read.All, Team.ReadBasic.All, TeamSettings.Read.All, ChannelSettings.Read.All, TeamsTab.Read.All

Write-Host "Finding Teams"
[array]$Teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All | Sort-Object DisplayName
If (!($Teams)) {Write-Host "For some reason, I can't find any Teams... exiting..."; break }

Clear-Host
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report
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
         # Debug line - uncomment it to see the channelid details as they are processed
         # Write-Host ("Processing tab {0} in channel {1} of team {2}" -f $Tab.DisplayName, $Channel.DisplayName, $Team.DisplayName) 
         ForEach ($Tab in $Tabs) {
           $WikiContent = $False; $TeamOwners = $null
           If ($Tab.TeamsApp.Id -eq "com.microsoft.teamspace.tab.wiki") { 
            # Check if the wiki configuration says that it has some content
             [array]$TabData = Get-MgTeamChannelTab -ChannelId $Channel.Id -TeamId $Team.Id -TeamsTabId $Tab.Id  -ExpandProperty TeamsApp
             $WikiContent = $TabData.Configuration.AdditionalProperties['hasContent']
            }  #End if Wiki tab
            Switch ($WikiContent) {
              $True     { 
                 # Get Team owners 
                 [array]$OwnerData = Get-MgGroupOwner -GroupId $Team.Id
                 $TeamOwners = $OwnerData.AdditionalProperties.displayName -Join ", "
                 $WikiCheck = "Wiki content exists"}
              $False    { $WikiCheck = "No content" }
              Default    { $WikiCheck = "No content" }
            }   
            $ReportLine = [PSCustomObject][Ordered]@{
              Team           = $Team.DisplayName
              Channel        = $Channel.DisplayName
             "Tab name"     = $Tab.DisplayName
              Owners         = $TeamOwners
              Check          = $WikiCheck
              TeamId         = $Team.Id    
              AppId          = $Tab.TeamsApp.Id
            }
         $Report.Add($ReportLine) } #End ForEach Tabs
         Start-Sleep -Milliseconds 50 # Brief pause before next channel
       } #End if Tabs
      } #End ForEach channel
     } #End If (archived check)
    Else { Write-Host "The" $Team.DisplayName "team is archived - no check done" }
    Start-Sleep -Milliseconds 100
} #End ForEach Team

$WikiReport = $Report | Where-Object {$_.AppId -eq "com.microsoft.teamspace.tab.wiki" -and $_.Check -eq "Wiki content exists"}
$WikiReport | Out-GridView
$WikiReport | Export-CSV -NoTypeInformation WikisToReview.CSV -Encoding UTF8 -Delimiter ";"
Write-Host ("All done. {0} Teams processed. The Wiki tab with content was found in {1} channels. CSV file generated in WikisToReview.csv" -f $Teams.count, $WikiReport.count)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
