s# ReportTeamsPrivateChannels.PS1
# Report the set of Teams private channels and their members for a tenant
#
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportTeamsPrivateChannels.PS1
# Uses the MicrosoftTeams PowerShell module and the Microsoft Graph PowerShell SDK

$Modules = @( "MicrosoftTeams", "Microsoft.Graph" )
# Requires -Modules $Modules

$Outputfile = "C:\temp\OutputChannels.csv"
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "MicrosoftTeams")) {Write-Host "Please connect to the Microsoft Teams module and then restart the script"; break}

Connect-MgGraph -Scopes Group.Read.All, Directory.Read.All
Select-MgProfile Beta

[array]$Teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All
If (!($Teams)) {Write-Host "Can't find any teams - exiting"; break} 
$Teams = $Teams | Sort-Object DisplayName

$ChannelsList = [System.Collections.Generic.List[Object]]::new()
[int]$i = 0
ForEach ($Team in $Teams) {
   $i++
   Write-Host ("Processing {0} ({1}/{2})" -f $Team.DisplayName, $i, $Teams.Count)
   [array]$Channels = Get-TeamAllChannel -GroupId $Team.Id -MembershipType "Private"
   ForEach ($Channel in $Channels) {
     Write-Host ("Found private channel {0} in team {1}" -f $Channel.DisplayName, $Team.DisplayName)
     [array]$ChannelMembers = Get-TeamChannelUser -GroupId $Team.Id -DisplayName $Channel.DisplayName
     ForEach ($Member in $ChannelMembers) {
       $ChannelLine = [PSCustomObject][Ordered]@{  # Write out details of the private channel and its members
          Team                = $Team.DisplayName
          Channel             = $Channel.DisplayName
          Description         = $Channel.Description
          Member              = $Member.Name
          MemberUPN           = $Member.User
          Role                = $Member.Role
          HostTeam            = $Channel.HostTeamId
          Id                  = $Channel.Id }
       $ChannelsList.Add($ChannelLine) } 
  } #End Foreach Member
} # End ForEach Team

[array]$TeamsWithPrivateChannels = $ChannelsList | Sort-Object HostTeam -Unique | Select-Object -ExpandProperty Team
[array]$ChannelOwners = $ChannelsList | Where-Object {$_.Role -eq "Owner"}
$PrivateChannels = $ChannelOwners.Channel | Sort-Object -Unique
CLS
Write-Host "Analysis of Teams Private Channels Complete"
Write-Host "-------------------------------------------"
Write-Host ""
Write-Host ""
Write-Host "Total Teams processed:       " $Teams.Count
Write-Host "Teams with private channels: " $TeamsWithPrivateChannels.count
Write-Host ""
Write-Host "Private channels found in the following teams:" ($TeamsWithPrivateChannels -join ", ")
Write-Host ""
Write-Host "The owners of the private channels are:"
Write-Host ""
$ChannelOwners | Format-Table Member, Role, Channel -AutoSize

$ChannelsList | Out-GridView
$ChannelsList | Export-CSV -NoTypeInformation $Outputfile
Write-Host ""
Write-Host "Channels list is in $Outputfile"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
