# Report-TeamsTags.PS1
# Report teams that have tags defined to tag members
# Doesn't support tags defined for shared and private channels yet...

$TenantId = 'Your tenant identifier'
# Identifier for the app holding the necessary permissions
$AppId = '1b58427d-1938-40de-9a5d-0b22c4f85c0c' 
# Thumbprint for an X.509 certificate uploaded to the app
$Thumbprint = "F79286DB88C21491110109A0222348FACF694CBD"

Connect-MgGraph -NoWelcome -TenantId $TenantId -AppId $AppId -CertificateThumbprint $Thumbprint
# Necessary scopes: Directory.Read.All, TeamworkTag.Read.All, Channel.ReadBasic.All, Team.ReadBasic.All
Write-Host "Finding Teams..."
[array]$Teams = Get-MgTeam -All | Sort-Object DisplayName
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Team in $Teams) {
  Write-Host "Processing team" $Team.DisplayName
  [array]$TeamTags = Get-MgTeamTag -TeamId $Team.Id -ErrorAction SilentlyContinue
  If ($TeamTags) { # The team has some tags
     ForEach ($Tag in $TeamTags) {
       $TagMembers = (Get-MgTeamTagMember -TeamId $Team.Id -TeamWorkTagId $Tag.Id) | Select-Object -ExpandProperty DisplayName 
       $ReportLine = [PSCustomObject][Ordered]@{
          Team        = $Team.DisplayName
          TeamId      = $Team.Id
          Tag         = $Tag.DisplayName
          Description = $Tag.Description
          Members     = $TagMembers -Join ", "
          TagId       = $Tag.Id }
       $Report.Add($ReportLine)
      } #End Foreach Tag
   } #End If TeamTags
} #End ForEach team
 
[array]$TeamsWithTags = $Report.TeamId | Sort-Object -Unique
[array]$UniqueTags = $Report.TagId | Sort-Object -Unique
Write-Host "Total teams:      " $Teams.Count
Write-Host "Teams with tags:  " $TeamsWithTags.Count
Write-Host "Total tags:       " $UniqueTags.Count

$Report | Select-Object Team, Tag, Description, Members | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.