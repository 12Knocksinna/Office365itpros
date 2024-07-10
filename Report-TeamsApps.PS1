# Report-TeamsApps.PS1
# Report specific apps install in teams
# Initially designed to help with the retirement of Office Connectors from Teams in August 2024

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-TeamsApps.PS1

Connect-MgGraph -NoWelcome -Scopes Directory.Read.All, Team.ReadBasic.All

Write-Host "Looking for teams to analyze..."
[array]$Teams = Get-MgTeam -All -PageSize 999 | Sort-Object DisplayName

If ($Teams) {
    Write-Host ("Found {0} teams - now analyzing their apps" -f $Teams.Count)
} Else {
    Write-Host "No teams found"
    Break
}   

[array]$TargetApps = "SharePoint News", "RSS", "Incoming Webhook", "Happy Communities"
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Team in $Teams) {
    $TeamName = $Team.DisplayName
    Write-Host "Analyzing team: $TeamName"
    $Apps = Get-MgTeamInstalledApp -TeamId $Team.id -ExpandProperty TeamsAppDefinition
    If ($Apps) {
        ForEach ($App in $Apps) {
            If ($App.TeamsAppDefinition.DisplayName -in $TargetApps) {
                $Report.Add([PSCustomObject]@{
                    TeamName = $TeamName
                    AppName = $App.TeamsAppDefinition.DisplayName
                    Description = $App.TeamsAppDefinition.ShortDescription
                    AppVersion = $App.TeamsAppDefinition.Version
                    AppState = $App.TeamsAppDefinition.PublishingState
                    AppId = $App.TeamsAppDefinition.Id
                })
            }
        }
    }
}

$Report | Out-GridView -Title "App connectors to check"


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
