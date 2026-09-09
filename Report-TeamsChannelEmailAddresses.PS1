# Report-TeamsChannelEmailAddresses.PS1
# Helper script to create a set of email addresses used by Teams channels
#
# V1.0 9-Sep-2026
# GitHub linl: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-TeamsChannelEmailAddresses.PS1
# Article link: https://office365itpros.com/2022/08/24/teams-channel-email-addresses/

Connect-MgGraph -Scopes GroupMember.Read.All, Channel.ReadBasic.All

Write-Output "Finding Teams..."
[array]$Teams = Get-MgTeam -All

If (!($Teams)) {
    Write-Host "Can't find any teams - exiting"
    Return
}

$Teams = $Teams | Sort-Object DisplayName
$ChannelsList = [System.Collections.Generic.List[Object]]::new()

# Loop through teams to find channels that have an email address
[int]$i = 0
ForEach ($Team in $Teams) {
    $i++
    $ProgressBar = "Processing Team " + $Team.DisplayName + " (" + $i + " of " + $Teams.Count + ")"
    Write-Progress -Activity "Checking Teams Information" -Status $ProgressBar -PercentComplete ($i/$Teams.Count*100)
    [array]$Channels = Get-MgTeamChannel -Teamid $Team.Id 
    ForEach ($Channel in $Channels) {
    # Check if the channel has an email address
    If ($Channel.Email) {
        $ChannelLine = [PSCustomObject][Ordered]@{  # Write out details of the private channel and its members
            DisplayName    = $Team.DisplayName
            Id             = $Team.Id
            Channel        = $Channel.DisplayName
            Description    = $Channel.Description
            EMail          = $Channel.Email
            'Channel Id'   = $Channel.Id 
        }
        $ChannelsList.Add($ChannelLine) } 
    } #End Foreach Channel
} # End ForEach Team

[array]$TeamsEmailAddresses = $ChannelsList | Where-Object {$_.Email -Like "*.teams.ms"}
Write-Output ("Found {0} channels enabled for inbound email" -f $TeamsEmailAddresses.count)

$OutputCSVFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Teams Channels with Email Addresses.csv"
Remove-Item -Path $OutputCSVFile -ErrorAction SilentlyContinue
$TeamsEmailAddresses | Export-Csv -Path $OutputCSVFile -NoTypeInformation -Encoding Utf8

Write-Output ("Exported Teams channels with email addresses to " + $OutputCSVFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.