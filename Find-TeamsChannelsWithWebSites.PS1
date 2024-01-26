# Find-TeamChannelsWithWebSites.PS1
# Quick and dirty script to find tabs pointing to websites in Teams channels
# https://github.com/12Knocksinna/Office365itpros/blob/master/Find-TeamsChannelsWithWebSites.PS1
# Provoked by change announced in MC708500 (20 Jan 2024)

Connect-MgGraph -NoWelcome `
  -Scopes Directory.Read.All, Team.ReadBasic.All, TeamSettings.Read.All, ChannelSettings.Read.All, TeamsTab.Read.All

Write-Host "Finding Teams"
[array]$Teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All | Sort-Object DisplayName
If (!($Teams)) {Write-Host "For some reason, I can't find any Teams... exiting..."; break }
$CSVOutputFile = "c:\temp\WebSitesinTeamsChannels.CSV"

# Clear-Host
Write-Host ("Processing {0} teams to check for tabs containing websites" -f $Teams.count)
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report
$i = 0
# Loop through each team to examine its channels, tabs, and apps
ForEach ($Team in $Teams) {
   $i++
   $ProgressBar = "Processing Team " + $Team.DisplayName + " (" + $i + " of " + $Teams.Count + ")"
    If ($TeamDetails.IsArchived -ne $True) { # Team is not archived, so we can fetch information
        Write-Progress -Activity "Checking Teams Information" -Status $ProgressBar -PercentComplete ($i/$Teams.Count*100)
        # Get Team owners 
        [array]$Owners = Get-MgGroupOwner -GroupId $Team.Id
        $TeamOwners = $Owners.AdditionalProperties.displayName -Join ", "
        [array]$TeamChannels = Get-MgTeamChannel -TeamId $Team.Id -ErrorAction SilentlyContinue
        ForEach ($Channel in $TeamChannels) {
            [array]$Tabs = Get-MgTeamChannelTab -ChannelId $Channel.Id -TeamId $Team.Id -ExpandProperty TeamsApp -ErrorAction SilentlyContinue
            If ($Tabs) {
            # Debug line - uncomment it to see the channelid details as they are processed
            # Write-Host ("Processing tab {0} in channel {1} of team {2}" -f $Tab.DisplayName, $Channel.DisplayName, $Team.DisplayName) 
            ForEach ($Tab in $Tabs) {
                $WebSiteURL = $null
                If ($Tab.TeamsApp.Id -eq "com.microsoft.teamspace.tab.web") { 
                    If ($Tab.WebURL) {
                        $StartPosition = $Tab.WebURL.IndexOf("webUrl=") + 7
                        # Some website tabs have &Label and some have ?Label, so we check for both
                        $LabelPosition =  $Tab.WebURL.IndexOf("&label")
                        If ($LabelPosition -eq -1) {
                            $LabelPosition = $Tab.WebURL.IndexOf("?label")
                        }
                        $URLLength = ($LabelPosition - ($StartPosition))
                        $WebSite = $Tab.WebURL.SubString($StartPosition, $URLLength)
                        $WebSiteURL = [System.Web.HttpUtility]::UrlDecode($WebSite)
                    } Else {
                        $WebSiteURL = "Can't find web site URL"
                    }
                    $ReportLine = [PSCustomObject][Ordered]@{
                        Team           = $Team.DisplayName
                        Channel        = $Channel.DisplayName
                        "Tab name"     = $Tab.DisplayName
                        URL            = $WebSiteURL
                        'Team owners'  = $TeamOwners
                        TeamId         = $Team.Id    
                        AppId          = $Tab.TeamsApp.Id
                    }
                    $Report.Add($ReportLine) 
                } 
            } #End ForEach Tab
            Start-Sleep -Milliseconds 250 # Brief pause before next channel
            } #End if Tabs
        } #End ForEach channel
    } Else { # End Archive check
        Write-Host ("The {0} team is archived - no check done" -f $Team.DisplayName)
    }
    Start-Sleep -Milliseconds 100
} #End ForEach Team

$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation $CSVOutputFile -Encoding UTF8 -Delimiter ";"
Write-Host ("All done. {0} Teams processed. A website tab was found in {1} channels. CSV file generated in {2}" -f $Teams.count, $Report.count, $CSVOutputFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.
