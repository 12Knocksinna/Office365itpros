# Report-TeamPrivateChannelMembership.PS1
# A script to report the membership of Teams private channels. This script is designed to be run by a tenant administrator or someone with the necessary permissions to access Teams data.

# Original idea described in https://office365itpros.com/2022/08/16/teams-private-channels-report/
# This version 7-May-2026
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-TeamPrivateChannelMembership.PS1

$Modules = Get-Module | Select-Object -ExpandProperty Name
If ($Modules -notcontains 'MicrosoftTeams') {
    Write-Host "Importing Microsoft Teams module..." -ForegroundColor Yellow
    Import-Module MicrosoftTeams -Force
    Try {
        Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Yellow
        Connect-MicrosoftTeams -ErrorAction Stop
    } Catch {
        Write-Host "Failed to connect to Microsoft Teams. Please ensure you have the necessary permissions and try again." -ForegroundColor Red
        Exit
    }
}

Write-Output "Fetching teams data..." -ForegroundColor Green
[array]$Teams = Get-Team 

$ChannelsList = [System.Collections.Generic.List[Object]]::new()
[int]$i = 0
ForEach ($Team in $Teams) {
    $i++
    Write-Host ("Processing {0} ({1}/{2})" -f $Team.DisplayName, $i, $Teams.Count)
    [array]$Channels = Get-TeamAllChannel -GroupId $Team.GroupId -MembershipType "Private"
    ForEach ($Channel in $Channels) {
        Write-Host ("Found private channel {0} in team {1}" -f $Channel.DisplayName, $Team.DisplayName)
        [array]$ChannelMembers = Get-TeamChannelUser -GroupId $Team.GroupId -DisplayName $Channel.DisplayName
        If ($ChannelMembers) {
            ForEach ($Member in $ChannelMembers) {
            $ChannelLine = [PSCustomObject][Ordered]@{ # Write out details of the private channel and its members
                Team                = $Team.DisplayName
                Channel             = $Channel.DisplayName
                Description         = $Channel.Description
                Member              = $Member.Name
                MemberUPN           = $Member.User
                Role                = $Member.Role
                HostTeam            = $Channel.HostTeamId
                Id                  = $Channel.Id 
            }
            $ChannelsList.Add($ChannelLine) 
            } # End ForEach Member
            
        }  Else { # Deal with private channels that have no members (which can happen if the channel is created and then all members are removed, or if the channel is created and never had any members added to it)
            $ChannelLine = [PSCustomObject][Ordered]@{ # Write out details of the private channel with no members
                Team                = $Team.DisplayName + " (No members)"
                Channel             = $Channel.DisplayName + " (No members)"
                Description         = $Channel.Description
                Member              = $Null
                MemberUPN           = $Null
                Role                = $Null
                HostTeam            = $Channel.HostTeamId
                Id                  = $Channel.Id 
            }
            $ChannelsList.Add($ChannelLine) 
        }
    }
} # End ForEach Team

$ChannelsList | Out-GridView -Title "Teams Private Channel Membership Report" 

# Set up to use the ImportExcel module to generate Excel worksheets if it is available. If not, the report will be generated in CSV format instead.
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
}

# Generate the attachment in either Excel worksheet or CSV format, depending on if the ImportExcel module is available
If ($ExcelGenerated) {
    $OutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\TeamsPrivateChannelMembership.xlsx"
    $ChannelsList | Export-Excel -Path $OutputFile -WorksheetName "PrivateChannelMembership" -Title ("Teams Private Channel Membership {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "PrivateChannels"         
} Else {
    $OutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\TeamsPrivateChannelMembership.csv"
    $ChannelsList | Export-Csv -Path $OutputFile     -NoTypeInformation -Encoding Utf8
}

Write-Host "Report generated in: $OutputFile" -ForegroundColor Green

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# The Automating Microsoft 365 with PowerShell eBook is available from https://o365itpros.gumroad.com/l/M365PS in EPUB and PDF format 
# or in paperback at https://amzn.to/3L4IBmU

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
