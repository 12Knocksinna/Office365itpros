# Find-InactiveTeamsChannels.PS1
# This script identifies inactive channels in Microsoft Teams based on the last activity date of a user message posted to the channel.

# V1.0 5-June-2026
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-InactiveTeamsChannels.PS1

Connect-MgGraph

# Threshold for inactivity in days
$ThresholdDays = 180
# Channel is deemed inactive if there has been no user activity for more than this number of days. Or we find no user messages in the channel.

Write-Host "Finding teams to check for inactive channels..." -ForegroundColor Green
Try {
    [array]$Teams = Get-MgTeam -All -ErrorAction Stop 
    Write-Host "Found $($Teams.Count) teams to check for inactive channels." -ForegroundColor Green
    $Teams = $Teams | Sort-Object DisplayName
} catch {
    Write-Host "Error fetching teams: $_" -ForegroundColor Red
    break
}

$Report = [System.Collections.Generic.List[Object]]::new()  
ForEach ($Team in $Teams) {
    Write-Host "Checking team: $($Team.DisplayName)" -ForegroundColor Cyan
    Try {
        [array]$Channels = Get-MgTeamChannel -TeamId $Team.Id -All `
            -Property Id, CreatedDateTime, DisplayName, Description, IsArchived, MembershipType -ErrorAction Stop 
        Write-Host "Found $($Channels.Count) channels in team: $($Team.DisplayName)." -ForegroundColor Green
        $Channels = $Channels | Sort-Object DisplayName
    } catch {
        Write-Host "Error fetching channels for team $($Team.DisplayName): $_" -ForegroundColor Red
        continue
    }

    ForEach ($Channel in $Channels) {
        $LastActivityDate = $null
        Write-Host "Checking channel: $($Channel.DisplayName)" -ForegroundColor Yellow
        Try {
            [array]$Messages = Get-MgTeamChannelMessage -TeamId $Team.Id -ChannelId $Channel.Id -Top 20 -ErrorAction Stop 
            if ($Messages.Count -gt 0) {
                [array]$UserMessages = $Messages | Where-Object {$_.body.content -ne "<systemeventmessage/>"}
                if ($UserMessages.Count -gt 0) {
                    $LastActivityDate = $UserMessages[0].CreatedDateTime
                    Write-Host "Last activity date for channel '$($Channel.DisplayName)' is: $LastActivityDate" -ForegroundColor Green
                } else {
                    Write-Host "No user messages found in channel '$($Channel.DisplayName)'. It may be inactive." -ForegroundColor Red
                    $ChannelStatus = "Inactive" 
                }
            } else {
                Write-Host "No messages found in channel '$($Channel.DisplayName)'. It may be inactive." -ForegroundColor Red
                $ChannelStatus = "Inactive"
            }
        } catch {
            Write-Host "Error fetching messages for channel '$($Channel.DisplayName)': $_" -ForegroundColor Red
            continue
        }
        Try {
            [array]$Members = Get-MgTeamChannelMember -TeamId $Team.Id -ChannelId $Channel.Id -All -ErrorAction Stop   
            [array]$Owners = $Members | Where-Object { $_.Roles -contains "owner" }
            [array]$Guests = $Members | Where-Object { $_.UserId -like "*#EXT#*" }
        } catch {
            Write-Host "Error fetching members for channel '$($Channel.DisplayName)': $_" -ForegroundColor Red
            continue
        }

        If ($LastActivityDate) { 
            [int]$DaysSinceLastActivity = (New-TimeSpan -Start $LastActivityDate -End (Get-Date)).Days
            If ($DaysSinceLastActivity -gt $ThresholdDays) {
                $ChannelStatus = "Inactive"
            } Else {
                $ChannelStatus = "Active"
            }
        } Else {
            [string]$DaysSinceLastActivity = $null
            $ChannelStatus = "Unknown"
        } 

        $ReportLine = [PSCustomObject][Ordered]@{
            TeamName        = $Team.DisplayName
            ChannelName     = $Channel.DisplayName
            ChannelStatus   = $ChannelStatus
            Description     = $Channel.Description
            LastActivityDate = [string]$LastActivityDate
            'Days since last activity' = [string]$DaysSinceLastActivity 
            IsArchived      = $Channel.IsArchived
            MembershipType  = $Channel.MembershipType
            OwnerCount      = $Owners.Count
            MemberCount     = $Members.Count
            GuestCount      = $Guests.Count
            TeamId          = $Team.Id
            ChannelId       = $Channel.Id
        }
        $Report.Add($ReportLine) | Out-Null
    }
}   



# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
