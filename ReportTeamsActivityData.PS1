# ReportTeamsActivityData.PS1
# Report the Teams usage data using the Microsoft Graph PowerShell SDK
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportTeamsActivityData.PS1

Connect-MgGraph -Scopes ReportSettings.ReadWrite.All
Select-MgProfile Beta

# first, find if the data is obscured
$Display = Invoke-MgGraphRequest -Method Get -Uri 'https://graph.microsoft.com/beta/admin/reportSettings'
If ($Display['displayConcealedNames'] -eq $True) { # data is obscured, so let's reset it to allow the report to run
   $ObscureFlag = $True
   Write-Host "Setting tenant data concealment for reports to False" -foregroundcolor red
   Invoke-MgGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/beta/admin/reportSettings' -Body (@{"displayConcealedNames"= $false} | ConvertTo-Json) }


$Uri = "https://graph.microsoft.com/beta/reports/getTeamsTeamActivityDetail(period='D180')?`$format=application/json&`$top=999"
[array]$Data = Invoke-MgGraphRequest -Uri $Uri -Method Get
$Data = $Data.Value

If ($Data.Count -gt 0) {
   $Report = [System.Collections.Generic.List[Object]]::new()
   ForEach ($Team in $Data) {
   $ReportLine = [PSCustomObject][Ordered]@{  
     Name             = $Team.teamName
     LastActivity     = $Team.lastActivityDate
     AccessType       = $Team.teamType
     Id               = $Team.teamId
     IsDeleted        = $Team.isDeleted
     ActiveUsers      = $Team.Details.activeUsers
     ActiveExtUsers   = $Team.Details.activeExternalUsers
     Guests           = $Team.Details.guests
     ActiveChannels   = $Team.Details.activeChannels
     SharedChannels   = $Team.Details.activeSharedChannels
     Posts            = $Team.Details.postMessages
     Replies          = $team.Details.replyMessages
     Channelmessages  = $Team.Details.channelMessages
     Reactions        = $Team.Details.reactions
     Mentions         = $Team.Details.mentions
     UrgentMessages   = $Team.Details.urgentMessages
     Meetings         = $Team.Details.meetingsOrganized
      }
   $Report.Add($ReportLine) }
}

$Report | Out-GridView

# And reset obscured data if necessary
If ($ObscureFlag -eq $True) {
   Write-Host "Resetting tenant data concealment for reports to True" -foregroundcolor red
   Invoke-MgGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/beta/admin/reportSettings' -Body (@{"displayConcealedNames"= $true} | ConvertTo-Json) }

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
  
