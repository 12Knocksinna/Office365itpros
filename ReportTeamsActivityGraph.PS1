# ReportTeamsActivityGraph.PS1
# A sample script showing how to fetch per-team usage information from the Graph
# V1.0 4-Apr-2020
# V2.0 21-Aug-2023 - Introduce pagination to make sure all teams are fetched
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportTeamsActivityGraph.PS1
# 
# Note: Guest user activity is not recorded by the Graph - only tenant accounts are processed
# Needs the Reports.Read.All permission to get user login data

# Define the values applicable for the application used to connect to the Graph (change these for your tenant)
$AppId = "b716b32c-0edb-48be-9385-30a9cfd96155"
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppSecret = 's_rkvIn1oZ1cNceUBvJ2or1lrrIsb*:='

# Construct URI and body needed for authentication
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials" }

# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Unpack Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

# Base URL
$headers = @{Authorization = "Bearer $token"}

Write-Host "Finding Teams Activity Data from the Graph..."
$Uri = "https://graph.microsoft.com/beta/reports/getTeamsTeamActivityDetail(period='D90')?`$top=500&format=application/json"
[array]$TeamsData =  (Invoke-RestMethod -Uri $Uri -Headers $headers -Method Get)

If (!($TeamsData)) { 
    Throw "Can't find any Teams activity details - exiting" 
  } Else {
   [array]$TeamsDataProcess = $TeamsData.Value
}

If ($TeamsDataProcess[0].teamid -eq "00000000-0000-0000-0000-000000000000") {
   Write-Host "Team information is obfuscated. Change the setting in the Microsoft 365 admin center"
   Write-Host "if you want to see team names and identifiers." }

# Check to see if we have more teams usage data to fetch
$NextLink = $TeamsData.'@Odata.NextLink'
While ($Null -ne $NextLink) {
   [array]$TeamsData =  (Invoke-RestMethod -Uri $NextLink -Headers $headers -Method Get)
   $TeamsDataProcess += $TeamsData.Value
   $NextLink = $TeamsData.'@odata.NextLink' 
}

Write-Host ("Activity records for {0} teams found - reporting." -f $TeamsData.Value.Count)

$Report = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Team in $TeamsDataProcess) {
   $DaysSinceActive = (New-Timespan -Start ($Team.LastActivityDate -as [datetime]) -End ($Team.Reportrefreshdate -as [datetime])).Days
   $ReportLine  = [PSCustomObject] @{   
      Team            = $Team.teamName
      Privacy         = $Team.teamType
      TeamId          = $Team.teamId
      LastActivity    = Get-Date ($Team.lastActivityDate) -format dd-MMM-yyyy
      ReportPeriod    = $Team.Details.reportPeriod
      DaysSinceActive = $DaysSinceActive
      ActiveUsers     = $Team.Details.activeUsers
      Posts           = $Team.Details.postMessages
      Replies         = $Team.Details.replyMessages
      Urgent          = $Team.Details.urgentMessages
      Mentions        = $Team.Details.mentions
      Guests          = $Team.Details.guests
      ActiveChannels  = $Team.Details.activeChannels
      Reactions       = $Team.Details.reactions }
 $Report.Add($ReportLine)
} #end ForEach

$Report | Sort-Object DaysSinceActive | Out-GridView
$Report | Sort-Object Team | Export-CSV -NoTypeInformation c:\temp\TeamsActivityData.csv
Write-Host ("Data for {0} teams exported to c:\temp\TeamsActivityData.csv" -f $TeamsData.Value.Count)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
