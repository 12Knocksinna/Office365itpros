Cls
# Define the values applicable for the application used to connect to the Graph
$AppId = "d716b32c-0edb-48be-9385-30a9cfd96154" # Will be different for your tenant
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f5f3478" # Will also be different for your tenant
$AppSecret = 's_rkvIn1oZ1cNceUBvJ2or1lrrIsb*:=' #And you'll need to get this too

# Construct URI and body needed for authentication
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"
}

# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Unpack Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

# Base URL
$uri = "https://graph.microsoft.com/beta/"
$headers = @{Authorization = "Bearer $token"}
$ctype = "application/json"

# Create list of Teams in the tenant
Write-Host "Fetching list of Teams in the tenant"
$Teams = Invoke-WebRequest -Method GET -Uri "$($uri)groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')" -ContentType $ctype -Headers $headers | ConvertFrom-Json

$Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report; $ReportLine = $Null
$i = 0
# Loop through each team to examine its channels and discover if any are email-enabled
ForEach ($Team in $Teams.Value) {
      $i++
      $ProgressBar = "Processing Team " + $Team.DisplayName + " (" + $i + " of " + $Teams.Value.Count + ")"
      Write-Progress -Activity "Checking Teams Information" -Status $ProgressBar -PercentComplete ($i/$Teams.Value.Count*100)
      # Get apps installed in the team
      $teamApps = Invoke-WebRequest -Headers $Headers -Uri "https://graph.microsoft.com/v1.0/teams/$($Team.id)/installedApps?`$expand=teamsApp" -ErrorAction Stop
     $teamApps = ($teamApps.content | ConvertFrom-Json).Value
     $TeamAppNumber = 0
     ForEach ($App in $TeamApps) {
         $TeamAppNumber++
         $ReportLine = [PSCustomObject][Ordered]@{
            Record = "App"
            Number = $TeamAppNumber
            Team   = $Team.DisplayName
            App    = $App.TeamsApp.DisplayName
            AppId  = $App.TeamsApp.Id
            Distribution = $App.TeamsApp.DistributionMethod }
        $Report.Add($ReportLine) }
            
# Get the channels so we can report the tabs created in each channel
 $TeamChannels = Invoke-WebRequest -Headers $Headers -Uri "https://graph.microsoft.com/beta/Teams/$($Team.id)/channels" -ErrorAction Stop
    $TeamChannels = ($TeamChannels.Content | ConvertFrom-Json).value
# Find the tabs created for each channel (standard tabs like Files don't show up here)    
ForEach ($Channel in $TeamChannels) {
        $Tabs = Invoke-WebRequest -Headers $Headers -Uri "https://graph.microsoft.com/beta/teams/$($Team.id)/channels/$($channel.id)/tabs?`$expand=teamsApp" 
        $Tabs = ($Tabs.Content | ConvertFrom-Json).value
        $TabNumber = 0
        ForEach ($Tab in $Tabs) {
            $TabNumber++
            $ReportLine = [PSCustomObject][Ordered]@{
              Record = "Channel tab"
              Number = $TabNumber
              Team   = $Team.DisplayName
              Channel = $Channel.DisplayName
              Tab    = $Tab.DisplayName
              AppId  = $Tab.TeamsApp.Id
              Distribution = $Tab.TeamsApp.DistributionMethod 
              WebURL  = $Tab.WebURL}
        $Report.Add($ReportLine) }}
}

$Report | Sort Team | Export-CSV C:\Temp\TeamsChannelsWithEmailAddress.Csv -NoTypeInformation
Write-Host $EmailAddresses "mail-enabled channels found. Details are in C:\Temp\TeamsChannelsWithEmailAddress.Csv"
