# https://github.com/12Knocksinna/Office365itpros/blob/master/FindTabsAndAppsInTeams.ps1
# Remember to insert the correct values for the tenant id, app id, and app secret before running the script.
Cls
# Define the values applicable for the application used to connect to the Graph
$AppId = "d716b32c-0edb-48be-9385-30a9cfd96153"    # Change this
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3476" # And this
$AppSecret = 's_rkvIn1oZ1cNceUBvJ2or1lrrIsb*:='    # and this

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

$Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $Token" 
            'ConsistencyLevel' = "eventual" }

# Create list of Teams in the tenant
Write-Host "Fetching list of Teams in the tenant"
$Uri = "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"
$Teams = Get-GraphData -AccessToken $Token -Uri $Uri

CLS
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report; $ReportLine = $Null
$i = 0
# Loop through each team to examine its channels and discover its channels, tabs, and apps
ForEach ($Team in $Teams) {
      $i++
      $ProgressBar = "Processing Team " + $Team.DisplayName + " (" + $i + " of " + $Teams.Value.Count + ")"
      $Uri = "https://graph.microsoft.com/v1.0/teams/" + $Team.Id
      $TeamDetails = Get-GraphData -AccessToken $Token -Uri $Uri
      If ($TeamDetails.IsArchived -ne $True) { # Team is not archived, so we can fetch information
       Write-Progress -Activity "Checking Teams Information" -Status $ProgressBar -PercentComplete ($i/$Teams.Value.Count*100)
       # Get apps installed in the team
       $Uri = "https://graph.microsoft.com/v1.0/teams/$($Team.id)/installedApps?`$expand=teamsApp"
       $TeamApps = Get-GraphData -AccessToken $Token -Uri $Uri
       $TeamAppNumber = 0
       ForEach ($App in $TeamApps) {
         $TeamAppNumber++
         $ReportLine = [PSCustomObject][Ordered]@{
            Record = "App"
            Number = $TeamAppNumber
            Team   = $Team.DisplayName
            TeamId = $Team.Id
            App    = $App.TeamsApp.DisplayName
            AppId  = $App.TeamsApp.Id
            Distribution = $App.TeamsApp.DistributionMethod }
         $Report.Add($ReportLine) }
            
      # Get the channels so we can report the tabs created in each channel
      $Uri = "https://graph.microsoft.com/beta/Teams/$($Team.id)/channels"
      $TeamChannels = Get-GraphData -AccessToken $Token -Uri $Uri
      # Find the tabs created for each channel (standard tabs like Files don't show up here)    
      ForEach ($Channel in $TeamChannels) {
        $Uri = "https://graph.microsoft.com/beta/teams/$($Team.id)/channels/$($channel.id)/tabs?`$expand=teamsApp"
        $Tabs = Get-GraphData -AccessToken $Token -Uri $Uri
        $TabNumber = 0
        ForEach ($Tab in $Tabs) {
            $TabNumber++
            $ReportLine = [PSCustomObject][Ordered]@{
              Record  = "Channel tab"
              Number  = $TabNumber
              Team    = $Team.DisplayName
              TeamId  = $TeamId
              Channel = $Channel.DisplayName
              Tab     = $Tab.DisplayName
              AppId   = $Tab.TeamsApp.Id
              Distribution = $Tab.TeamsApp.DistributionMethod 
              WebURL  = $Tab.WebURL}
         $Report.Add($ReportLine) }}
     } #End If (archived check)
    Else {
       Write-Host "The" $Team.DisplayName "team is archived - no check done" }
}


$Report | Sort Team, Record, App | Export-CSV C:\Temp\TeamsChannelsAppInfo.Csv -NoTypeInformation
Write-Host $EmailAddresses "Info about Teams channels, apps and tabs exported to C:\Temp\TeamsChannelsAppInfo.Csv"

function Get-GraphData {
# GET data from Microsoft Graph.
# Based on https://danielchronlund.com/2018/11/19/fetch-data-from-microsoft-graph-with-powershell-paging-support/
    param (
        [parameter(Mandatory = $true)]
        $AccessToken,

        [parameter(Mandatory = $true)]
        $Uri
    )

    # Check if authentication was successful.
    if ($AccessToken) {
        # Format headers.
        $Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $AccessToken" 
            'ConsistencyLevel' = "eventual"   }

        # Create an empty array to store the result.
        $QueryResults = @()

        # Invoke REST method and fetch data until there are no pages left.
        do {
            $Results = ""
            $StatusCode = ""

            do {
                try {
                    $Results = Invoke-RestMethod -Headers $Headers -Uri $Uri -UseBasicParsing -Method "GET" -ContentType "application/json" 

                    $StatusCode = $Results.StatusCode
                } catch {
                    $StatusCode = $_.Exception.Response.StatusCode.value__

                    if ($StatusCode -eq 429) {
                        Write-Warning "Got throttled by Microsoft. Sleeping for 45 seconds..."
                        Start-Sleep -Seconds 45
                    }
                    else {
                        Write-Error $_.Exception
                    }
                }
            } while ($StatusCode -eq 429)

            if ($Results.value) {
                $QueryResults += $Results.value
            }
            else {
                $QueryResults += $Results
            }

            $uri = $Results.'@odata.nextlink'
        } until (!($uri))

        # Return the result.
        $QueryResults
    }
    else {
        Write-Error "No Access Token"
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
