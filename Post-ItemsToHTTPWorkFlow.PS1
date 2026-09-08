# Post-ItemstoHTTPWorkFlow.ps1
# An example of posting an adaptive card to the Power Automate 'Post to a channel when a webhook request is received' workflow
# V1.0 12-Jun-2024
# V1.1 8-Sep-2026 
# Github link: https://github.com/12Knocksinna/Office365itpros/blob/master/Post-ItemsToHTTPWorkFlow.PS1
# Article: https://office365itpros.com/2024/06/17/teams-post-to-channel-workflow/

# Define a protype adaptive card to use to post to the webhook
$CardTemplate = @'
{
  "type": "message",
  "attachments": [
    {
      "contentType": "application/vnd.microsoft.card.adaptive",
      "contentUrl": null,
      "content": {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "msteams": {
          "width": "Full"
        },
        "body": [
          {
            "type": "Image",
            "url": "https://office365itpros.com/wp-content/uploads/2024/06/ProblemIndicator.jpg",
            "height": "30px",
            "altText": "Problem report"
          },
          {
            "type": "TextBlock",
            "text": "**MessageHeader**",
            "size": "Large",
            "weight": "Bolder",
            "wrap": true
          },
          {
            "type": "FactSet",
            "facts": InsertFacts
          }
        ]
      }
    }
  ]
}
'@

# Webhook to post the adaptive card to - each team channel will have a different webhook. The webhook below is for illustrative purposes only and is invalid.
#You will need to create your own webhook for the channel you want to post to.
# Use the Workflows option to get a webhook for the 'Post to a channel when a webhook request is received' workflow in Power Automate
$WebHook = "https://prod-161.westeurope.logic.azure.com:443/workflows/643e69f83c8944438d68119179a10axx/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=KODuebWbDGYFr0z0eu-6Rj8aUKz7108W3wrNJZxFE5A"

# Connect to the Microsoft Graph PowerShell SDK with the permission to read service health data
Connect-MgGraph -NoWelcome -Scopes ServiceHealth.Read.All

# Define the set of services we're interested in knowing about
[array]$ImportantServices = "Exchange", "Teams", "SharePoint", "OrgLiveID", "Planner", "microsoftteams", "Purview", "OneDriveForBusiness", "Copilot"

# Grab the service health data
Write-Host "Retrieving service health data..."
$Uri = "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews"
Try {
    [array]$Data = Invoke-MgGraphRequest -Method Get -Uri $Uri -OutputType "PSObject" -ErrorAction Stop
} Catch {
    Write-Host "Failed to retrieve service health data: $_"
    return
}

# Extract the service health data for the services we're interested in
[array]$ImportantServiceStatus = $Data.Value | Where-Object {$_.Id -in $ImportantServices}

If (!($ImportantServiceStatus)) {
    Write-Host "No service health data found - exiting!"; break
}

# Create a PowerShell list to hold the service health data
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Service in $ImportantServiceStatus) {
    # add a graphic indicator based on the service status
    $ServiceIndicator = $null
    Switch ($Service.Status.tolower()) {
        "serviceoperational"    { $ServiceIndicator = "✅" }
        "servicerestored"       { $ServiceIndicator = "🟢" }
        "servicedegradation"    { $ServiceIndicator = "🟡" }
        "serviceinterruption"   { $ServiceIndicator = "🔴" }
        "extendedrecovery"      { $ServiceIndicator = "🟠" }
        "falsepositive"         { $ServiceIndicator = "🟣" }
        "serviceincident"       { $ServiceIndicator = "🔵" }
        "servicerestoration"    { $ServiceIndicator = "🟢" }
        "servicemaintenance"    { $ServiceIndicator = "🔧" }
        "serviceinfo"           { $ServiceIndicator = "ℹ️" }
        "servicewarning"        { $ServiceIndicator = "⚠️" }
        "serviceunknown"        { $ServiceIndicator = "❓" }
        Default                 { $ServiceIndicator = "❌" }
    }
    # Nicer names for the services
    $ServiceName = $Service.Id
    Switch ($Service.Id) {
        "Exchange"              { $ServiceName = "Exchange Online" }
        "Teams"                 { $ServiceName = "Microsoft Teams" }
        "SharePoint"            { $ServiceName = "SharePoint Online" }
        "OrgLiveID"             { $ServiceName = "Entra ID" }
        "Planner"               { $ServiceName = "Microsoft Planner" }
        "microsoftteams"        { $ServiceName = "Microsoft Teams" }
        "Purview"               { $ServiceName = "Microsoft Purview" }
        "OneDriveForBusiness"   { $ServiceName = "OneDrive for Business" }
        "Copilot"               { $ServiceName = "Microsoft 365 Copilot" }
        default                 { $ServiceName = $Service.Id }
    }
    # Update the list of services
    $ServiceHealth = [PSCustomObject] @{
        Service     = $Service.Id
        Status      = $Service.Status
        "Indicator" = $ServiceIndicator
        ServiceName = $ServiceName
    }
    $Report.Add($ServiceHealth)
}

# Sort by service name
$Report = $Report | Sort-Object ServiceName
# Format the facts about service health for the adaptive card
$Facts = ForEach ($R in $Report) {
    Switch ($R.Status.ToLowerInvariant()) {
        "serviceoperational"  { $RStatus = "Operational" }
        "servicerestored"     { $RStatus = "Restored" }
        "servicedegradation"  { $RStatus = "Degradation" }
        "serviceinterruption" { $RStatus = "Interruption" }
        "extendedrecovery"    { $RStatus = "Extended Recovery" }
        "falsepositive"       { $RStatus = "False Positive" }
        "serviceincident"     { $RStatus = "Incident" }
        "servicerestoration"  { $RStatus = "Restoration" }
        "servicemaintenance"  { $RStatus = "Maintenance" }
        "serviceinfo"         { $RStatus = "Information" }
        "servicewarning"      { $RStatus = "Warning" }
        "serviceunknown"      { $RStatus = "Unknown" }
        default               { $RStatus = "Unknown" }
    }

    [PSCustomObject]@{
        title = $R.ServiceName
        value = '{0} {1}' -f $RStatus, $R.Indicator
    }
}

$FactsJson = ConvertTo-Json -InputObject @($Facts) -Compress

# Create the message header
$MessageHeader = "Microsoft Critical Service Status\n$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')"
# Replace the fields in the protype card with the actual data
$Card = $CardTemplate.Replace("MessageHeader",$MessageHeader).Replace("InsertFacts",$FactsJson)

# And post the card to the webhook
Write-Host "Posting adaptive card to webhook..."
Try {
    Invoke-MgGraphRequest -uri $WebHook -Method Post -body $Card -ErrorAction Stop
    Write-Host "Service health status successfully posted to Teams channel"
} Catch {
    Write-Host "Failed to post adaptive card to webhook: $_"
}   

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.