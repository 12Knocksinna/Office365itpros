# Post-ItemstoHTTPWorkFlow.ps1
# An example of posting an adaptive card to the Power Automate 'Post to a channel when a webhook request is received' workflow
# V1.0 12-Jun-2024
# Github link: https://github.com/12Knocksinna/Office365itpros/blob/master/Post-ItemsToHTTPWorkFlow.PS1

# Define a protype adaptive card to use to post to the webhook
$Card = '{
       "type": "message",
       "attachments":[
          {
             "contentType":"application/vnd.microsoft.card.adaptive",
             "contentUrl":null,
             "content":{
                "$schema":"http://adaptivecards.io/schemas/adaptive-card.json",
                "type":"AdaptiveCard",
                "version":"1.4",
                "body":[
                    {
                        "type": "Image",
                        "url": "https://office365itpros.com/wp-content/uploads/2024/06/ProblemIndicator.jpg",
                        "height": "30px",
                        "altText": "Problem report"
                    },
                    {
                    "type": "TextBlock",
                    "text": "**MessageHeader**",
                    "style": "heading"
                    },
                    {
                    "type": "FactSet",
                    "facts": [
                       InsertFacts
                    ]
                }
       ]
}'

# Webhook to post the adaptive card to - each team channel will have a different webhook
$WebHook = "https://prod-161.westeurope.logic.azure.com:443/workflows/643e69f83c8944438d68119179a10a64/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=KODuebWbDGYFr0z0eu-6Rj8aUKz7108W3wrNJZxFE5A"

# Connect to the Microsoft Graph PowerShell SDK with the permission to read service health data
Connect-MgGraph -NoWelcome -Scopes ServiceHealth.Read.All

# Define the set of services we're interested in knowing about
[array]$ImportantServices = "Exchange", "Teams", "SharePoint", "OrgLiveID", "Planner", "microsoftteams", "O365Client", "OneDriveForBusiness"

# Grab the service health data
Write-Host "Retrieving service health data..."
$Uri = "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews"
[array]$Data = Invoke-MgGraphRequest -Method Get -Uri $Uri

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
    Switch ($Service.Status) {
        "ServiceOperational"    { $ServiceIndicator = "✅" }
        "ServiceRestored"       { $ServiceIndicator = "🟢" }
        "serviceDegradation"    { $ServiceIndicator = "🟡" }
        "ServiceInterruption"   { $ServiceIndicator = "🔴" }
        "ExtendedRecovery"      { $ServiceIndicator = "🟠" }
        "FalsePositive"         { $ServiceIndicator = "🟣" }
        "ServiceIncident"       { $ServiceIndicator = "🔵" }
        "ServiceRestoration"    { $ServiceIndicator = "🟢" }
        "ServiceMaintenance"    { $ServiceIndicator = "🔧" }
        "ServiceInfo"           { $ServiceIndicator = "ℹ️" }
        "ServiceWarning"        { $ServiceIndicator = "⚠️" }
        "ServiceUnknown"        { $ServiceIndicator = "❓" }
        Default                 { $ServiceIndicator = "❌" }
    }
    # Nicer names for the services
    Switch ($Service.Id) {
        "Exchange"              { $ServiceName = "Exchange Online" }
        "Teams"                 { $ServiceName = "Microsoft Teams" }
        "SharePoint"            { $ServiceName = "SharePoint Online" }
        "OrgLiveID"             { $ServiceName = "Entra ID" }
        "Planner"               { $ServiceName = "Microsoft Planner" }
        "microsoftteams"        { $ServiceName = "Microsoft Teams" }
        "O365Client"            { $ServiceName = "Office 365 Client" }
        "OneDriveForBusiness"   { $ServiceName = "OneDrive for Business" }
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
# Format the facts about service health to insert into the adaptive card
[string]$Facts = $null
[int]$Count = 1
ForEach ($R in $Report) {
   # Nicer names for the operational status
    Switch ($R.Status) {
        "ServiceOperational"    { $RStatus = "Operational" }
        "ServiceRestored"       { $RStatus = "Restored" }
        "serviceDegradation"    { $RStatus = "Degradation" }
        "ServiceInterruption"   { $RStatus = "Interruption" }
        "ExtendedRecovery"      { $RStatus = "Extended Recovery" }
        "FalsePositive"         { $RStatus = "False Positive" }
        "ServiceIncident"       { $RStatus = "Incident" }
        "ServiceRestoration"    { $RStatus = "Restoration" }
        "ServiceMaintenance"    { $RStatus = "Maintenance" }
        "ServiceInfo"           { $RStatus = "Information" }
        "ServiceWarning"        { $RStatus = "Warning" }
        "ServiceUnknown"        { $RStatus = "Unknown" }
        Default                 { $RStatus = "Unknown" }
    }
    # Generat5e the facts to insert into the adaptive card
    $RStatus = ("{0} {1}" -f $RStatus, $R.Indicator)
    $Fact = '{ "title": "' + $R.ServiceName + '",' + ' "value": "' + $RStatus + '" }'
    If ($Count -lt $Report.count) {
        $Fact = $Fact + ","
        $Count++
    }
    $Facts = $Facts + $Fact
}

# Create the message header
$MessageHeader = "Microsoft 365 Critical Service Status at " + (Get-Date).ToString("dd-MMM-yyyy HH:mm:ss") 
# Replace the fields in the protype card with the actual data
$Card = $Card.Replace("MessageHeader",$MessageHeader).Replace("InsertFacts",$Facts)

# And post the card to the webhook
Write-Host "Posting adaptive card to webhook..."
Try {
    Invoke-MgGraphRequest -uri $WebHook -Method Post -body $Card 
    Write-Host "Service health status successfully posted to Teams channel"
} Catch {
    Write-Host "Failed to post adaptive card to webhook: $_"
}   

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
