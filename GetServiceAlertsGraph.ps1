# GetServiceAlertsGraph.ps1
# https://github.com/12Knocksinna/Office365itpros/blob/master/GetServiceAlertsGraph.ps1
# Define the values applicable for the application used to connect to the Graph
$AppId = "d716b32c-0edb-48be-9385-30a9cfd96155"
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppSecret = 's_rkvIn1oZ1cNceUBvJ2or1lrrIsb*:='

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
$uri = "https://graph.microsoft.com/beta/Security/Alerts"
$headers = @{Authorization = "Bearer $token"}

$Alerts = (Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get -ContentType "application/json").Value
$Report = [System.Collections.Generic.List[Object]]::new()
[String]$User
ForEach ($Alert in $Alerts) {
  $ExtraInfo = $Null
  Switch ($Alert.Title) {
   "Email messages containing phish URLs removed after delivery" {
       $User = $Alert.UserStates.UserPrincipalName[1]  }
    "User restricted from sending email" {
       $User = $Alert.UserStates.UserPrincipalName }
    "Data Governance Activity Policy" {
       $User = "N/A" }
    "Admin Submission Result Completed" {
       $User = $Alert.UserStates.UserPrincipalName[0] 
       $ExtraInfo = "Email from " + $Alert.UserStates.UserPrincipalName[1] + " reported for " + $Alert.UserStates.UserPrincipalName[2] }
    "Default" {
       $User = $Alert.UserStates.UserPrincipalName } 
  } # End Switch

If ([string]::IsNullOrEmpty($Alert.Description)) { $AlertDescription = "Office 365 alert" }
  Else { $AlertDescription = $Alert.Description } 

# Unpack comments
 [String]$AlertComments = $Null; $i = 0
 ForEach ($Comment in $Alert.Comments) {
   If ($i -eq 0) { 
      $AlertComments = $Comment; $i++ }
   Else { $AlertComments = $AlertComments + "; " + $Comment }
 }

Switch ($Alert.Status) {
   "newAlert"   { $Color = "ff0000"  }
   "inProgress" { $Color = "ffff00"  }
   "Default"    { $Color = "00cc00"  }
}

 $ReportLine = [PSCustomObject][Ordered]@{
   Title       = $Alert.Title
   Category    = $Alert.Category
   User        = $User
   Description = $AlertDescription
   Date        = Get-Date($Alert.EventDateTime) -format g
   Status      = $Alert.Status
   Severity    = $Alert.Severity
   ViewAlert   = $Alert.SourceMaterials[0]
   Comments    = $AlertComments
   ExtraInfo   = $ExtraInfo
   Color       = $Color }
 $Report.Add($ReportLine)   

} # End ForEach

$URI = "https://outlook.office.com/webhook/0b9313ca-5b39-43a9-bde3-e0cd4e6ca4e0@b662313f-14fc-43a2-9a7a-d2e27f4f3478/IncomingWebhook/dd85ea98300a4fc3ba28eb7724a224ad/eff4cd58-1bb8-4899-94de-795f656b4a18"
Write-Host "Posting about new alerts..."
ForEach ($Item in $Report) {
   If ($Item.Status -ne "resolved" ) {

   # Convert MessageText to JSON beforehand, if not the payload will fail.
     $MessageText = ConvertTo-Json $Item.Description
   #  Generate payload(s)          
     $Payload = @" 
{
    "@context": "https://schema.org/extensions",
    "@type": "MessageCard",
    "potentialAction": [
            {
            "@type": "OpenUri",
            "name": "More info",
            "targets": [
                {
                    "os": "default",
                    "uri": "$($Item.ViewAlert)"
                }
            ]
        },
     ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Status:",
                    "value": "$($Item.Status)"
                },
                {
                    "name": "User:",
                    "value": "$($Item.User)"
                },
            {
                    "name": "Date:",
                    "value": "$($Item.Date)"
                }
            ],
            "text": $($MessageText)
        }
    ],
    "summary": "$($Item.Description)",
    "themeColor": "$($Item.Color)",
    "title": "$($Item.Title)"
    }
"@

# Post to Teams via webhook
    $Status = (Invoke-RestMethod -uri $URI -Method Post -body $Payload -ContentType 'application/json; charset=utf-8')
    }
}
   
# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
