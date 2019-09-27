# A script to report the email addresses for Teams channels that are mail-enabled
# V1.0 August 2019
Cls
# Define the values applicable for the application used to connect to the Graph (different for each tenant, app, and secret)
$AppId = "d716b32c-0edb-48be-9385-30a9cfd96155"
$TenantId = "b662313f-14fc-33a2-9a7a-d2e27f4f3478"
$AppSecret = 's_rkvIn1oZ1cNceUBvJ2or10rrIsb*:='

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

# Loop through each team to examine its channels and discover if any are email-enabled
$i = 0; $EmailAddresses = 0; $Report = @(); $ReportLine = $Null
ForEach ($Team in $Teams.Value) {
      $i++
      $ProgressBar = "Processing Team " + $Team.DisplayName + " (" + $i + " of " + $Teams.Value.Count + ")"
      Write-Progress -Activity "Checking Teams Information" -Status $ProgressBar -PercentComplete ($i/$Teams.Value.Count*100)
      # Get owners of the team
      $TeamOwners = Invoke-WebRequest -Method GET -Uri "$($uri)groups/$($team.id)/owners" -ContentType $ctype -Headers $headers | ConvertFrom-Json  
      If ($TeamOwners.Value.Count -eq 1) {$TeamOwner = $TeamOwners.Value.DisplayName}
      Else { # More than one team owner, so let's split them out and make the string look pretty
         $Count = 1
         ForEach ($Owner in $TeamOwners.Value) {
            If ($Count -eq 1) {  # First owner in the list
               $TeamOwner = $Owner.DisplayName
               $Count++ }
            Else { $TeamOwner = $TeamOwner + "; " + $Owner.DisplayName }
       }}                         
      # Fetch list of channels for the team
      $Channels = Invoke-WebRequest -Method GET -Uri "$($uri)teams/$($team.id)/channels" -ContentType $ctype -Headers $headers | ConvertFrom-Json
      #Loop through each channel and get its email address if set
      ForEach ($Channel in $Channels.Value) {
        If (-Not [string]::IsNullOrEmpty($Channel.Email)) {
            # Write-Host "Email address found" $Channel.Email
            $EmailAddresses++
            $ReportLine = [PSCustomObject][Ordered]@{
              Team                = $Team.DisplayName
              TeamEmail           = $Team.Mail
              Owners              = $TeamOwner
              Channel             = $Channel.DisplayName
              ChannelDescription  = $Channel.Description
              ChannelEmailAddress = $Channel.Email   }
            # And store the line in the report object
            $Report += $ReportLine }}
} 
$Report | Sort Team | Export-CSV C:\Temp\TeamsChannelsWithEmailAddress.Csv -NoTypeInformation
Write-Host $EmailAddresses "mail-enabled channels found. Details are in C:\Temp\TeamsChannelsWithEmailAddress.Csv"
 
