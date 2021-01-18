# A script to report the email addresses for Teams channels that are mail-enabled
# V1.0 August 2019; V2.0 January 2020; V3.0 February 2020 (add pagination support)
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportTeamsChannelEmailAddresses.ps1

Cls
# Define the values applicable for the application used to connect to the Graph (these are specific to a tenant)
$AppId = "s716b32c-0edb-48be-9385-30a9cfd96155"
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3478"
$AppSecret = 'j_rkvIn1oZ1cNceUBvJ2or1lrrIsb*:='
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
$TeamsHash = @{}
$Teams.Value.ForEach( {
   $TeamsHash.Add($_.Id, $_.DisplayName) } )
$NextLink = $Teams.'@Odata.NextLink'
While ($NextLink -ne $Null) {
   $Teams = Invoke-WebRequest -Method GET -Uri $NextLink -ContentType $ctype -Headers $headers | ConvertFrom-Json
   $Teams.Value.ForEach( {
      $TeamsHash.Add($_.Id, $_.DisplayName) } )
   $NextLink = $Teams.'@odata.NextLink' }

# All teams found...
CLS
Write-Host "Processing" $TeamsHash.Count "Teams..."
# Loop through each team to examine its channels and discover if any are email-enabled
$i = 0; $EmailAddresses = 0; $Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report; $ReportLine = $Null
ForEach ($Team in $TeamsHash.Keys) {
      $i++
      $TeamId = $($Team); $TeamDisplayName = $TeamsHash[$Team]  #Populate variables to identify the team
      $ProgressBar = "Processing Team " + $TeamDisplayName + " (" + $i + " of " + $TeamsHash.Count + ")"
      Write-Progress -Activity "Checking Teams Information" -Status $ProgressBar -PercentComplete ($i/$TeamsHash.Count*100)
      Try { # Get owners of the team
       $TeamOwners = Invoke-WebRequest -Method GET -Uri "$($uri)groups/$($TeamId)/owners" -ContentType $ctype -Headers $headers | ConvertFrom-Json  
       If ($TeamOwners.Value.Count -eq 1) {$TeamOwner = $TeamOwners.Value.DisplayName}
       Else { # More than one team owner, so let's split them out and make the string look pretty
         $Count = 1
         ForEach ($Owner in $TeamOwners.Value) {
            If ($Count -eq 1) {  # First owner in the list
               $TeamOwner = $Owner.DisplayName
               $Count++ }
            Else { $TeamOwner = $TeamOwner + "; " + $Owner.DisplayName }
       }}}
     Catch {Write-Host "Unable to get owner information for team" $TeamDisplayName }                        
     
     Try {  # Fetch list of channels for the team
      $Channels = Invoke-WebRequest -Method GET -Uri "$($uri)teams/$($TeamId)/channels" -ContentType $ctype -Headers $headers | ConvertFrom-Json 
      ForEach ($Channel in $Channels.Value) {
       If (-Not [string]::IsNullOrEmpty($Channel.Email)) {
          $EmailAddresses++
          $ReportLine = [PSCustomObject][Ordered]@{
              Team                = $TeamDisplayName
              TeamEmail           = $Team.Mail
              Owners              = $TeamOwner
              Channel             = $Channel.DisplayName
              ChannelDescription  = $Channel.Description
              ChannelEmailAddress = $Channel.Email 
              TeamId              = $TeamId  }
            # And store the line in the report object
            $Report.Add($ReportLine) }}
     }
    Catch { Write-Host "Unable to fetch channels for" $Team.DisplayName }
} 
$Report | Sort Team | Export-CSV C:\Temp\TeamsChannelsWithEmailAddress.Csv -NoTypeInformation
Write-Host $EmailAddresses "mail-enabled channels found. Details are in C:\Temp\TeamsChannelsWithEmailAddress.Csv"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
