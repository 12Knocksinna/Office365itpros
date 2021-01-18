# GetUserSignInDataGraph.PS1
# A script to fetch user sign-in data from the Microsoft Graph
# https://github.com/12Knocksinna/Office365itpros/blob/master/GetUserSignInDataGraph.PS1
#
CLS
# Define the values applicable for the application used to connect to the Graph (change these for your tenant)
$AppId = "d716b32c-0edb-48be-9385-30a9cfd96155"
$TenantId = "c662313f-14fc-43a2-9a7a-d2e27f4f3478"
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

# Get User sign in data
Write-Host "Accessing the Graph to get user sign-in data..."
$URI = "https://graph.microsoft.com/beta/users?`$select=displayName,userPrincipalName, mail, id, CreatedDateTime, signInActivity, UserType&`$top=999"
$SignInData = (Invoke-RestMethod -Uri $URI -Headers $Headers -Method Get -ContentType "application/json") 
$Report = [System.Collections.Generic.List[Object]]::new() 

Foreach ($User in $SignInData.Value) {  
   If ($Null -ne $User.SignInActivity)     {
      $LastSignIn = Get-Date($User.SignInActivity.LastSignInDateTime) -format g
      $DaysSinceSignIn = (New-TimeSpan $LastSignIn).Days }
   Else { #No sign in data for this user account
      $LastSignIn = "Never or > 180 days" 
      $DaysSinceSignIn = "N/A" }
     
   $ReportLine  = [PSCustomObject] @{          
     UPN                = $User.UserPrincipalName
     DisplayName        = $User.DisplayName
     Email              = $User.Mail
     ObjectId           = $User.Id
     Created            = Get-Date($User.CreatedDateTime) -format g      
     LastSignIn         = $LastSignIn
     DaysSinceSignIn    = $DaysSinceSignIn
     UserType           = $User.UserType }
   $Report.Add($ReportLine) 
} # End ForEach

# Do we have extra data to fetch?
$NextLink = $SignInData.'@Odata.NextLink'

While ($NextLink -ne $Null) { # We do... so process them.
   Write-Host "Still processing..."
   $SignInData = Invoke-WebRequest -Method GET -Uri $NextLink -ContentType "application/json" -Headers $Headers
   $SignInData = $SignInData | ConvertFrom-JSon
   ForEach ($User in $SignInData.Value) {  

   If ($Null -ne $User.SignInActivity)     {
      $LastSignIn = Get-Date($User.SignInActivity.LastSignInDateTime) -format g
      $DaysSinceSignIn = (New-TimeSpan $LastSignIn).Days }
   Else { #No sign in data for this user account
      $LastSignIn = "Never or > 180 days" 
      $DaysSinceSignIn = "N/A" }
     
   $ReportLine  = [PSCustomObject] @{  
     UPN                = $User.UserPrincipalName
     DisplayName        = $User.DisplayName
     Email              = $User.Mail
     ObjectId           = $User.Id
     Created            = Get-Date($User.CreatedDateTime) -format g      
     LastSignIn         = $LastSignIn
     DaysSinceSignIn    = $DaysSinceSignIn
     UserType           = $User.UserType        }
     $Report.Add($ReportLine) } 

   # Check for more data
   $NextLink = $SignInData.'@Odata.NextLink'
} # End While

Write-Host "All done. " $Report.Count "accounts processed - output available in c:\Temp\ReportUserSignin.csv."
$Report | Sort DisplayName | Out-GridView
$Report | Export-CSV -NoTypeInformation c:\Temp\ReportUserSignin.csv

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
