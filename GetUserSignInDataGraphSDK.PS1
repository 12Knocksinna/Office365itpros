# GetUserSignInDataGraph.PS1
# A script to fetch user sign-in data from the Microsoft Graph using Graph SDK Cmdlets
# https://github.com/12Knocksinna/Office365itpros/blob/master/GetUserSignInDataGraphSDK.PS1
# 
Connect-MgGraph -Scopes User.Read.All -NoWelcome
Clear-Host
$CSVFile = "c:\Temp\ReportUserSignin.csv"
# Get User sign in data
Write-Host "Accessing the Graph to get user sign-in data..."
[array]$Users = Get-MgUser -All `
   -Property Id, Mail, displayName, UserPrincipalName, SignInActivity, CreatedDateTime, usertype 
$Report = [System.Collections.Generic.List[Object]]::new() 

ForEach ($User in $Users) {  
   [string]$LastSignInOutput = $Null; $DaysSinceSignIn = $Null
   If ($Null -ne $User.SignInActivity)     {
      $LastUserSignInDate = $User.SignInActivity.lastSignInDateTime
      If ($Null -ne $LastUserSignInDate) {
          $LastSignInOutput = Get-Date($LastUserSignInDate) -format g
          $DaysSinceSignIn = (New-TimeSpan $LastUserSignInDate).Days
      }
   } Else { #No sign in data for this user account
      $LastSignInOutput = "Never or > 180 days" 
      $DaysSinceSignIn = "N/A" 
   }
 
   $ReportLine  = [PSCustomObject] @{          
     UPN                = $User.UserPrincipalName
     DisplayName        = $User.DisplayName
     Email              = $User.Mail
     ObjectId           = $User.Id
     Created            = Get-Date($User.CreatedDateTime) -format g      
     LastSignIn         = $LastSignInOutput
     DaysSinceSignIn    = $DaysSinceSignIn
     UserType           = $User.UserType }
   $Report.Add($ReportLine) 
} # End ForEach

Write-Host ("All done. {0} accounts processed - output available in {1}" -f $users.count, $CSVFile)
$Report | Sort-Object UserType, DisplayName | Out-GridView
$Report | Export-CSV -NoTypeInformation $CSVFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
