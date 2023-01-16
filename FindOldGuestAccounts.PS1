# FindOldGuestAccounts.PS1
# Script to find Guest User Accounts in an Office 365 Tenant that are older than 365 days and the groups they belong to
# V2.0 10-Oct-2022
# V3.0 13-Jan-2023 - Removed the Exchange Online management module and do everything with the Graph

# https://github.com/12Knocksinna/Office365itpros/blob/master/FindOldGuestAccounts.PS1

Connect-MgGraph  -Scopes AuditLog.Read.All, Directory.Read.All
Select-MgProfile Beta

# Set age threshold for reporting a guest account
[int]$AgeThreshold = 365
# Output report name
$OutputReport = "c:\Temp\OldGuestAccounts.csv"
# Get all guest accounts in the tenant
Write-Host "Finding Guest Accounts..."
[Array]$GuestUsers = Get-MgUser -Filter "userType eq 'Guest'" -All -Property signInActivity 
$Today = (Get-Date); $i = 0; $StaleGuests = 0; $Report = [System.Collections.Generic.List[Object]]::new()
# Loop through the guest accounts looking for old accounts 
CLS
ForEach ($Guest in $GuestUsers) {
# Check the age of the guest account, and if it's over the threshold for days, report it
   $AADAccountAge = ($Guest.CreatedDateTime | New-TimeSpan).Days
   $i++
   If ($AADAccountAge -gt $AgeThreshold) {
      $ProgressBar = "Processing Guest " + $Guest.DisplayName + " " + $AAdAccountAge + " days old " +  " (" + $i + " of " + $GuestUsers.Count + ")"
      Write-Progress -Activity "Checking Guest Account Information" -Status $ProgressBar -PercentComplete ($i/$GuestUsers.Count*100)
      $GroupNames = $Null
     
# Find what Microsoft 365 Groups the guest belongs to... if any
$NumberOfGroups = 0
$Uri = ("https://graph.microsoft.com/beta/users/{0}/memberOf/microsoft.graph.group?`$filter=groupTypes/any(a:a eq 'unified')&`$top=200&$`orderby=displayName&`$count=true" -f $Guest.Id)
[array]$Data = Invoke-MgGraphRequest -Uri $Uri
[array]$GuestGroups = $Data.Value

If ($GuestGroups) { 
   $GroupNames = $Data.Value.displayName -join ", " 
   $NumberOfGroups = $GuestGroups.count
} 

$DaysSinceSignIn = "N/A"; $Year = $Null
# Some Azure AD accounts without sign in data return a sign in date of Monday 1 January 0001 00:00:00
# which makes it difficult to assess how many days since the last sign in
If ($Guest.SignInActivity.LastSignInDateTime) {
   [datetime]$UserLastSignInDate = $Guest.SignInActivity.LastSignInDateTime
   $Year = (Get-Date($UserLastSignInDate) -format "yyyy") 
}

Switch ($Year) {
  $Null   {  
     [string]$UserLastLogonDate = "No recent sign in records found" 
     $DaysSinceSignIn = "N/A" }
  "0001" {
     [string]$UserLastLogonDate = "No recent sign in records found" 
     $DaysSinceSignIn = "N/A" }
  Default {
     [string]$UserLastLogonDate = Get-Date ($UserLastSignInDate) -format g 
     $DaysSinceSignIn = ($UserLastLogonDate | New-TimeSpan).Days }
  }

 $Staleness = "Acceptable"
 If (($AADAccountAge -ge 365) -and ($AADAccountAge -le 730) -and ($NumberOfGroups -eq 0)) {
    $Staleness = "Stale" 
 } Elseif ($AADAccountAge -gt 730 -and $NumberOfGroups -eq 0) {
    $Staleness = "Very Stale" }

 $ReportLine = [PSCustomObject][Ordered]@{
      UPN                  = $Guest.UserPrincipalName
      Name                 = $Guest.DisplayName
      Age                  = $AADAccountAge
      "Account created"    = $Guest.CreatedDateTime
      "Last sign in"       = $UserLastLogonDate 
      "Days since sign in" = $DaysSinceSignIn
      "Number of groups"   = $NumberOfGroups
      Groups               = $GroupNames 
      "Staleness"          = $Staleness 
      Id                   = $Guest.Id }      
     $Report.Add($ReportLine) 

 } #End if guest age met threshold
} #End Foreach guest

$Report | Sort Name | Export-CSV -NoTypeInformation $OutputReport
$PercentStale = ($Report.Count/$GuestUsers.Count).toString("P")
Write-Host ("Script complete. {0} guest accounts found aged over {1} days ({2} of {3} accounts). Output CSV file is in {4}" -f $Report.Count, $AgeThreshold, $PercentStale, $GuestUsers.count, $OutputReport)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
