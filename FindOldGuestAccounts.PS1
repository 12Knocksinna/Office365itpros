# FindOldGuestAccounts.PS1
# Script to find Guest User Accounts in an Office 365 Tenant that are older than 365 days and the groups they belong to
# V2.0 10-Oct-2022
# V3.0 13-Jan-2023 - Removed the Exchange Online management module and do everything with the Graph
# V3.1 9-Jul-2023    Updated for SDK V2.0
# V3.2 18-Sep-2024   Checked against SDK V2.22 and updated to use the SDK to find groups a guest belongs to

# https://github.com/12Knocksinna/Office365itpros/blob/master/FindOldGuestAccounts.PS1

Connect-MgGraph  -Scopes AuditLog.Read.All, Directory.Read.All -NoWelcome

# Set age threshold for reporting a guest account
[int]$AgeThreshold = 365
# Output report name
$OutputReport = "c:\Temp\OldGuestAccounts.csv"
# Get all guest accounts in the tenant
Write-Host "Finding Guest Accounts..."
[Array]$GuestUsers = Get-MgUser -Filter "userType eq 'Guest'" -All -PageSize 999 -Property id, displayName, userPrincipalName, userType, SignInActivity, createdDateTime
$i = 0; $Report = [System.Collections.Generic.List[Object]]::new()
# Loop through the guest accounts looking for old accounts 
Clear-Host
ForEach ($Guest in $GuestUsers) {

# Check the age of the guest account, and if it's over the threshold for days, report it
   $i++
   Write-Host ("Processing Guest {0} ({1} of {2})" -f $Guest.DisplayName, $i, $GuestUsers.Count)
   $AADAccountAge = ($Guest.CreatedDateTime | New-TimeSpan).Days

   If ($AADAccountAge -gt $AgeThreshold) {
      $ProgressBar = "Processing Guest " + $Guest.DisplayName + " " + $AAdAccountAge + " days old " +  " (" + $i + " of " + $GuestUsers.Count + ")"
      Write-Progress -Activity "Checking Guest Account Information" -Status $ProgressBar -PercentComplete ($i/$GuestUsers.Count*100)
      $GroupNames = $null; $RealGroups = $null
     
# Find what Microsoft 365 Groups the guest belongs to... if any
   [array]$GuestGroups = Get-MgUserMemberOf -UserId $Guest.id
#$Uri = ("https://graph.microsoft.com/V1.0/users/{0}/memberOf/microsoft.graph.group?`$filter=groupTypes/any(a:a eq 'unified')&`$top=200&$`orderby=displayName&`$count=true" -f $Guest.Id)
#[array]$Data = Invoke-MgGraphRequest -Uri $Uri
#[array]$GuestGroups = $Data.Value

   If ($GuestGroups) { 
      # Exclude administrative units
      $RealGroups = $GuestGroups | Where-Object {$_.additionalproperties.'@odata.type' -eq '#microsoft.graph.group'}
      $GroupNames = $RealGroups.additionalProperties.displayName -join ", " 
   } 

   $DaysSinceSignIn = "N/A"; $Year = $Null
# Some Entra ID accounts without sign in data return a sign in date of Monday 1 January 0001 00:00:00
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
      $Staleness = "Very Stale" 
   }

   $ReportLine = [PSCustomObject][Ordered]@{
      UPN                  = $Guest.UserPrincipalName
      Name                 = $Guest.DisplayName
      Age                  = $AADAccountAge
      "Account created"    = $Guest.CreatedDateTime
      "Last sign in"       = $UserLastLogonDate 
      "Days since sign in" = $DaysSinceSignIn
      "Number of groups"   = $RealGroups.Count
      Groups               = $GroupNames 
      "Staleness"          = $Staleness 
      Id                   = $Guest.Id }      
     $Report.Add($ReportLine) 
   } #End if guest age met threshold
} #End Foreach guest

$Report | Export-CSV -NoTypeInformation $OutputReport
$PercentStale = ($Report.Count/$GuestUsers.Count).toString("P")
Write-Host ("Script complete. {0} guest accounts found aged over {1} days ({2} of {3} accounts). Output CSV file is in {4}" -f $Report.Count, $AgeThreshold, $PercentStale, $GuestUsers.count, $OutputReport)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
