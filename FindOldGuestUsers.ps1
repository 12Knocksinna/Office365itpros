# FindOldGuestUsers.PS1
# Script to find Guest User Accounts in an Office 365 Tenant that are older than 365 days (update the $GuestAccountAge variable to set a different
# number of days to check for) and the groups they belong to
# Script needs to connect to the Microsoft Graph PowerShell SDK Exchange Online PowerShell.
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindOldGuestUsers.ps1
# V2.0 10-Oct-2022
# V2.1 19-Jul-2022 Updated for Graph SDK V2

Connect-MgGraph -Scopes AuditLog.Read.All, Directory.Read.All -NoWelcome

# Set age threshold for reporting a guest account
[int]$AgeThreshold = 365
# Output report name
$OutputReport = "c:\Temp\OldGuestAccounts.csv"
# Get all guest accounts in the tenant
Write-Host "Finding Guest Accounts..."
[Array]$GuestUsers = Get-MgUser -Filter "userType eq 'Guest'" -All -PageSize 999 -Property Id, displayName, userPrincipalName, createdDateTime, signInActivity `
    | Sort-Object displayName
$i = 0; $Report = [System.Collections.Generic.List[Object]]::new()
# Loop through the guest accounts looking for old accounts 
Clear-Host
ForEach ($Guest in $GuestUsers) {
# Check the age of the guest account, and if it's over the threshold for days, report it
   $AccountAge = ($Guest.CreatedDateTime | New-TimeSpan).Days
   $i++
   If ($AccountAge -gt $AgeThreshold) {
      $ProgressBar = "Processing Guest " + $Guest.DisplayName + " " + $AccountAge + " days old " +  " (" + $i + " of " + $GuestUsers.Count + ")"
      Write-Progress -Activity "Checking Guest Account Information" -Status $ProgressBar -PercentComplete ($i/$GuestUsers.Count*100)
      $StaleGuests++
      $GroupNames = $Null
      # Find what Microsoft 365 Groups the guest belongs to... if any
     [array]$GuestGroups = (Get-MgUserMemberOf -UserId $Guest.Id).additionalProperties.displayName
     If ($GuestGroups) {
        $GroupNames = $GuestGroups -Join ", " 
     } Else {
        $GroupNames = "None"
     }
  
#    Find the last sign-in date for the guest account, which might indicate how active the account is
     $UserLastLogonDate = $Null
     $UserLastLogonDate = $Guest.SignInActivity.LastSignInDateTime
     If ($Null -ne $UserLastLogonDate) {
        $UserLastLogonDate = Get-Date ($UserLastLogonDate) -format g
     } Else {
        $UserLastLogonDate = "No recent sign in records found" 
     }

     $ReportLine = [PSCustomObject][Ordered]@{
           UPN               = $Guest.UserPrincipalName
           Name              = $Guest.DisplayName
           Age               = $AccountAge
           "Account created" = $Guest.createdDateTime 
           "Last sign in"    = $UserLastLogonDate 
           Groups            = $GroupNames }     
     $Report.Add($ReportLine) }
} # End Foreach Guest

$Report |  Export-CSV -NoTypeInformation $OutputReport
$PercentStale = ($StaleGuests/$GuestUsers.Count).toString("P")
Write-Host ("Script complete. {0} guest accounts found aged over {1} days ({2} of total). Output CSV file is in {3}" -f $StaleGuests, $AgeThreshold, $PercentStale, $OutputReport)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
