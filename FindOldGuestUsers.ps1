# FindOldGuestUsers.PS1
# Script to find Guest User Accounts in an Office 365 Tenant that are older than 365 days (update the $GuestAccountAge variable to set a different
# number of days to check for) and the groups they belong to
# Script needs to connect to Azure Active Directory and Exchange Online PowerShell.
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindOldGuestUsers.ps1
Write-Host "Finding Guest Users..."
$GuestAccountAge = 365 # Value used for guest age comparison. If you want this to be a different value (like 30 days), change this here.
$GuestUsers = Get-AzureADUser -All $true -Filter "UserType eq 'Guest'" | Sort DisplayName
$Today = (Get-Date); $StaleGuests = 0
$Report = [System.Collections.Generic.List[Object]]::new()
CLS
$ProgressDelta = 100/($GuestUsers.Count); $PercentComplete = 0; $GuestNumber = 0
# Check each account and find those over 365 days old
ForEach ($Guest in $GuestUsers) {
   $CreatedDate = ((Get-AzureADUser -ObjectId $Guest.UserPrincipalName).ExtensionProperty.createdDateTime)
   $AccountAge = ($CreatedDate | New-TimeSpan).Days
   If ($AccountAge -gt $GuestAccountAge) {
      $StaleGuests++; $GuestNumber++
      $CurrentStatus = $Guest.DisplayName + " ["+ $GuestNumber +"/" + $GuestUsers.Count + "]"
      Write-Progress -Activity "Extracting information for guest account" -Status $CurrentStatus -PercentComplete $PercentComplete
      $PercentComplete += $ProgressDelta
      $i = 0; $GroupNames = $Null
      # Find what Office 365 Groups the guest belongs to... if any
      $DN = (Get-Recipient -Identity $Guest.UserPrincipalName).DistinguishedName 
      $GuestGroups = (Get-Recipient -Filter "Members -eq '$Dn'" -RecipientTypeDetails GroupMailbox | Select DisplayName, ExternalDirectoryObjectId)
      If ($GuestGroups -ne $Null) {
         ForEach ($G in $GuestGroups) { 
           If ($i -eq 0) { $GroupNames = $G.DisplayName; $i++ }
         Else 
           {$GroupNames = $GroupNames + "; " + $G.DisplayName }
      }}
      $ReportLine = [PSCustomObject]@{
           UPN     = $Guest.UserPrincipalName
           Name    = $Guest.DisplayName
           Age     = $AccountAge
           Created = $CreatedDate
           Groups  = $GroupNames
           DN      = $DN}      
     $Report.Add($ReportLine) }
   Else { # Update the number of guests processed so our progress bar looks good
         $GuestNumber++
         $CurrentStatus = $Guest.DisplayName + " ["+ $GuestNumber +"/" + $GuestUsers.Count + "]"
         Write-Progress -Activity "Extracting information for guest account" -Status $CurrentStatus -PercentComplete $PercentComplete
         $PercentComplete += $ProgressDelta} 
}
$Report | Out-GridView
$Report | Sort Name | Export-CSV -NoTypeInformation c:\Temp\OldGuestAccounts.csv

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
