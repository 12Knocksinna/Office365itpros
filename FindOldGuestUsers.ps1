# Script to find Guest User Accounts in an Office 365 Tenant that are older than 365 days and the groups they belong to
# Find guest accounts
$GuestUsers = Get-AzureADUser -All $true -Filter "UserType eq 'Guest'"
$Today = (Get-Date); $StaleGuests = 0; $Report = @()
# Check each account and find those over 365 days old
ForEach ($Guest in $GuestUsers) {
   $AADAccountAge = ($Guest.RefreshTokensValidFromDateTime | New-TimeSpan).Days
   If ($AADAccountAge -gt 365) {
      $StaleGuests++
      Write-Host "Processing" $Guest.DisplayName
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
      $ReportLine = [PSCustomObject][Ordered]@{
           UPN     = $Guest.UserPrincipalName
           Name    = $Guest.DisplayName
           Age     = $AADAccountAge
           Created = $Guest.RefreshTokensValidFromDateTime  
           Groups  = $GroupNames
           DN      = $DN}      
     $Report += $ReportLine }
}
$Report | Sort Name | Export-CSV -NoTypeInformation c:\Temp\OldGuestAccounts.csv
