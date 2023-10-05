# ConvertMailContactsGuests.PS1
# Convert Exchange Online mail contacts to Azure AD guest accounts
# https://github.com/12Knocksinna/Office365itpros/blob/master/ConvertMailContactsGuests.PS1

Connect-ExchangeOnline
Connect-MgGraph -Scopes Directory.ReadWrite.All

[array]$Contacts = Get-ExoRecipient -RecipientTypeDetails MailContact -ResultSize Unlimited -Filter {CustomAttribute2 -ne "Migrated"} -PropertySets All
If (!($Contacts)) { Write-Host "No mail contacts found... " ; break }
Write-Host ("Found {0} mail contacts - now processing..." -f $Contacts.count)

Add-Type -AssemblyName 'System.Web'
# Get email addresses for current guest accounts
[array]$GuestEmail = Get-MgUser -All -Filter "userType eq 'Guest'"  | Sort-Object Mail | Select-Object -ExpandProperty Mail
[int]$i = 0
$DLUpdates = [System.Collections.Generic.List[Object]]::new() 

ForEach ($Contact in $Contacts) {
 $i++
 Write-Host ("Processing mail contact {0} ({1}/{2})" -f $Contact.PrimarySmtpaddress, $i, $Contacts.count) 
 If ($Contact.PrimarySmtpAddress -in $GuestEmail) { 
   Write-Host ("Contact {0} with email {1} is already registered as guest account - hiding mail contact" -f $Contact.DisplayName, $Contact.PrimarySmtpAddress)
   Set-MailContact -Identity $Contact.Alias -HiddenFromAddressListsEnabled $True -CustomAttribute2 "Migrated" 
  } Else {
  # Create a password for the new account   
  $NewPassword = [System.Web.Security.Membership]::GeneratePassword(10, 3)
  $NewPasswordProfile = @{}
  $NewPasswordProfile["Password"]= $NewPassword
  $NewPasswordProfile["ForceChangePasswordNextSignIn"] = $True
  # Determine usage location
  $UsageLocation = "US"
  Switch ($Contact.CountryOrRegion) {
   "Bulgaria"       { $UsageLocation = "BG" }
   "Canada"         { $UsageLocation = "CA" }
   "France"         { $UsageLocation = "FR" }
   "Germany"        { $UsageLocation = "DE" }
   "Ireland"        { $UsageLocation = "IE" }
   "Italy"          { $UsageLocation = "IT" }
   "Switzerland"    { $UsageLocation = "CH" }
   "United States"  { $UsageLocation = "US" }
   "United Kingdom" { $UsageLocation = "UK" }
  } #End Switch
  
  # New-MgUser gets upset if null strings are passed in parameters
  [string]$City = " "; [string]$Office = " "; [string]$JobTitle = " "; [string]$Department = " "
  [String]$Country = " "; [string]$PostalCode = " "; [string]$Company = " "; [string]$FirstName = "Unknown"
  [string]$LastName = "Unknown"; [string]$DisplayName = " "

  If ($Contact.City)            { $City = $Contact.City }
  If ($Contact.Office)          { $Office = $Contact.Office }
  If ($Contact.Title)           { $JobTitle = $Contact.Title }
  If ($Contact.CountryOrRegion) { $Country = $Contact.CountryOrRegion }
  If ($Contact.PostalCode)      { $PostalCode = $Contact.PostalCode }
  If ($Contact.Company)         { $Company = $Contact.Company }
  If ($Contact.Department)      { $Department = $Contact.Department }
  If ($Contact.FirstName)       { $FirstName = $Contact.FirstName }
  If ($Contact.LastName)        { $LastName = $Contact.LastName }
  If ($Contact.DisplayName)     { $DisplayName = $Contact.DisplayName }

  # Calculate values for mail nickname and user principal name for the guest account
  # Use your own domain rather than office365itpros.com....
  $Alias = $Contact.alias -replace '[?]',''
  $NickName = $Alias + ".Contact"
  $UPN = $NickName + "#EXT#@Office365itpros.com"

  # Give mail contact a different SMTP address so it doesn't clash
  $NewPrimarySmtpAddress = $NickName + ".temp@Office365itpros.com"
  Set-MailContact -Identity $Contact.Alias -EmailAddresses $NewPrimarySmtpAddress

  # Populate hash table with properties for the new account 
  $NewUserProperties = @{
    UserType = "Guest"
    GivenName = $FirstName
    Surname = $LastName
    DisplayName = $DisplayName
    JobTitle = $JobTitle
    Department = $Department
    MailNickname = $NickName
    Mail = $Contact.PrimarySmtpAddress
    UserPrincipalName = $UPN
    Country = $Country
    City  = $City
    PostalCode = $PostalCode
    OfficeLocation = $Office
    Company = $Company
    UsageLocation = $UsageLocation
    PasswordProfile = $NewPasswordProfile
    AccountEnabled = $true }

 # Try to create new guest account
 Try { 
    $NewGuestAccount = New-MgUser @NewUserProperties }
 Catch {
    Write-Host ("Couldn't create new Guest account using these properties")
    Write-Host $NewUserProperties 
    Break }

  # Let the new guest appear in Exchange address lists
  Update-MgUser -UserId $NewGuestAccount.Id -ShowInAddressList:$True
  # Hide the mail contact and keep a record of the old email address (that the guest account now has
  Set-MailContact -Identity $Contact.Alias -HiddenFromAddressListsEnabled $True -CustomAttribute1 $Contact.PrimarySmtpAddress -CustomAttribute2 "Migrated"

  # Update distribution groups...
  # Because Exchange Online doesn't create the new Mail User object immediately, we need to wait before we can
  # swap DL membership and replace the old mail contact records with the new guest accounts. So we write out the
  # information into a list and process the updates later
  $DN = $Contact.DistinguishedName
  [array]$DLs = Get-ExoRecipient -ResultSize Unlimited -Filter "Members -eq '$DN'" -RecipientTypeDetails MailUniversalDistributionGroup -ErrorAction SilentlyContinue
  If ($DLs) {
   Write-Host ("User is a member of {0} groups" -f $DLs.count)
   ForEach ($DL in $DLs) {
    $DataLine  = [PSCustomObject] @{
     DLName      = $DL.DisplayName
     DLAlias     = $DL.Alias
     DLId        = $DL.ExternalDirectoryObjectId
     DLOldMember = $DN
     DLNewMember = $NewGuestAccount.Id }
   $DLUpdates.Add($Dataline) }
  } #End If $DLs

 } #End if Not found in existing guest accounts
} #End ForEach Contact

# Export all the DL Updates to process
$DLUpdates | Export-CSV -NoTypeInformation c:\temp\DLUpdateToProcess.csv

# --- End of script

# This is the code needed to process the distribution list updates in the CSV file created by the code above.

[array]$DLUpdatesToProcess = Import-CSV c:\temp\DLUpdateToProcess.csv
ForEach ($Update in $DLUpdatesToProcess) {
     Remove-DistributionGroupMember -Identity $Update.DLAlias -Member $Update.DLOldMember -Confirm:$False
     Add-DistributionGroupMember -Identity $Update.DLAlias -Member $Update.DLNewMember
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
