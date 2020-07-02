# ArchiveMicrosoft365Groups.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ArchiveMicrosoft365Groups.PS1
# A script to process a set of Microsoft 365 Groups and Archive them so that the information is kept and available
# online for eDiscovery but can't be accessed nby users
# Uses the Exchange Online Module. Must be run by a tenant admin.

Write-Host "Preparing to archive" $ArchiveGroups.Count "Microsoft 365 Groups"
# Find list of groups to be archived
$ArchiveGroups = Get-UnifiedGroup -Filter {CustomAttribute1 -eq "Archive"} -ResultSize 1000 | Select DisplayName, DistinguishedName, Alias, PrimarySmtpAddress, SensitivityLabel

If ($ArchiveGroups.Count -eq 0) {
    Write-Host "No groups found to archive"; break}

# This is the address of the account that will continue to access the group
$AdminAccount = "Rory.Best@Office365itpros.com"
# If you use sensitivity labels, we need to define a label to assign to the archived groups
$SensitivityLabel = "27451a5b-5823-4853-bcd4-2204d03ab477"

$ProgressDelta = 100/($ArchiveGroups.Count); $PercentComplete = 0; $GroupNumber = 0
$Report = [System.Collections.Generic.List[Object]]::new()

CLS
# Main Loop
Foreach ($Group in $ArchiveGroups) {
  $GroupNumber++
  $CurrentStatus = $Group.DisplayName + " ["+ $GroupNumber +"/" + $ArchiveGroups.Count + "]"
  Write-Progress -Activity "Archiving Microsoft 365 Group" -Status $CurrentStatus -PercentComplete $PercentComplete
  $PercentComplete += $ProgressDelta

# Get lists of current owners and members
  $CurrentOwners = (Get-UnifiedGroupLinks -Identity $Group.DistinguishedName -LinkType Owners | Select Name)
  $CurrentMembers = (Get-UnifiedGroupLinks -Identity $Group.DistinguishedName -LinkType Members | Select Name)
# Add the administrator as an owner
  Add-UnifiedGroupLinks -Identity $Group.DistinguishedName -LinkType Members -Links $AdminAccount
  Add-UnifiedGroupLinks -Identity $Group.DistinguishedName -LinkType Owners -Links $AdminAccount
# Remove the other members and owners
  ForEach ($Owner in $CurrentOwners) { 
        Remove-UnifiedGroupLinks -Identity $Group.DistinguishedName -LinkType Owners -Links $Owner.Name -Confirm:$False }
  ForEach ($Member in $CurrentMembers) { 
        Remove-UnifiedGroupLinks -Identity $Group.DistinguishedName -LinkType Members -Links $Member.Name -Confirm:$False }

# Create SMTP Address for the archived group
  $OldSmtpAddress = $Group.PrimarySmtpAddress # Just for reporting
  $SmtpAddress = $Group.PrimarySmtpAddress -Split "@"
  $NewSmtpAddress = $SmtpAddress[0] +  "_archived" + "@" + $SmtpAddress[1]
  $AddressRemove = "smtp:"+ $Group.PrimarySmtpAddress

# Update the archive info for the group… $O365Cred is a credentials object that we fetch 
# the username from. Adjust the script for your own credentials.
  $ArchiveInfo = "Archived " + (Get-Date)  + " by " + $O365cred.username 
  $NewDisplayName = $Group.DisplayName + " (Archived)" 
# Update Group properties
  If ($Group.SensitivityLabel -eq $Null) {
      Set-UnifiedGroup -Identity $Group.DistinguishedName -AccessType Private -RequireSenderAuthenticationEnabled `
         $True -HiddenFromExchangeClientsEnabled -CustomAttribute1 $ArchiveInfo -PrimarySmtpAddress $NewSmtpAddress `
         -DisplayName $NewDisplayName -HiddenFromAddressListsEnabled $True }
  Elseif ($Group.SensitivityLabel -ne $Null) {
      Set-UnifiedGroup -Identity $Group.DistinguishedName -RequireSenderAuthenticationEnabled `
         $True -HiddenFromExchangeClientsEnabled -CustomAttribute1 $ArchiveInfo -PrimarySmtpAddress $NewSmtpAddress `
         -DisplayName $NewDisplayName -HiddenFromAddressListsEnabled $True `
         -SensitivityLabel $SensitivityLabel }

# Update Group email address
  Set-UnifiedGroup -Identity $Group.DistinguishedName -EmailAddresses @{remove=$AddressRemove}

# Report what we've done
 $ReportLine  = [PSCustomObject] @{
   Group   = $Group.DisplayName
   DN      = $Group.DistinguishedName
   NewName = $NewDisplayName
   Info    = $ArchiveInfo
   Owner   = $AdminAccount
   OldSmtp = $OldSmtpAddress
   NewSmtp = $NewSmtpAddress }
  $Report.Add($ReportLine) 
}

$Report | Out-GridView
