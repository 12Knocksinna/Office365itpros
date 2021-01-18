# https://github.com/12Knocksinna/Office365itpros/blob/master/ArchiveMicrosoft365Groups.PS1
# A script to process a set of Microsoft 365 Groups and Archive them so that the information is kept and available
# online for eDiscovery but can't be accessed nby users
# Uses the Exchange Online Module. Must be run by a tenant admin.
# Also needs the Microsoft Teams module to check for private channels - updated 17-Jul-2020

# Check that we're connected to Exchange Online and Teams
Write-Host "Checking that prerequisite PowerShell modules are loaded..."
Try { $OrgName = (Get-OrganizationConfig).Name }
   Catch  {
      Write-Host "Your PowerShell session is not connected to Exchange Online."
      Write-Host "Please connect to Exchange Online using an administrative account and retry."
      Break }

$TeamsCheck = Get-Module -Name MicrosoftTeams
If ($TeamsCheck -eq $Null) {
     Write-Host "Your PowerShell session is not connected to Microsoft Teams."
     Write-Host "Please connect to Microsoft Teams using an administrative account and retry."; Break }


Write-Host "Preparing to archive" $ArchiveGroups.Count "Microsoft 365 Groups"
# Find list of groups to be archived
$ArchiveGroups = Get-UnifiedGroup -Filter {CustomAttribute1 -eq "Archive"} -ResultSize 1000 | Select DisplayName, DistinguishedName, Alias, PrimarySmtpAddress, SensitivityLabel, ExternalDirectoryObjectId

# If you don't want to use custom attributes to mark groups to be archived, you can export the groups to a CSV file
# review them with Excel, remove groups that you don't want to archive, and use the remaining set as the input to the
# script. To do this, create the CSV file with:
# $Groups = Get-UnifiedGroup -ResultSize Unlimited | Select DisplayName, Notes, DistinguishedName, Alias, PrimarySmtpAddress, SensitivityLabel, ExternalDirectoryObjectId
# $Groups | Sort DisplayName | Export-CSV -NoTypeInformation c:\temp\GroupsForReview.CSV
# and after the review is done, replace the call to Get-UnifiedGroup above with:
# $ArchiveGroups = Import-CSV c:\temp\GroupsForReview.CSV

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
# Need to check if the group is team-enabled
  Try {
      $Team = Get-Team -GroupId $Group.ExternalDirectoryObjectId }
  Catch
      { $Status = 0 }

# Get lists of current owners and members and add the compliance admin as an owner, then remove the existing owners and members
If ($Team) { # This group is team-enabled, so we process membership details with Teams cmdlets
   $CurrentOwners = Get-TeamUser -GroupId $Group.ExternalDirectoryObjectId -Role Owner
   $CurrentMembers = Get-TeamUser -GroupId $Group.ExternalDirectoryObjectId -Role Member
   Add-TeamUser -GroupId $Group.ExternalDirectoryObjectId -User $AdminAccount -Role Owner
   Start-Sleep -Seconds 2 # Let membership settle down
   $TeamChannels = Get-TeamChannel -GroupId $Group.ExternalDirectoryObjectId -Membershiptype Private 
   If ($TeamChannels) { # Add compliance admin as a member and owener for each private channel
       ForEach ($Channel in $TeamChannels) {
         Add-TeamChannelUser -GroupId $Group.ExternalDirectoryObjectId -User $AdminAccount -DisplayName $Channel.DisplayName
         Add-TeamChannelUser -GroupId $Group.ExternalDirectoryObjectId -User $AdminAccount -DisplayName $Channel.DisplayName -Role Owner }
    } #End TeamChannel If
   ForEach ($Owner in $CurrentOwners) { 
      Remove-TeamUser -GroupId $Group.ExternalDirectoryObjectId -User $Owner.User }
   ForEach ($Member in $CurrentMembers) { 
      Remove-TeamUser -GroupId $Group.ExternalDirectoryObjectId -User $Member.User }
} # End If to process group using Teams cmdlets
Else { # it's a normal group, so use the EXO cmdlets to update group membership
  $CurrentOwners = (Get-UnifiedGroupLinks -Identity $Group.DistinguishedName -LinkType Owners | Select Name)
  $CurrentMembers = (Get-UnifiedGroupLinks -Identity $Group.DistinguishedName -LinkType Members | Select Name)
  Add-UnifiedGroupLinks -Identity $Group.DistinguishedName -LinkType Members -Links $AdminAccount
  Add-UnifiedGroupLinks -Identity $Group.DistinguishedName -LinkType Owners -Links $AdminAccount
  Start-Sleep -Seconds 1
  ForEach ($Owner in $CurrentOwners) { 
        Remove-UnifiedGroupLinks -Identity $Group.DistinguishedName -LinkType Owners -Links $Owner.Name -Confirm:$False }
  ForEach ($Member in $CurrentMembers) { 
        Remove-UnifiedGroupLinks -Identity $Group.DistinguishedName -LinkType Members -Links $Member.Name -Confirm:$False }
} # End Else

# Create SMTP Address for the archived group
  $OldSmtpAddress = $Group.PrimarySmtpAddress # Just for reporting
  $NewSmtpAddress = $Group.PrimarySmtpAddress.Split("@")[0] +  "_archived" + "@" + $Group.PrimarySmtpAddress.Split("@")[1]
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
$Report | Export-CSV -NoTypeInformation c:\temp\ArchivedGroups.csv

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
