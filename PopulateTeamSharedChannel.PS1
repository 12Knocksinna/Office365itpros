# PopulateTeamSharedChannel.PS1
# An example of how to populate the membership of a shared channel with every user account (with licenses) in a tenant,
# the idea being that the shared channel is available to everyone.

$Modules = @( "MicrosoftTeams", "Microsoft.Graph" )
# Requires -Modules $Modules

$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "MicrosoftTeams")) {Write-Host "Please connect to the Microsoft Teams module and then restart the script"; break}

Connect-MgGraph -Scopes User.Read.All
Select-MgProfile Beta

# Find user accounts with licenses
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All
If (!($Users)) { Write-Host "No user accounts found - exiting" ; break }
# Filter out any accounts marked that shouldn't be added to team membership
$FilteredUsers = $Users | ? {$_.OfficeLocation -ne "XXX"}

# Now check that each user actually has a Teams service plan
$UsersWithTeams =  [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $FilteredUsers) {
  $TeamsLicense =  Get-MgUserLicenseDetail -UserId $User.Id | Select-Object -ExpandProperty ServicePlans | ? {$_.ServicePlanId -eq "57ff2da0-773e-42df-b2af-ffb7a2317929"} | Select-Object -ExpandProperty ProvisioningStatus
 If ($TeamsLicense -eq "Success") {
     $UserData = [PSCustomObject][Ordered]@{  # Write out details of the user
       Id          = $User.Id
       DisplayName = $User.DisplayName }
     $UsersWithTeams.Add($UserData)
 } #End if
} #End ForEach

# These variables will differ depending on the host team and channel name you decide to use
$GroupId = (Get-Team -DisplayName "HR Questions and Answers").GroupId
$ChannelName = "Questions and Answers"

Write-Host ("Finding the membership of of the {0} channel" -f $ChannelName)
# Find current members and owners and add them to a hash table that we can lookup
$ChannelMembers = Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName -Role Member
$ChannelOwners = Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName -Role Owner

$CurrentMembers = @{}
ForEach ($Member in $ChannelMembers) {
   $CurrentMembers.Add($Member.UserId,$Member.User) }
ForEach ($Member in $ChannelOwners) {
   $CurrentMembers.Add($Member.UserId,$Member.User) }


$i = 0
# Check each user and add them if they are not found
ForEach ($User in $UsersWithTeams) {
   If (!($CurrentMembers[$User.Id])) {
      Write-Host ("Adding {0} to the {1} channel" -f $User.DisplayName, $ChannelName)
      Add-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName -User $User.Id; $i++ }
}

Write-Host ("{0} new members added to channel" -f $i)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
