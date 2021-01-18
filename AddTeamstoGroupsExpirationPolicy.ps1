# Add all Teams that aren't already covered by the Groups expiration policy 
# to the policy
# https://github.com/12Knocksinna/Office365itpros/blob/master/AddTeamstoGroupsExpirationPolicy.ps1
$PolicyId = (Get-AzureADMSGroupLifecyclePolicy).Id
$TeamsCount = 0
Write-Host "Fetching list of Teams in the tenant..."
$Teams = Get-Team
ForEach ($Team in $Teams) {
  $CheckPolicy = (Get-UnifiedGroup -Identity $Team.GroupId).CustomAttribute3
  If ($CheckPolicy -eq $PolicyId) {
    Write-Host "Team" $Team.DisplayName "is already covered by the expiration policy" }
  Else { 
    Write-Host "Adding team" $Team.DisplayName "to group expiration policy"
    Add-AzureADMSLifecyclePolicyGroup -GroupId $Team.GroupId -Id $PolicyId -ErrorAction SilentlyContinue
    Set-UnifiedGroup -Identity $Team.GroupId -CustomAttribute3 $PolicyId
    $TeamsCount++ }}
Write-Host "All done." $TeamsCount "teams added to policy"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization.  Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
