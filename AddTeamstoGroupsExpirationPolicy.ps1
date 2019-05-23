# Add all Teams that aren't already covered by the Groups expiration policy 
# to the policy
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
