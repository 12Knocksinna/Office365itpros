# FindOffice365GroupsWithBadGuests.PS1
# A script to scan Office 365 Groups to find groups with guest members that also have a sensitivity label that blocks guests
# V1.0 March 2020

# First Check for connection to compliance endpoint
$TenantLabels = @{}
Write-Host "Finding the sensitvity labels defined in the tenant"
Try {
    $Labels = Get-Label }
Catch {
    Write-Host "Your PowerShell session must be connected to the Compliance endpoint to fetch label data" ; break}
# Now Populate hash table with label data  
$Labels.ForEach( {
       $TenantLabels.Add([String]$_.ImmutableId, $_.DisplayName) } )
# Now figure out which labels block the adding of guests to a group's membership
$LabelsBlockingGuests = @{}
ForEach ($Label in $Labels) { 
       $LabelGuestAccess = $True
       $LabelActions = $Label.LabelActions | ConvertFrom-Json      
       ForEach ($LabelAction in $LabelActions) {
          If ($LabelAction.Type -eq "protectgroup") {
             $Settings = $LabelAction.Settings
             ForEach ($Setting in $Settings) {
                If ($Setting.Key -eq "allowaccesstoguestusers" -and $Setting.Value -eq "false") {
                  $LabelsBlockingGuests.Add([String]$Label.ImmutableId, $Label.DisplayName)}}}
}}
       
CLS; Write-Host "Finding Office 365 Groups with sensitivity labels that block guests..."
# Find groups that have a sensitivity label and have 
$Groups = Get-UnifiedGroup -ResultSize UnLimited | ? {$_.SensitivityLabel -ne $Null -and $_.GroupExternalMemberCount -gt 0} 
CLS
If (!$Groups.Count) { Write-Host "No Office 365 Groups found with guest users"}
  Else {
     $Report = [System.Collections.Generic.List[Object]]::new(); $NumberGuests = 0
     $ProgressDelta = 100/($Groups.count); $PercentComplete = 0; $GrpNumber = 0
     ForEach ($Group in $Groups) {
       $LabelDisplayName = $Null; $GrpNumber++
       $GrpStatus = $Group.DisplayName + " ["+ $GrpNumber +"/" + $Groups.Count + "]"
       Write-Progress -Activity "Processing group" -Status $GrpStatus -PercentComplete $PercentComplete
       $PercentComplete += $ProgressDelta
       $LabelDisplayName = $LabelsBlockingGuests.Item($Group.SensitivityLabel.Guid)
       If ($Null -ne $LabelDisplayName) { # We have a group with a label that blocks guests
          $Users = Get-UnifiedGroupLinks -Identity $Group.Alias -LinkType Members
          ForEach ($U in $Users) {
             If ($U.Name -Match "#EXT#" -and $U.Name -NotLike "*teams.ms*") {
## Remember to edit the string to make sure it’s your tenant name…
##          $CheckName = $U.Name + "@EditMeTenantName.onmicrosoft.com"
            $CheckName = $U.Name + "@redmondassociates.onmicrosoft.com"
            $User = (Get-AzureADUser -ObjectId $CheckName).DisplayName 
            $ReportLine = [PSCustomObject]@{
               Email           = $U.Name
               User            = $User
               Group           = $Group.DisplayName
               Site            = $Group.SharePointSiteURL
               Label           = $LabelDisplayName
               LabelGuid       = $Group.SensitivityLabel.Guid }
            $Report.Add($ReportLine)
            $NumberGuests++ }}         
}}}

$Report | Sort Email | Out-GridView
$Report | Export-CSV -NoTypeInformation c:\temp\GroupsWithGuestsBlocked.csv
Write-Host "All done." $NumberGuests "guests found in" $Groups.Count "groups. Output available in c:\temp\GroupsWithGuestsBlocked.csv"
