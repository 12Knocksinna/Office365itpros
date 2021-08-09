# EnableMicrosoft365GroupCreationControl.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/EnableMicrosoft365GroupCreationControl.PS1

# Check that we have the right module loaded
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "AzureADPreview")) {Write-Host "Please run Connect-AzureAD using the Azure AD Preview module and then restart the script"; break}

# Define values for the group used to control group creation - change this for your tenant
[string]$GroupAllowedToCreate = "GroupCreationControl"
# Define whether you want to turn group creation on or off ($True or $False)
$OnOffSwitch = $True
# Just for formatting...
Switch ($OnOffSwitch) {
   $True { $Control = "On" }
   $False { $Control = "Off" }
}

# Check whether we can find the group and if we find more than one, ask the administrator to select a group
[array]$Group = (Get-AzureADGroup -SearchString $GroupAllowedToCreate)
If (!$Group) { Write-Host ("Can't find the group identifier for {0} - is it the correct group name?" -f $GroupAllowedToCreate) ; break }
If ($Group.Count -eq 1) { # Just one group found
   [string]$GroupId = $Group.ObjectId }
Elseif ($Group.Count -gt 1) { # More than one groupfound. Ask which to use
   CLS; Write-Host "More than one matching group was found."; [int]$i=1
   ForEach ($GroupOption in $Group) {
      Write-Host $i ":" $GroupOption.DisplayName; $i++ }
   [Int]$Answer = Read-Host "Enter the number of the group to use"
   If (($Answer -gt 0) -and ($Answer -le $i)) {
      [int]$i = ($Answer-1)
      [string]$GroupId = $Group[$i].ObjectId 
      [string]$GroupAllowedToCreate = $Group[$i].DisplayName
      Write-Host "OK. Selected group is" $GroupAllowedToCreate
    } #end if
  }


Write-Host ("Setting group creation control to {0} using group name {1}." -f $Control, $GroupAllowedToCreate)

$PolicySettingsId = (Get-AzureADDirectorySetting | ? {$_.DisplayName -eq "Group.Unified"}).Id
If (!$PolicySettingsId) { # No policy settings found for the tenant, so create it and extract the identifier
  $PolicyTemplate = Get-AzureADDirectorySettingTemplate | ? {$_.DisplayName -eq "Group.Unified"}
  $PolicySettings = $PolicyTemplate.CreateDirectorySetting()
  New-AzureADDirectorySetting -DirectorySetting $PolicySettings
  $PolicySettingsId = (Get-AzureADDirectorySetting | ? {$_.DisplayName -eq "Group.Unified"}).Id
} # End If

$PolicySettings = Get-AzureADDirectorySetting -Id $PolicySettingsId
$PolicySettings["EnableGroupCreation"] = $OnOffSwitch
$PolicySettings["GroupCreationAllowedGroupId"] = $GroupId
Set-AzureADDirectorySetting -Id $PolicySettingsId -DirectorySetting $PolicySettings

# Check what we have done and report the current status
$CurrentValues = Get-AzureADDirectorySetting | ?{$_.DisplayName -eq "Group.Unified"}
$GroupId = $CurrentValues.Values |?{$_.Name -eq "GroupCreationAllowedGroupId" } | Select -ExpandProperty Value
$OnOffSwitch = $CurrentValues.Values |?{$_.Name -eq "EnableGroupCreation" } | Select -ExpandProperty Value

Switch ($OnOffSwitch) {
   $True { $Control = "Restricted" }
   $False { $Control = "Unrestricted" }
}
CLS
Write-Host ""
Write-Host ("Group creation is currently {0}. The name of the group defined by policy to control group creation is {1} and its object identifier is {2}" -f $Control, (Get-AzureADGroup -ObjectId $GroupId).DisplayName, $GroupId)
# List the members of the group
Get-AzureADGroupMember -ObjectId $GroupId

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
