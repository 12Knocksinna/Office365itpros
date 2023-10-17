# EnableMicrosoft365GroupCreationControl.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/EnableMicrosoft365GroupCreationControl.PS1

# A script to update the setting in the Azure AD policy for Microsoft 365 Groups used to control if group creation is restricted.

# Usage .\EnableMicrosoft365GroupCreationControl.PS1 -GroupAllowedToCreate GroupUsedtoControlGroups -OnOffSwitch $True/$False
# e.g.  .\EnableMicrosoft365GroupCreationControl.PS1 -GroupAllowedToCreate GroupUsedtoControlGroups -OnOffSwitch $True

# GroupAllowedToCreate is the display name of a Microsoft 365 Group or security group used to hold the set of users allowed to create new Microsoft 365 Groups
# OnOffSwitch is $True to enable control of group creation or $False to disable it

Param ([Parameter(Mandatory)]$GroupAllowedToCreate, [string] [ValidateSet( "True", "False")]$OnOffSwitch=$True)

If ($GroupAllowedToCreate.length -eq $0) { Write-Host "No group name specified - please rerun"; break }
If ($OnOffSwitch -ne $True -and $OnOffSwitch -ne $False) { Write-Host "No mode specified - please rerun" ; break }

# Check that we have the right module loaded
Connect-MgGraph -Scopes Directory.ReadWrite.All, Group.Read.All

# Just for formatting...
If ($OnOffSwitch -eq $True) {
   $Control = "On"
} Else {
   $Control = "Off"
}

# Check whether we can find the group and if we find more than one, ask the administrator to select a group
$Filter = "displayname eq '" + $GroupAllowedToCreate + "'"
[array]$Group = (Get-MgGroup -filter $Filter)

If (!$Group) { 
   Write-Host ("Can't find the group identifier for {0} - is it the correct group name?" -f $GroupAllowedToCreate) ; break }
If ($Group.Count -eq 1) { # Just one group found
   [string]$GroupId = $Group.Id 
} Elseif ($Group.Count -gt 1) { # More than one groupfound. Ask which to use
   Clear-Host; Write-Host "More than one matching group was found."; [int]$i=1
   ForEach ($GroupOption in $Group) {
      Write-Host $i ":" $GroupOption.DisplayName; $i++ 
   }
   [Int]$Answer = Read-Host "Enter the number of the group to use"
   If (($Answer -gt 0) -and ($Answer -le $i)) {
      [int]$i = ($Answer-1)
      [string]$GroupId = $Group[$i].ObjectId 
      [string]$GroupAllowedToCreate = $Group[$i].DisplayName
      Write-Host "OK. Selected group is" $GroupAllowedToCreate
    } #end if
}

Write-Host ("Setting group creation control to {0} using group name {1}." -f $Control, $GroupAllowedToCreate)
$PolicySettingsId = Get-MgBetaDirectorySetting | Where-Object {$_.DisplayName -eq "Group.Unified"} | Select-Object -ExpandProperty Id

If (!$PolicySettingsId) { # No policy settings found for the tenant, so create it and extract the identifier
   $PolicyTemplate = Get-MgBetaDirectorySettingTemplate | Where-Object {$_.DisplayName -eq "Group.Unified"}
   New-MgBetaDirectorySetting -TemplateId $PolicyTemplate.Id
   $PolicySettingsId = (Get-MgBetaDirectorySetting | Where-Object {$_.DisplayName -eq "Group.Unified"}).Id
} # End If

$PolicySettings = Get-MgBetaDirectorySetting -DirectorySettingId $PolicySettingsId
$Values = $PolicySettings.Values
($Values | Where-Object Name -eq 'EnableGroupCreation').Value = $OnOffSwitch
($Values | Where-Object Name -eq 'GroupCreationAllowedGroupId').Value = $GroupId

Update-MgBetaDirectorySetting -DirectorySettingId $PolicySettingsId -Values $Values

# Check what we have done and report the current status
$CurrentValues = Get-MgBetaDirectorySetting | Where-Object {$_.DisplayName -eq "Group.Unified"}
$GroupId = $CurrentValues.Values | Where-Object {$_.Name -eq "GroupCreationAllowedGroupId" } | Select-Object -ExpandProperty Value
$OnOffSwitch = $CurrentValues.Values | Where-Object {$_.Name -eq "EnableGroupCreation" } | Select-Object -ExpandProperty Value

Switch ($OnOffSwitch) {
   $True { $Control = "Unrestricted" }
   $False { $Control = "Restricted" }
}
Clear-Host
Write-Host ""
[array]$Owners = (Get-MgGroupOwner -GroupId $GroupId)
[array]$OwnerNames = $Null
ForEach ($Owner in $Owners) { 
   $OwnerNames += (Get-MgUser -UserId $Owner.Id).DisplayName 
}
[string]$OwnerNames = $OwnerNames -join ", "
Write-Host ("The name of the group defined to control group creation is {0} and its identifier is {1}. Its owners are {2}." -f (Get-MgGroup -GroupId $GroupId).DisplayName, $GroupId, $OwnerNames)
Write-Host ""
Write-Host “The accounts allowed to create new Microsoft 365 groups are:”
(Get-MgGroupMember -GroupId $GroupId).additionalProperties.displayName


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
