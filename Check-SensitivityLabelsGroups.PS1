# Originally featured in https://practical365.com/monitor-changes-sensitivity-labels-container-management/

$Status = Get-ConnectionInformation -ErrorAction SilentlyContinue
If (!($Status)) {
  Connect-ExchangeOnline -SkipLoadingCmdletHelp 
}
Connect-IPPSSession

# Define the default sensitivity label
$DefaultSensitivityLabel = "e42fd42e-7240-4df0-9d8f-d14658bcf7ce" # Guid for General Access

# Create a list of the sensitivity labels used for container management in the organization
[array]$Labels = Get-Label
$ContainerLabels = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Label in $Labels) { 
    If ($Label.ContentType -Like "*UnifiedGroup*") { # It's a label for container management
      $DataLine = [PSCustomObject] @{
        LabelId     = $Label.ImmutableId
        DisplayName = $Label.DisplayName
        Priority    = $Label.Priority } 
      $ContainerLabels.Add($DataLine) 
    }
}
# Validate that the default sensitivity label is OK

If ($DefaultSensitivityLabel -notin $ContainerLabels.LabelId) {
    Write-Host ("Default label {0} is not valid -exiting" -f $DefaultSensitivityLabel)
}

[int]$LabelsAssigned = 0
Write-Host "Fetching Microsoft 365 Groups"
[array]$Groups = Get-UnifiedGroup -ResultSize Unlimited
ForEach ($Group in $Groups) {
   If ($null -eq $Group.SensitivityLabel) { # No label assigned so let's assign the default label
      Write-Host ("{0} has no sensitivity label - assigning the default label" -f $Group.DisplayName) -Foregroundcolor Red
      Set-UnifiedGroup -Identity $Group.ExternalDirectoryObjectId -SensitivityLabel $DefaultSensitivityLabel -CustomAttribute14 $DefaultSensitivityLabel
      $LabelsAssigned++
    } Else { # Just update the Custom Attribute
      Set-UnifiedGroup -Identity $Group.ExternalDirectoryObjectId  -CustomAttribute14 $Group.SensitivityLabel
    }
} # End For
Write-Host ("Labels assigned to {0} Microsoft 365 Groups; checked label tracking attribute for {1} groups" -f $LabelsAssigned, $Groups.Count)


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
