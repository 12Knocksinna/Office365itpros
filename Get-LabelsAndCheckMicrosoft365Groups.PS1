# Get-LabelsAndCheckMicrosoft365Groups.PS1
# Find any groups without a sensitivity (container) label and apply a default label. And then check all the
# Microsoft 365 groups in the tenant to make sure that they have the correct container management label assigned.
# An example of using Graph schema extensions.

# V1.0 22-Oct-2024
# GitHub link:

# Requires the Groups.ReadWrite.All permission to read and write group settings.

Connect-MgGraph -Scopes Group.ReadWrite.All

$DefaultSensitivityLabel = "e42fd42e-7240-4df0-9d8f-d14658bcf7ce" # Guid for General Access
$DefaultSensitivityLabelName = "General Access"
[int]$LabelsAssigned = 0

Write-Host "Fetching Microsoft 365 Groups..."
[array]$Groups = Get-MgGroup -All -PageSize 500 -Filter "(groupTypes/any(c:c eq 'unified'))" `
    -Property DisplayName, Id, assignedLabels, office365itpros_labels

Write-Host ("{0} Microsoft 365 Groups found" -f $Groups.Count)
[array]$GroupsWithLabels = $Groups | Where-Object { $_.assignedLabels -ne $null} | Sort-Object DisplayName
[array]$GroupsNoLabels = $Groups | Where-Object { $_.assignedLabels.displayName -eq $null} | Sort-Object DisplayName
If ($GroupsNoLabels) {
    Write-Host ("{0} Microsoft 365 Groups found without a container management label" -f $GroupsNoLabels.Count)
    $GroupsNoLabels | Format-Table DisplayName, Id
} Else {
    Write-Host "All Microsoft 365 Groups have a container management label"
}

# Add the settings to store the default container management label to the properties defined for the schema extensions
$DateLabelUpdated = Get-Date -format 'yyyy-MM-ddTHH:mm:ssZ'
$LabelSchemaSettings  = @{}
$LabelSchemaSettings.Add("LabelId", $DefaultSensitivityLabel)
$LabelSchemaSettings.Add("LabelDisplayName", $DefaultSensitivityLabelName)
$LabelSchemaSettings.Add("DateLabelUpdated", $DateLabelUpdated)

# Define the label to assign for the parameters passed to Update-MgGroup
$AssignedLabels = @{}
$AssignedLabels.Add("LabelId", $DefaultSensitivityLabel)

# Set up the body parameter for Update-MgGroup containing the default container management label and the 
# schema extension peroperties
$Parameters = @{}
$Parameters.Add("office365itpros_labels", $LabelSchemaSettings)
$Parameters.Add("assignedLabels", @($AssignedLabels))

# Update the groups that didn't have an assigned sensitivity label
ForEach ($Group in $GroupsNoLabels) {
    Write-Host ("Assigning default container management label {0} to group {1}" -f $DefaultSensitivityLabelName, $Group.DisplayName)
    Try {
        Update-MgGroup -GroupId $Group.Id -BodyParameter $Parameters
        $LabelsAssigned++
    } Catch {
        Write-Host ("Failed to assign default container management label to group {0}" -f $Group.DisplayName)
    }
}
Write-Host ("{0} Microsoft 365 Groups updated with the default container management label" -f $LabelsAssigned)

# Now that the unlabeled groups have a container management label, let's check the other groups to make sure that they 
# have the correct label assigned
Write-Host ("Checking {0} Microsoft 365 Groups to verify that they have the correct container management label" -f $GroupsWithLabels.Count)
[int]$i = 0
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Group in $GroupsWithLabels) {
    $i++
    Write-Host ("Updating label schema for group {0} ({1}/{2})" -f $Group.DisplayName, $i, $GroupsWithLabels.count)
    # Retrieve current container management label
    $CurrentLabelName = $Group.assignedLabels.DisplayName
    $CurrentLabelId = $Group.assignedLabels.LabelId
    # Retrieve the current schema extension properties
    $StoredLabelId = $Group.additionalProperties.office365itpros_labels['LabelId']
    $StoredLabelName = $Group.additionalProperties.office365itpros_labels['LabelDisplayName']
    If ($CurrentLabelId -ne $StoredLabelId) {
        Write-Host ("Mismatch in label for group {0}. Current label is {1}, stored label is {2}" -f $Group.DisplayName, $CurrentLabelName, $StoredLabelName) -ForegroundColor Red
        Write-Host "Updating label to revert to stored label" -ForegroundColor Yellow
        $LabelSchemaSettings  = @{}
        $DateLabelUpdated = Get-Date -format 'yyyy-MM-ddTHH:mm:ssZ'
        $LabelSchemaSettings.Add("DateLabelUpdated", $DateLabelUpdated)
        $AssignedLabels = @{}
        $AssignedLabels.Add("LabelId", $StoredLabelId)
        $Parameters = @{}
        $Parameters.Add("office365itpros_labels", $LabelSchemaSettings)
        $Parameters.Add("assignedLabels", @($AssignedLabels))
        Try {
            Update-MgGroup -GroupId $Group.Id -BodyParameter $Parameters
            $ReportLine = [PSCustomObject][Ordered]@{ 
                'Group Name'    = $Group.DisplayName
                'Current Label' = $CurrentLabelName
                'Stored Label'  = $StoredLabelName
                Action          = "Reverted"
                Date            = $DateLabelUpdated
            }
            $Report.Add($ReportLine)
        } Catch {
            Write-Host ("Failed to update group {0} with label {1}" -f $Group.DisplayName, $CurrentLabelName)
            $ReportLine = [PSCustomObject][Ordered]@{ 
                'Group Name'    = $Group.DisplayName
                'Current Label' = $CurrentLabelName
                'Stored Label'  = $StoredLabelName
                Action          = "Failed to Revert"
                Date            = $DateLabelUpdated
            }
            $Report.Add($ReportLine)
        }
    }
}

Write-Host ""
Write-Host ("Monitoring completed. Inconsistent container management label detected for {0} groups" -f $Report.Count)
$Report | Format-Table -AutoSize

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.