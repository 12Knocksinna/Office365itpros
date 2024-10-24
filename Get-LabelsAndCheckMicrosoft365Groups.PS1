# Get-LabelsAndCheckMicrosoft365Groups.PS1
# Find any groups without a sensitivity (container) label and apply a default label. And then check all the
# Microsoft 365 groups in the tenant to make sure that they have the correct container management label assigned.
# An example of using Graph schema extensions.

# V1.0 22-Oct-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-LabelsAndCheckMicrosoft365Groups.PS1

# Requires the Groups.ReadWrite.All permission to read and write group settings.

Connect-MgGraph -Scopes Group.ReadWrite.All

$DefaultSensitivityLabel = "e42fd42e-7240-4df0-9d8f-d14658bcf7ce" # Guid for General Access
$DefaultSensitivityLabelName = "General Access"
[int]$LabelsAssigned = 0
[int]$LabelDateUpdated = 0

Write-Host "Scanning for Microsoft 365 Groups..."
# Look for groups that don't have a container management label or don't have the correct schema extension properties

[array]$Groups = Get-MgGroup -PageSize 500 -All -Property Id, displayName, assignedLabels, office365itpros_labels `
    -Filter "(groupTypes/any(c:c eq 'unified')) and (office365itpros_labels/datelabelupdated eq null or office365itpros_labels/labelid eq null)" -ConsistencyLevel Eventual -CountVariable Count
Write-Host ("{0} Microsoft 365 Groups found that need their container management settings adjusted" -f $Groups.Count)

# Check each group. Add the default container management label if it's missing. Update the schema extension date label added property if it's missing
ForEach ($Group in $Groups) {
    If (!($Group.assignedLabels)) {
        Write-Host ("Assigning default container management label {0} to group {1}" -f $DefaultSensitivityLabelName, $Group.DisplayName)
        # Populate hash table with schema extension label properties
        $DateLabelUpdated = Get-Date -format 'yyyy-MM-ddTHH:mm:ssZ'
        $LabelSchemaSettings  = @{}
        $LabelSchemaSettings.Add("LabelId", $DefaultSensitivityLabel)
        $LabelSchemaSettings.Add("LabelDisplayName", $DefaultSensitivityLabelName)
        $LabelSchemaSettings.Add("DateLabelUpdated", $DateLabelUpdated)
        # Populate hash table with label id for the default container management label
        $AssignedLabels = @{}
        $AssignedLabels.Add("LabelId", $DefaultSensitivityLabel)
        # Populate hash table to use with Update-MgGroup
        $Parameters = @{}
        $Parameters.Add("office365itpros_labels", $LabelSchemaSettings)
        $Parameters.Add("assignedLabels", @($AssignedLabels))
        # And update the group
        Try {
            Update-MgGroup -GroupId $Group.Id -BodyParameter $Parameters
            $LabelsAssigned++
        } Catch {
            Write-Host ("Failed to assign default container management label to group {0}" -f $Group.DisplayName)
        }   
    } Elseif (($Group.assignedLabels) -and (!($Group.additionalProperties.office365itpros_labels))) {
        # This code handles the situation where a new group is created with a container management label but the data has never
        # been written into the schema extension properties. 
        Write-Host ("Updating label assignment data for group {0}" -f $Group.DisplayName)
        $CurrentLabelName = $Group.assignedLabels.DisplayName
        $CurrentLabelId = $Group.assignedLabels.LabelId
        $DateLabelUpdated = Get-Date -format 'yyyy-MM-ddTHH:mm:ssZ'
        $LabelSchemaSettings = @{}
        $LabelSchemaSettings.Add("LabelId", $CurrentLabelId)
        $LabelSchemaSettings.Add("LabelDisplayName", $CurrentLabelName)
        $LabelSchemaSettings.Add("DateLabelUpdated", $DateLabelUpdated)
        $Parameters = @{}
        $Parameters.Add("office365itpros_labels", $LabelSchemaSettings)
        Try {
            Update-MgGroup -GroupId $Group.Id -BodyParameter $Parameters
            $LabelDateUpdated++
        } Catch {
            Write-Host ("Failed to assign default container management label to group {0}" -f $Group.DisplayName)
        }   
    }
}

If ($LabelsAssigned -gt 0) {
    Write-Host ("{0} Microsoft 365 Groups updated with the default container management label" -f $LabelsAssigned)
}
If ($LabelDateUpdated -gt 0) {      
    Write-Host ("{0} Microsoft 365 Groups updated with an assignment date for the container management label" -f $LabelDateUpdated)
}

# Now that the unlabeled groups have a container management label, let's check the other groups to make sure that they 
# have the correct label assigned
[array]$Groups = Get-MgGroup -All -PageSize 500 -Filter "(groupTypes/any(c:c eq 'unified'))" `
    -Property DisplayName, Id, assignedLabels, office365itpros_labels | Sort-Object DisplayName
Write-Host ("Checking {0} Microsoft 365 Groups to verify that they have the correct container management label" -f $GroupsWithLabels.Count)
[int]$i = 0
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Group in $Groups) {
    $i++
    $CurrentLabelName = $null; $CurrentLabelId = $null; $StoredLabelId = $null; $StoredLabelName = $null
    Write-Host ("Checking label schema for group {0} ({1}/{2})" -f $Group.DisplayName, $i, $GroupsWithLabels.count)
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
            $Parameters
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

If ($Report) {
    Write-Host ""
    Write-Host ("Monitoring completed. Inconsistent container management label detected for {0} groups" -f $Report.Count)
    $Report | Format-Table -AutoSize
} Else {
    Write-Host ""
    Write-Host "All Microsoft 365 Groups have the correct container management label" -ForegroundColor Green
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
