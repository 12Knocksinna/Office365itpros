# Remove-OldPlannerTasks.PS1
# This script removes old completed tasks from Microsoft Planner based on a specified age threshold.

# GitHub link:
# V1.0 25-Jul-2025

Connect-MgGraph -Scopes "Group.ReadWrite.All", "Tasks.ReadWrite"

$ThresholdDate = (Get-Date).AddDays(-365) # Set the threshold date to 365 days ago
$Report = [System.Collections.Generic.List[object]]::new()

# Find all Microsoft 365 Groups
[array]$Groups = $groups = 
Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified')" -Property Id, DisplayName
If ($Groups) {
    Write-Host "Found $($Groups.Count) Microsoft 365 Groups - checking each group for Plans."
}

ForEach ($Group in $Groups) {
    # Get all plans for the group
    [array]$Plans = Get-MgGroupPlannerPlan -GroupId $Group.Id -All
    If ($Plans) {
        Write-Host "Found $($Plans.Count) Plans in Group '$($Group.DisplayName)' - checking each plan for tasks."
    }

    ForEach ($Plan in $Plans) {
        # Get all tasks for the plan that are completed and the completion date is less than the threshold date
        [array]$Tasks = Get-MgPlannerPlanTask -PlannerPlanId $Plan.Id -All | Where-Object { $_.PercentComplete -eq '100' -as [int] -and $_.CompletedDateTime -lt $ThresholdDate}
        If ($Tasks) {
            Write-Host "Found $($Tasks.Count) old completed Tasks in Plan '$($Plan.Title)' - evaluating for removal."
        }

        ForEach ($Task in $Tasks) {
            # Check if the task is completed and older than the threshold date
            
                # Remove the old completed task
                Try {
                    Remove-MgPlannerTask -PlannerTaskId $Task.Id -ErrorAction Stop
                    Write-Host "Removed Task '$($Task.Title)' from Plan '$($Plan.Title)' in Group '$($Group.DisplayName)'."
                    # Add to report
                    $Report.Add([PSCustomObject]@{
                        GroupName       = $Group.DisplayName
                        PlanTitle       = $Plan.Title
                        TaskTitle       = $Task.Title
                        CompletedDate   = $Task.CompletedDateTime
                        Timestamp       = Get-Date -Format "s"
                    })
                } Catch {
                    Write-Host "Failed to remove Task '$($Task.Title)' from Plan '$($Plan.Title)' in Group '$($Group.DisplayName)': $($_.Exception.Message)"
                }
                
                
            
        }
    }
}

Write-Host "Task removal process completed. Total tasks removed: $($Report.Count). using a threshold date of $($ThresholdDate.ToString('dd-MMM-yyyy'))."
Write-Host "Number of plans scanned: $($Plans.count)"
Write-Host "Number of groups scanned: $($Groups.count)"
Write-Host "Tasks removed from unique plans: $($Report | Select-Object -Property PlanTitle -Unique | Measure-Object).Count"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.