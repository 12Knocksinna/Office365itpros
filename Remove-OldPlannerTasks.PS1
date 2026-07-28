# Remove-OldPlannerTasks.PS1
# This script removes old completed tasks from Microsoft Planner based on a specified age threshold.

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Remove-OldPlannerTasks.PS1
# V1.0 25-Jul-2025

Connect-MgGraph -Scopes "Group.ReadWrite.All", "Tasks.ReadWrite"

$ThresholdDate = (Get-Date).AddYears(-5) # Set the threshold date to five years ago. Set this to be whatever age threshold you want

# Find all Microsoft 365 Groups
Write-Host "Searching for Microsoft 365 Groups..."
[array]$Groups = Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified')" -Property Id, DisplayName -PageSize 500
If ($Groups) {
    $Groups = $Groups | Sort-Object DisplayName
    Write-Host "Found $($Groups.Count) Microsoft 365 Groups - checking each group for Plans."
    [int]$PlansScanned = 0
} Else {
    Write-Host "No Microsoft 365 Groups found."
    return
}

$Report = [System.Collections.Generic.List[object]]::new()
[datetime]$StartTime = Get-Date
ForEach ($Group in $Groups) {
    # Get all plans for the group
    [array]$Plans = Get-MgGroupPlannerPlan -GroupId $Group.Id -All
    If ($Plans) {
        Write-Host "Found $($Plans.Count) Plan(s) in Group '$($Group.DisplayName)' - checking each plan for tasks." -ForegroundColor Yellow
        $PlansScanned++
    }

    ForEach ($Plan in $Plans) {
        [int]$TasksRemovedFromPlan = 0
        # Get all tasks for the plan that are completed and the completion date is less than the threshold date
        [array]$Tasks = Get-MgPlannerPlanTask -PlannerPlanId $Plan.Id -All -PageSize 500| `
            Where-Object { $_.PercentComplete -eq '100' -as [int] -and $_.CompletedDateTime -lt $ThresholdDate}
        If ($Tasks) {
            Write-Host "Found $($Tasks.Count) old completed Tasks in Plan '$($Plan.Title)' - evaluating for removal." -ForegroundColor Cyan
        }

        ForEach ($Task in $Tasks) {           
            # Remove the old completed task
            Try {
                # Remove-MgPlannerTask -PlannerTaskId $Task.Id -ErrorAction Stop
                #Write-Host "Removed Task '$($Task.Title)' from Plan '$($Plan.Title)' in Group '$($Group.DisplayName)'."
                # Add to report
                    $Report.Add([PSCustomObject]@{
                        GroupName       = $Group.DisplayName
                        PlanTitle       = $Plan.Title
                        TaskTitle       = $Task.Title
                        CreationDate    = $Task.CreatedDateTime
                        CompletedDate   = $Task.CompletedDateTime
                        Timestamp       = Get-Date -Format "s"
                        PlanId          = $Plan.Id
                    })
                    $TasksRemovedFromPlan++
                } Catch {
                    Write-Host "Failed to remove Task '$($Task.Title)' from Plan '$($Plan.Title)' in Group '$($Group.DisplayName)': $($_.Exception.Message)"
                }             
            }
            If ($TasksRemovedFromPlan -gt 0) {
                Write-Host "Removed $TasksRemovedFromPlan old completed Tasks from Plan '$($Plan.Title)' in Group '$($Group.DisplayName)'."
                # Small pause to avoid hitting API rate limits
                Start-Sleep -Seconds 1
            }
        }
    }
[datetime]$EndTime = Get-Date
$Duration = $EndTime - $StartTime
$UniquePlansCount = ($Report | Select-Object -Property PlanId -Unique | Measure-Object).Count
Write-Host ""
Write-Host "Task removal process completed. $($Report.Count) tasks removed using a threshold date of $($ThresholdDate.ToString('dd-MMM-yyyy'))."
Write-Host "Number of plans scanned: $PlansScanned" 
Write-Host "Number of groups scanned: $($Groups.count)"
Write-Host "Tasks removed from unique plans: $UniquePlansCount"
Write-Host "Duration: $($Duration.ToString())"

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.