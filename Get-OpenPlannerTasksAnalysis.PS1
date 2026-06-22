# Get-OpenPlannerTasksAnalysis.PS1
# This script retrieves the last six weeks of Planner task data from a SharePoint list and generates a summary of tasks per user per week, including an average tasks per week metric.

# V1.0 22-June-2026
# GitHub Link:

Connect-MgGraph -Scopes "Sites.Read.All", "Sites.ReadWrite.All"
# Define URL for the SharePoint site and construct the SiteId
$Uri = "https://redmondassociates.sharepoint.com/sites/O365ExchPro"
$SiteId = $Uri.Split('//')[1].split("/")[0] + ":/sites/" + $Uri.Split('//')[1].split("/")[2]
$Site = Get-MgSite -SiteId $SiteId
# Find the list named "Planner Task Burndown" in the SharePoint site
$List = Get-MgSiteList -SiteId $Site.Id -Filter "displayName eq 'Planner Task Burndown'"   
If ($List) {
    # Get list items and load them into a PowerShell list
    [array]$ListItems = Get-MgSiteListItem -ListId $List.Id -SiteId $Site.Id -ExpandProperty "fields(`$select=id,title,plan,upn,username,Tasks,RunDate,NewTasksAdded,Averagedaysoutstanding,Objectid)" -PageSize 999 -All
	
    $ItemData = [System.Collections.Generic.List[Object]]::new()
    ForEach ($Item in $ListItems.fields) { 
    
        $ReportLine = [PSCustomObject] @{ 
            Id                     = $Item.Id
            ItemId                 = $Item.additionalProperties.Title
            Plan                   = $Item.additionalProperties.Plan
            UPN                    = $Item.additionalProperties.UPN
            UserName               = $Item.additionalProperties.Username
            Tasks                  = $Item.additionalProperties.Tasks
            Rundate                = $Item.additionalProperties.RunDate
            NewTasksAdded          = $Item.additionalProperties.NewTasksAdded
            Averagedaysoutstanding = $Item.additionalProperties.Averagedaysoutstanding
            ObjectId               = $Item.additionalProperties.ObjectId
        }
        $ItemData.Add($ReportLine)
    }
}

# Define the date range for the last six weeks, starting from the most recent Saturday
$Today = (Get-Date).Date

$LastSaturday = $Today.AddDays(-(($Today.DayOfWeek - [System.DayOfWeek]::Saturday + 7) % 7))
[datetime]$SixWeeksAgo = $LastSaturday.AddDays(-42)

# Create six week buckets, oldest to newest
$Weeks = 5..0 | ForEach-Object {
    $WeekStart = $LastSaturday.AddDays(-7 * $_)

    [PSCustomObject]@{
        WeekNumber = 6 - $_
        ColumnName = $WeekStart.ToString("dd-MMM-yyyy")
        WeekStart  = $WeekStart
        WeekEnd    = $WeekStart.AddDays(7)
    }
}

# Data window to select items for the six week period: from oldest Saturday to end of latest Saturday-based week
$StartDate = ($Weeks | Sort-Object WeekNumber | Select-Object -First 1).WeekStart
$EndDate = ($Weeks | Sort-Object WeekNumber | Select-Object -Last 1).WeekEnd

# Find set of unique users who have had outstanding tasks
$UniqueUsersWithTasks = $ItemData | Sort-Object UPN -Unique | Select-Object UPN, UserName

# Find set of items logged for the last six weeks
[array]$SixWeeksTasks = $ItemData | Where-Object {[datetime]$_.RunDate -ge $StartDate -and [datetime]$_.RunDate -lt $EndDate}

# Create a hash table of tasks for each user. Without the hash table, we would scan $SixWeeksTasks six times per user
# or 600 times for 100 users.This way we do the scan one and direct lookups thereafter
$TasksByUPN = @{}
Foreach ($Task in $SixWeeksTasks) {
    If (-not $TasksByUPN.ContainsKey($Task.UPN)) {
        $TasksByUPN[$Task.UPN] = [System.Collections.Generic.List[Object]]::new()
    }
    $TasksByUPN[$Task.UPN].Add($Task)
}

# Create output file
$TaskWeeklySummary = [System.Collections.Generic.List[Object]]::new()

ForEach ($User in $UniqueUsersWithTasks) {
    # Pre-define columns in the desired order
    $Row = [ordered]@{
        UserPrincipalName = $User.UPN
        DisplayName = $User.UserName
    }
    Foreach ($Week in $Weeks) {
        $Row[$Week.ColumnName] = 0
    }
    $Row["Avg. Tasks Per Week"] = 0

    # Fetch tasks for the user for the last six weeks from the hash table
    $UPN = $User.UPN
    $UserTasks = If ($TasksByUPN.ContainsKey($UPN)) {
        $TasksByUPN[$UPN]
    } Else {
        @()
    }

    [int]$TotalTasks = 0
    Foreach ($Week in $Weeks) {
        $WeekRecord = $UserTasks | Where-Object {$_.UPN -eq $UPN -and [datetime]$_.RunDate -ge [datetime]$Week.WeekStart -and [datetime]$_.RunDate -lt [datetime]$Week.WeekEnd } | Sort-Object { [datetime]$_.RunDate } -Descending | Select-Object -First 1
        If ($WeekRecord) {
            $TaskCount = [int]$WeekRecord.Tasks
            $Row[$Week.ColumnName] = $TaskCount
            $TotalTasks += $TaskCount
        }
    }
    # Calculate average tasks per week and round to two decimal places
    $Row["Avg. Tasks Per Week"] = [math]::Round(($TotalTasks / 6), 2)
    $TaskWeeklySummary.Add([PSCustomObject]$Row)
}

# Select output columns in the preferred order
$OutputColumns = @("DisplayName") + ($Weeks | Select-Object -ExpandProperty ColumnName) + "Avg. Tasks Per Week"

# Display Data
$TaskWeeklySummary | Select-Object $OutputColumns | Sort-Object UserPrincipalName | Format-Table -AutoSize -Wrap

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.