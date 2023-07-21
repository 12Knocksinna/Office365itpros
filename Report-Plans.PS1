# Report-Plans.PS1
# Simple script to show how to generate a simple report of the teams linked to Microsoft 365 Groups in a tenant
# The script finds all Microsoft 365 groups and checks each group to determine if it has any plans. If plans are
# found, the script retrieves the plan data and reports the number of tasks, how many tasks are completed, active, or
# in progress, and what tasks are in the various buckets.

# V2.0 adds reporting of incomplete tasks assigned to group members and an analysis of how well each group member
# is doing to clear up old tasks
#
# The script uses a registered Azure AD app for access. The app must have consent for these Graph application permissions:
# "Group.Read.All", "Directory.Read.All", "User.Read.All", "Tasks.Read.All"

function Generate-IndividualStatistics {
param (
       [parameter(Mandatory = $true)]
        $ActiveTasks 
       )

# function to take a set of active tasks and figure out how well each group member is doing

  $IndividualStats = [System.Collections.Generic.List[Object]]::new()
  ForEach ($Member in $GroupMembers) {
    [array]$Global:MemberTasks = $ActiveTasks | Where-Object {$_.Assignee -eq $Member.displayname}
    [array]$InProgressTasks = $MemberTasks | Where-Object {$_."Task Status" -eq "In progress"}
    [array]$NotStartedTasks = $MemberTasks | Where-Object {$_."Task Status" -eq "Not started"}
    $AvgDays = ($MemberTasks.DaysOld |Measure-Object -Average).average
    $DataLine = [PSCustomObject][Ordered]@{
       DisplayName        = $Member.displayName
       Tasks              = $MemberTasks.count
       "Not started"      = $NotStartedTasks.count
       "In progress"      = $InProgressTasks.count
       "Average days old" = ("{0:N2}" -f $AvgDays) }
      $IndividualStats.Add($DataLine) 
  }
  Return $IndividualStats
}

function Process-Tasks {
param (
       [parameter(Mandatory = $true)]
        $UncompletedTasks 
       )

# Return a set of uncompleted tasks for a plan so that we can analyze who needs to do more to close
# their tasks!
Write-Host ("Analyzing assignments for {0} uncompleted tasks" -f $UncompletedTasks.count)
$Assignments = [System.Collections.Generic.List[Object]]::new()
 
ForEach ($Task in $UncompletedTasks) {
   # Write-Host ("Processing task {0}" -f $Task.title)
   $TaskData = @{} 
   # Convert assignment data to a hash table for processing
   ($Task.assignments).psObject.Properties | ForEach { $TaskData[$_.Name] = $_.Value}
   [array]$TaskAssignments = $TaskData.Keys
   [array]$TaskAssignmentDates = $TaskData.Values.assignedDateTime
   [int]$i = 0; $DaysSinceAssignment = $Null
   ForEach ($Assignment in $TaskAssignments) {
      $Assignee = $GroupMembers | Where-Object {$_.Id -eq $Assignment} | Select-Object -ExpandProperty displayName
      If ($Assignee) {
           $AssignedDate = ($TaskAssignmentDates[$i]) 
           $DaysSinceAssignment = (New-TimeSpan $AssignedDate).Days
      } Else {
         $Assignee = "Unassigned"
         $AssignedDate = $Null
      }
      $i++; $Status = $Null; $Priority = $Null; $DaysTaskOld
      Switch ($Task.percentComplete) {
       "0"   { $Status = "Not started" }
       "50"  { $Status = "In progress"}
       "100" { $Status = "Complete" }
      }
      Switch ($Task.Priority) {
        "1"  { $Priority = "Urgent" }
        "3"  { $Priority = "Important" }
        "5"  { $Priority = "Medium" }
        "9"  { $Priority = "Low" }
      }
      If ($Task.createdDateTime) {
         $DaysTaskOld = (New-TimeSpan $Task.createdDateTime).days
      }
      If ($Task.dueDateTime) {
         $TaskDueDate = Get-Date($Task.dueDateTime) -format g
      }
      if ($AssignedDate) {
         $AssignedDate = Get-Date($AssignedDate) -format g
      }
      $DataLine = [PSCustomObject][Ordered]@{
         Plan                = $Task.planId
         PlanTitle           = $PlanTitle
         TaskId              = $Task.id
         Title               = $Task.title
         Bucket              = ($BucketsTable[$Task.bucketId])
         Created             = Get-Date ($Task.createdDateTime) -format g
         DueDate             = $TaskDueDate
         percentComplete     = $Task.percentComplete
         DaysOld             = $DaysTaskOld
         "Task Status"       = $Status
         Priority            = $Priority
         Assignee            = $Assignee
         AssignedDate        = ($AssignedDate) 
         DaysSinceAssignment = $DaysSinceAssignment
      }
      $Assignments.Add($DataLine)
    } #EndForeach Assignment            
  } #End Foreach Tasks
  Return $Assignments
}

function Get-GraphData {
# Based on https://danielchronlund.com/2018/11/19/fetch-data-from-microsoft-graph-with-powershell-paging-support/
# GET data from Microsoft Graph.
    param (
        [parameter(Mandatory = $true)]
        $AccessToken,

        [parameter(Mandatory = $true)]
        $Uri
    )

    # Check if authentication was successful.
    if ($AccessToken) {
    $Headers = @{
         'Content-Type'  = "application\json"
         'Authorization' = "Bearer $AccessToken" 
         'ConsistencyLevel' = "eventual"  }

        # Create an empty array to store the result.
        $QueryResults = @()

        # Invoke REST method and fetch data until there are no pages left.
        do {
            $Results = ""
            $StatusCode = ""

            do {
                try {
                    $Results = Invoke-RestMethod -Headers $Headers -Uri $Uri -UseBasicParsing -Method "GET" -ContentType "application/json"

                    $StatusCode = $Results.StatusCode
                } catch {
                    $StatusCode = $_.Exception.Response.StatusCode.value__

                    if ($StatusCode -eq 429) {
                        Write-Warning "Got throttled by Microsoft. Sleeping for 45 seconds..."
                        Start-Sleep -Seconds 45
                    }
                    else {
                        Write-Error $_.Exception
                    }
                }
            } while ($StatusCode -eq 429)

            if ($Results.value) {
                $QueryResults += $Results.value
            }
            else {
                $QueryResults += $Results
            }

            $uri = $Results.'@odata.nextlink'
        } until (!($uri))

        # Return the result.
        $QueryResults
    }
    else {
        Write-Error "No Access Token"
    }
}

function Get-AccessToken {
# function to return an Oauth access token

# Define the values applicable for the application used to connect to the Graph
$TenantId = "xxxxx3f-14fc-43a2-9a7a-d2e27f4f3478"
$AppId = "xxxxx-026b-4c29-ab81-fa1264139c9c"
$AppSecret = "szM8Q~dfpy9VvLqWGJW8Wr1SPdVby6TpWPryxb5M"

# Construct URI and body needed for authentication
$Uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"
}

# Get OAuth 2.0 Token
  $TokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
# Unpack Access Token
  $Global:Token = ($tokenRequest.Content | ConvertFrom-Json).access_token

   Write-Host ("Retrieved new access token at {0}" -f (Get-Date)) -foregroundcolor red

   $Global:Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $Token" 
            'ConsistencyLevel' = "eventual" }
	    
  Return $Token
}

# Start Processing

$Version = "2.0"
$HtmlReportFile = "c:\temp\GroupsPlans.html"
$CSVReportFile = "c:\temp\GroupPlans.CSV"

# Get access token (hopefully with the correct permissions...)
$Token = Get-AccessToken

# Fetch organization information
$Uri = "https://graph.microsoft.com/v1.0/organization"
[array]$OrgData = Get-GraphData -Uri $Uri -AccessToken $Token

# Get the Microsoft 365 groups in the tenant
$Uri = "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(a:a eq 'unified')"
[array]$Groups = Get-GraphData -AccessToken $Token -Uri $uri
If (!($Groups)) { Write-Host "Can't find any groups, so there's no plans to find either..."; break }

$Groups = $Groups | Sort-Object displayName
Write-Host ("Processing {0} groups" -f $Groups.count)

# Check each group for plans and process those plans
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Group in $Groups) {
  $Uri = ("https://graph.microsoft.com/v1.0/groups/{0}/planner/plans" -f $Group.Id)
  [array]$Plans = Get-GraphData -Uri $Uri -AccessToken $Token
  If ($Plans.container) {
     Write-Host ("{0} plans found in group {1}" -f $Plans.count, $Group.displayName) 
     ForEach ($Plan in $Plans) {
       $Global:PlanTitle = $Plan.title
       Write-Host ("Processing plan {0}" -f $PlanTitle)
       $FirstTask = $NUll; $NewestTask = $Null; [int]$TaskCount = 0; [array]$LowTasks = $Null; [array]$MediumTasks = $Null; [array]$UrgentTasks = $Null
       [array]$ImportantTasks = $Null; [array]$NotStartedTasks = $Null; [array]$InProgressTasks = $Null; [array]$CompletedTasks = $Null
       [array]$TaskAssignments = $Null; [array]$UncompletedTasks = $Null
       $DaysSinceTaskCreated = "N/A"
       # Get group members so that we can track assignments
       $Uri = ("https://graph.microsoft.com/v1.0/groups/{0}/members" -f $Group.Id)
       [array]$Global:GroupMembers = Get-GraphData -Uri $Uri -AccessToken $Token | Select-Object Id, displayName
       $Uri =  ("https://graph.microsoft.com/v1.0/planner/plans/{0}/tasks" -f $Plan.id)
       [array]$Tasks = Get-GraphData -Uri $Uri -AccessToken $Token
       If ($Tasks.value) {
          Write-Host ("Found {0} tasks in plan {1}" -f $Tasks.count, $Plan.title)
          $FirstTask = (Get-Date($Tasks.createdDateTime[($Tasks.count-1)]) -format g)
          $NewestTask = Get-Date($Tasks.createdDateTime[0]) -format g
          # How many days since a task was created in this plan?   
          $DaysSinceTaskCreated = (New-TimeSpan $NewestTask).Days
          [int]$TaskCount = $Tasks.count
          # Process each task to find assignment data
          [array]$UrgentTasks = $Tasks | Where-Object {$_.Priority -eq 1}
          [array]$ImportantTasks = $Tasks | Where-Object {$_.Priority -eq 3}
          [array]$MediumTasks = $Tasks | Where-Object {$_.Priority -eq 5}
          [array]$LowTasks = $Tasks | Where-Object {$_.Priority -eq 9}
          [array]$NotStartedTasks = $Tasks | Where-Object {$_.percentComplete -eq 0}
          [array]$InProgressTasks = $Tasks | Where-Object {$_.percentComplete -eq 50}
          [array]$CompletedTasks = $Tasks | Where-Object {$_.percentComplete -eq 100}
          # Get bucket data
          $Uri =  ("https://graph.microsoft.com/v1.0/planner/plans/{0}/buckets" -f $Plan.id)
          [array]$Buckets = Get-GraphData -Uri $Uri -AccessToken $Token
          $BucketStats = [System.Collections.Generic.List[Object]]::new()
            ForEach ($Bucket in $Buckets) {
              [array]$BucketTasks = $Tasks | Where-Object {$_.bucketId -eq $Bucket.id}
              [array]$BucketComplete = $Tasks | Where-Object {$_.percentComplete -eq 100 -and $_.bucketId -eq $Bucket.id}
              [int]$ActiveBucketTasks = ($BucketTasks.count - $BucketComplete.count)
              If ($ActiveBucketTasks -gt 0) {
                 $PercentActiveTasks = ($ActiveBucketTasks/$BucketTasks.count).toString("P")
              } Else {
                 $PercentActiveTasks = "N/A" }
              $DataLine = [PSCustomObject][Ordered]@{
                 Bucket     = $Bucket.name
                 Tasks      = $BucketTasks.count 
                 Complete   = $BucketComplete.count 
                 Active     = $ActiveBucketTasks
                 "% Active" = $PercentActiveTasks
                 Plan       = $Plan.title
                 PlanId     = $Plan.Id
              }
              $BucketStats.Add($DataLine)
            }
          $Global:BucketsTable = @{}
          ForEach ($Bucket in $Buckets) { $BucketsTable.Add([string]$Bucket.id,[string]$Bucket.name) }
          # Get assignments for all uncompleted tasks
          [array]$UncompletedTasks = $InProgressTasks + $NotStartedTasks
          If ($UncompletedTasks.count -gt 0) {
             [array]$TaskAssignments = Process-Tasks -UncompletedTasks $UncompletedTasks
             # Make sure that we have plan data in all records
             $TaskAssignments = $TaskAssignments | Where-Object {$_.Plan -ne $Null}
          }
       }
       $Buckets = $Buckets | Sort-Object Name 
       # Generate report line for the plan
       $ReportLine = [PSCustomObject][Ordered]@{
             Plan                = $Plan.title
             Created             = Get-Date($plan.createddatetime) -format g
             Tasks               = $Taskcount
             "Oldest task"       = $FirstTask
             "Newest task"       = $NewestTask
             "Days since task"   = $DaysSinceTaskCreated
             "Urgent tasks"      = $UrgentTasks.count
             "Important tasks"   = $ImportantTasks.count
             "Medium tasks"      = $MediumTasks.count
             "Low tasks"         = $LowTasks.count
             "Completed tasks"   = $CompletedTasks.count
             "In progress tasks" = $InProgressTasks.count
             "Not started tasks" = $NotStartedTasks.count
             Buckets             = ($Buckets.name -join ", ")
             PlanId              = $Plan.Id
             Group               = $Group.displayName
             GroupId             = $Group.Id 
             TaskStats           = $TaskAssignments
             BucketStats         = $BucketStats 
	     GroupMembers        = $GroupMembers        }
      $Report.Add($ReportLine) 
    } # End Foreach Plan
  }  # End if 
}

# Find the set of Microsoft 365 groups with plans
$GroupsWithPlans = $Report | Select-Object Group, GroupId | Sort-Object GroupId -Unique | Sort-Object Group

$CountOfPlans = ($Report.PlanId | Sort-Object -Unique).count
$CountOfTasks = ($Report.Tasks | Measure-Object -Sum).sum
$CountOfCompletedTasks = ($Report."Completed Tasks" | Measure-Object -Sum).sum
$CountOfActiveTasks = $CountOfTasks - $CountOfCompletedTasks
$PercentCompletedTasks = ($CountOfCompletedTasks/$CountOfTasks).toString("P")

Write-Host "Generating analysis..."

# Generate the report files
$HtmlHeading ="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.pass{background: #B7EB83;}
	   td.warn{background: #FFF275;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Microsoft 365 Groups and Plans Report</h1></p>
           <p><h3>Generated: " + (Get-Date -format 'dd-MMM-yyyy hh:mm tt') + "</h3></p></div>"

$HtmlReport = $HtmlHeading

ForEach ($G in $GroupsWithPlans) {
   # Report the basic statistics for the plan and bucket statistics if available
   $HtmlHeadingSection = ("<p><h2>Plans for Group <b><u>{0}</h2></b></u></p>" -f $G.Group)
   # Get the group members of the plan so that we can report individual assignments. Because a group can host multiple plans, we select the first record
   [array]$Global:GroupMembers = $Report | Where-Object {$_.GroupId -eq $G.GroupId} | Select-Object -First 1 |Select-Object -ExpandProperty GroupMembers

   # Extract Plans
   $GroupPlans = $Report | Where-Object {$_.GroupId -eq $G.GroupId} | Select-Object Plan, Created, Tasks, "Oldest Task", "Newest Task", "Days Since Task", "Urgent Tasks", "Important Tasks", "Medium Tasks", "Low Tasks", "Completed Tasks", "In progress Tasks", "Not started Tasks", Buckets, PlanId
   # Extract Bucket data for plan
   $GroupBuckets = $Report | Where-Object {$_.GroupId -eq $G.GroupId} | Select-Object -ExpandProperty BucketStats | Sort-Object Bucket
   # Extract assignments for uncompleted tasks
   [array]$GroupAssignments = $Report | Where-Object {$_.GroupId -eq $G.GroupId} | Select-Object -ExpandProperty TaskStats | Sort-Object Assignee

   $HtmlReport = "<p>" + $HtmlReport + "<p>" + $HtmlHeadingSection

   ForEach ($P in $GroupPlans) {
     # Add the basic statistics for the plan
     $IndividualStats = $Null
     $HtmlData = $P | ConvertTo-Html -Fragment
     $HtmlPlanHeading = ("<p><h3>Plan name: {0}</h3><p>" -f $P.Plan)
     # If it has any tasks, report the buckets
     If ($P.Tasks -gt 0) {
        $HtmlData2 = $GroupBuckets | Where-Object {$_.PlanId -eq $P.PlanId} | Sort-Object Bucket -Unique | ConvertTo-Html -Fragment
        $HtmlHeadingBuckets = ("<p><h3>Bucket Analysis for the <u>{0}</u> plan</h3></p>" -f $P.Plan)
        $HtmlReport = $HtmlReport + "<p>" + $HtmlPlanHeading + $HtmlData + $HtmlHeadingBuckets + $HtmlData2 + "<h4></h5><p><p>"
    } Else {
        $HtmlReport = $HtmlReport + "<p>" + $HtmlPlanHeading + $HtmlData + "<p>"
    }
    If ($P.Tasks -gt $P.'Completed Tasks' -and $GroupAssignments) { # We have some uncompleted tasks to report for assigned members
       [array]$Global:ActiveTasks = $GroupAssignments | Where-Object {$_.Plan -eq $P.PlanId} | Select-Object PlanTitle, Title, Assignee, Bucket, StartDate, DueDate, AssignedDate, "Task Status", Priority, DaysOld, DaysSinceAssignment 
       If ($ActiveTasks) {
          $HtmlData3 = $ActiveTasks | ConvertTo-html -Fragment
          $HtmlHeadingAssignments = ("<p><h3>Incomplete Tasks for the <u>{0}</u> plan</h3></p>" -f $P.Plan)
          $HtmlReport = $HtmlReport + "<p>" + $HtmlHeadingAssignments + $HtmlData3 + "<h4></h5><p><p>"
          $IndividualStats = Generate-IndividualStatistics -ActiveTasks $ActiveTasks
          $HtmlData4 = $IndividualStats | ConvertTo-html -Fragment   
          $HtmlHeadingIndividualStats = ("<p><h3>Indivdual Member Statistics for Incomplete Tasks for the <u>{0}</u> plan</h3></p>" -f $P.Plan)
          $HtmlReport = $HtmlReport + "<p>" + $HtmlHeadingIndividualStats + $HtmlData4 + "<h4></h5><p><p>"
       }
    }
}

} #End reporting plans for the groups
   
# Create the HTML report
$Htmltail = "<p><p>Report created for: " + ($OrgData.DisplayName) + "</p><p>" +
             "<p>Number of Microsoft 365 Groups with plans:   " + $GroupsWithPlans.count + "</p>" +
             "<p>Number of individual Plans:                  " + $CountOfPlans + "</p>" +
             "<p>Number of individual Tasks:                  " + $CountOfTasks + "</p>" +
             "<p>Number of Completed Tasks:                   " + $CountOfCompletedTasks + "</p>" +
             "<p>Percentage of Completed Tasks:               " + $PercentCompletedTasks + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------" +
             "<p>Microsoft 365 Groups and Plans <b>" + $Version + "</b>"	
$HtmlReport = $HtmlHead + $HtmlReport + $HtmlTail
$HtmlReport | Out-File $HtmlReportFile  -Encoding UTF8
$Report | Export-CSV $CSVReportFile -Notypeinformation

CLS
Write-Host "Finishing processing plans. Here's what we found"
Write-Host "------------------------------------------------"
Write-Host ""
Write-Host ("Microsoft 365 Groups with Plans: {0}" -f $GroupsWithPlans.count)
Write-Host ("Number of individual Plans:      {0}" -f $CountOfPlans)
Write-Host ("Number of individual Tasks:      {0}" -f $CountOfTasks)
Write-Host ("Number of Completed Tasks:       {0}" -f $CountOfCompletedTasks)
Write-Host ("Percentage of Completed Tasks:   {0}" -f $PercentCompletedTasks)
Write-Host ""
Write-Host ("The output files are {0} (HTML) and {1} (CSV)" -f $HtmlReportFile, $CSVReportFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
