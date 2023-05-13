# Report-Plans.PS1
# Simple script to show how to generate a simple report of the teams linked to Microsoft 365 Groups in a tenant
# The script finds all Microsoft 365 groups and checks each group to determine if it has any plans. If plans are
# found, the script retrieves the plan data and reports the number of tasks, how many tasks are completed, active, or
# in progress, and what tasks are in the various buckets.
#
# The script uses a registered Azure AD app for access. The app must have consent for these Graph application permissions:
# "Group.Read.All", "Directory.Read.All", "User.Read.All", "Tasks.Read.All"

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
$TenantId = "xxx"
$AppId = "xxx"
$AppSecret = "xxx"

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

$Version = "1.0"
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
       Write-Host ("Processing plan {0}" -f $Plan.title)
       $FirstTask = $NUll; $NewestTask = $Null; [int]$TaskCount = 0; [array]$LowTasks = $Null; [array]$MediumTasks = $Null; [array]$UrgentTasks = $Null
       [array]$ImportantTasks = $Null; [array]$NotStartedTasks = $Null; [array]$InProgressTasks = $Null; [array]$CompletedTasks = $Null
       $DaysSinceTaskCreated = "N/A"
       $Uri =  ("https://graph.microsoft.com/v1.0/planner/plans/{0}/tasks" -f $Plan.id)
       [array]$Tasks = Get-GraphData -Uri $Uri -AccessToken $Token
       If ($Tasks.value) {
          Write-Host ("Found {0} tasks in plan {1}" -f $Tasks.count, $Plan.title)
          $FirstTask = (Get-Date($Tasks.createdDateTime[($Tasks.count-1)]) -format g)
          $NewestTask = Get-Date($Tasks.createdDateTime[0]) -format g
          # How many days since a task was created in this plan?   
          $DaysSinceTaskCreated = (New-TimeSpan $NewestTask).Days
          [int]$TaskCount = $Tasks.count
          [array]$UrgentTasks = $Tasks | Where-Object {$_.Priority -eq 1}
          [array]$ImportantTasks = $Tasks | Where-Object {$_.Priority -eq 3}
          [array]$MediumTasks = $Tasks | Where-Object {$_.Priority -eq 5}
          [array]$LowTasks = $Tasks | Where-Object {$_.Priority -eq 9}
          [array]$NotStartedTasks = $Tasks | Where-Object {$_.percentComplete -eq 0}
          [array]$InProgressTasks = $Tasks | Where-Object {$_.percentComplete -eq 50}
          [array]$CompletedTasks = $Tasks | Where-Object {$_.percentComplete -eq 100}
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
                 Tasks      = $BucketCount.count 
                 Complete   = $BucketComplete.count 
                 Active     = $ActiveBucketTasks
                 "% Active" = $PercentActiveTasks
                 Plan       = $Plan.title
                 PlanId     = $Plan.Id
              }
              $BucketStats.Add($DataLine)
            }
       }
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
             BucketStats         = $BucketStats }
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
#  Report the basic statistics for the plan and bucket statistics if available
   $HtmlHeadingSection = ("<p><h2>Plans for Group <b><u>{0}</h2></b></u></p>" -f $G.Group)

   # Extract Plans
   $GroupPlans = $Report | Where-Object {$_.GroupId -eq $G.GroupId} | Select-Object Plan, Created, Tasks, "Oldest Task", "Newest Task", "Days Since Task", "Urgent Tasks", "Important Tasks", "Medium Tasks", "Low Tasks", "Completed Tasks", "In progress Tasks", "Not started Tasks", Buckets
   # Extract Bucket data for plan
   $GroupBuckets = $Report | Where-Object {$_.GroupId -eq $G.GroupId} | Select-Object -ExpandProperty BucketStats

   $HtmlReport = "<p>" + $HtmlReport + "<p>" + $HtmlHeadingSection

   ForEach ($P in $GroupPlans) {
     # Add the basic statistics for the plan
     $HtmlData = $P | ConvertTo-Html -Fragment
     $HtmlPlanHeading = ("<p><h3>Plan name: {0}</h3><p>" -f $P.Plan)
     # If it has any tasks, report the buckets
     If ($P.Tasks -gt 0) {
        $HtmlData2 = $GroupBuckets | Where-Object {$_.Plan -eq $P.Plan} | ConvertTo-Html -Fragment
        $HtmlHeadingBuckets = ("<p><h3>Bucket Analysis for the <u>{0}</u> plan</h3></p>" -f $P.Plan)
        $HtmlReport = $HtmlReport + "<p>" + $HtmlPlanHeading + $HtmlData + $HtmlHeadingBuckets + $HtmlData2 + "<h4></h5><p><p>"
    } Else {
        $HtmlReport = $HtmlReport + "<p>" + $HtmlPlanHeading + $HtmlData + "<p>"
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
