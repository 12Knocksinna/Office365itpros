# Update-PlanWithTasks.PS1
# An example of using the Microsoft Graph PowerShell SDK to read message center notifications
# and use that data to update tasks in a Planner Plan
# V1.0 31-Jan-2024

Connect-MgGraph -NoWelcome -Scopes Group.Read.All, Tasks.ReadWrite

$GroupId = '33b07753-efc6-47f5-90b5-13bef01e25a6'
$TargetPlanName = 'Admin Center Tasks'
$CheckDate = (Get-Date).AddDays(-7)
$Now = Get-Date

Write-Host "Finding message center announcement posts"
[array]$Announcements = Get-MgServiceAnnouncementMessage -Sort 'LastmodifiedDateTime desc' -All | `
  Where-Object {$_.StartDateTime -as [datetime] -gt $CheckDate} 
If ($Announcements.Count -eq 0) {
    Write-Host ("No new message center posts found since {0} - exiting" -f $CheckDate)
    Break
} Else {
    Write-Host ("{0} message center posts found to process..." -f $Announcements.count)
}

[array]$Plans = Get-MgGroupPlannerPlan -GroupId $GroupId
$TargetPlan = $Plans | Where-Object Title -Match $TargetPlanName
If (!$TargetPlan) {
    Write-Host ("Unable to find the target plan ({0}) - exiting" -f $TargetPlanName)
    Break
}

[array]$Buckets = Get-MgPlannerPlanBucket -PlannerPlanId $TargetPlan.Id
If (!$Buckets) {
    Write-Host "No buckets found in the target plan - exiting"
    Break
}
$ExchangeBucket = $Buckets | Where-Object Name -match 'Exchange Online'
$SharePointBucket = $Buckets | Where-Object Name -match 'SharePoint Online'
$TeamsBucket =  $Buckets | Where-Object Name -match 'Teams'
$PlannerBucket =  $Buckets | Where-Object Name -match 'Planner'
$GeneralBucket =  $Buckets | Where-Object Name -match 'To be assigned'
$PowerBucket = $Buckets | Where-Object Name -match 'Power Platform'
$AdminBucket = $Buckets | Where-Object Name -match 'Administration'
$EntraBucket = $Buckets | Where-Object Name -match 'Entra'

[array]$PlanMembers = Get-MgGroupMember -GroupId $GroupId

# Populate the assignees from the group members
$ExchangeAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'michel@eightwone.com'
$TeamsAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'Tony.Redmond@redmondassociates.org'
$AdminAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'pro@keepit.com'
$SharePointAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'jcgonzalez@itechcs.onmicrosoft.com'
$PowerAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'christina@cmportalsolutions.com'
$PlannerAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'Tony.Redmond@redmondassociates.org'
$EntraAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'bdesmond@ravenswoodtechnology.com'

# Planner likes to assign people via user id, so we make sure that we have it
$PowerAssignee = (Get-MgUser -UserId $PowerAssignee.userPrincipalName)
$TeamsAssignee = (Get-MgUser -UserId $TeamsAssignee.userPrincipalName)
$SharePointAssignee = (Get-MgUser -UserId $SharePointAssignee.userPrincipalName)
$AdminAssignee = (Get-MgUser -UserId $AdminAssignee.userPrincipalName)
$ExchangeAssignee = (Get-MgUser -UserId $ExchangeAssignee.userPrincipalName)
$PlannerAssignee = (Get-MgUser -UserId $PlannerAssignee.userPrincipalName)
$EntraAssignee = (Get-MgUser -UserId $EntraAssignee.userPrincipalName)

# Some generic settings needed for an assignment
$GenericTaskData = @{}
$GenericTaskData.Add("@odata.type", "#microsoft.graph.plannerAssignment")
$GenericTaskData.Add("orderHint"," !")

ForEach ($Announcement in $Announcements) {

    [array]$Services = $Announcement.Services
    $TaskTitle = ("[{0}] {1} [{2}]" -f ($Services -join ","), $Announcement.Title, $Announcement.Id)
    Write-Host ("Processing taskl {0}" -f $TaskTile)
    $Body = $Announcement | Select-Object -ExpandProperty Body

    $AssignedUserId = $AdminAssignee.Id
    $TargetBucket = $GeneralBucket.Id
    # Figure out who should be assigned
    Switch -wildcard ($Services) {
        "Microsoft Teams" {
            $AssignedUserId = $TeamsAssignee.Id
            $AssignedName = $TeamsAssignee.DisplayName
            $TargetBucket = $TeamsBucket.Id
        }
        "Power*" {
            $AssignedUserId = $PowerAssignee.Id
            $AssignedName = $PowerAssignee.DisplayName
            $TargetBucket = $PowerBucket.Id
        }
        "Exchange*" {
            $AssignedUserId = $ExchangeAssignee.Id
            $AssignedName = $ExchangeAssignee.DisplayName
            $TargetBucket = $ExchangeBucket.Id
        }
        "SharePoint*"{
            $AssignedUserId = $SharePointAssignee.Id
            $AssignedName = $SharePointAssignee.DisplayName
            $TargetBucket = $SharePointBucket.Id
        }
        "Plan*" {
            $AssignedUserId = $PlannerAssignee.Id
            $AssignedName = $PlannerAssignee.DisplayName
            $TargetBucket = $PlannerBucket.Id
        }
        "Microsoft 365*" {
            $AssignedUserId = $AdminAssignee.Id
            $AssignedName = $AdminAssignee.DisplayName
            $TargetBucket = $AdminBucket.Id
        }
        "Entra*" {
            $AssignedUserId = $EntraAssignee.Id
            $AssignedName = $EntraAssignee.DisplayName
            $TargetBucket = $EntraBucket.Id
        }
    }

    $TaskAssignments = @{}
    $TaskAssignments.Add($AssignedUserId, $GenericTaskData)

    $HTML = New-Object -Com "HTMLFile"
    $HTML.write([ref]$body.content)
    $TextOnly = $HTML.body.innerText

    $TaskDescription = @{}
    $TaskDescription.Add('description', $TextOnly)

    $TaskParameters = @{}
    $TaskParameters.Add('planId',$TargetPlan.Id)
    $TaskParameters.Add('bucketid',$TargetBucket)
    $TaskParameters.Add('title', $TaskTitle)
    $TaskParameters.Add('assignments', $TaskAssignments)
    $TaskParameters.Add('priority', '5')
    $TaskParameters.Add('startDateTime', $Now)
  
    
    $NewTask = New-MgPlannerTask -BodyParameter $TaskParameters
    If ($NewTask) {
        $TaskDetails = Get-MgPlannerTaskDetail -PlannerTaskId $NewTask.Id
        Update-MgPlannerTaskDetail -PlannerTaskId $TaskDetails.Id -BodyParameter $TaskDescription -IfMatch $TaskDetails.additionalProperties.'@odata.etag'
        Write-Host ("Task {0} assigned to {1}" -f $TaskTitle, $AssignedName)
    }
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
