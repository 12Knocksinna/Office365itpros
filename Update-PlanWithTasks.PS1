# Update-PlanWithTasks.PS1
# An example of using the Microsoft Graph PowerShell SDK to read message center notifications
# and use that data to update tasks in a Planner Plan
# V1.0 31-Jan-2024

Connect-MgGraph -NoWelcome -Scopes Group.Read.All, Tasks.ReadWrite

$ProcessedDataFile = "c:\temp\Announcements.csv"
$GroupId = '78b47932-b35f-4b26-94c2-3228cb234b07'
$TargetPlanName = 'Admin Task Assignment'
$CheckDate = (Get-Date).AddDays(-7)
$Now = Get-Date

# Attempt to import the file of previously processed data. If we have a file, set the window to query for new 
# announcements to 7 days. If not, use 60 days
[array]$ProcessedData = Import-CSV $ProcessedDataFile
If ($ProcessedData) {
    $CheckDate = (Get-Date).AddDays(-7)
} Else {
    $CheckDate = (Get-Date).AddDays(-60)
}

Write-Host "Finding message center announcement posts"
# Fetch all available announcements to create a new data file of previously processed data. Then trim
# the data to find the announcements created in the window we want to process
[array]$AllAnnouncements = Get-MgServiceAnnouncementMessage -Sort 'LastmodifiedDateTime desc' -All 
[array]$AnnouncementsForPeriod = $AllAnnouncements | Where-Object {$_.StartDateTime -as [datetime] -gt $CheckDate} 
If ($Announcements.Count -eq 0) {
    Write-Host ("No new message center posts found since {0} - exiting" -f $CheckDate)
    Break
} Else {
    Write-Host ("{0} message center posts found to process..." -f $AnnouncementsForPeriod.count)
}

# Now find what announcements we need to process (not seen before or not changed since we processed)
[array]$Announcements = $null
ForEach ($Announcement in $AnnouncementsForPeriod) {
    If ($Announcement.Id -notin $ProcessedData.Id) {
        $Announcements += $Announcement
    }
}

Write-Host "Checking for target plan and buckets..."
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
$GeneralBucket =  $Buckets | Where-Object Name -match 'For assignment'
$PowerBucket = $Buckets | Where-Object Name -match 'Power Platform'
$AdminBucket = $Buckets | Where-Object Name -match 'Administration'
$EntraBucket = $Buckets | Where-Object Name -match 'Entra'
$OfficeAppsBucket = $Buckets | Where-Object Name -match 'Office Apps'
$PurviewBucket = $Buckets | Where-Object Name -match 'Purview'

Write-Host "Finding plan members and creating assignee details..."
[array]$PlanMembers = Get-MgGroupMember -GroupId $GroupId

# Populate the assignees from the group members
$ExchangeAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'lotte.vetler@office365itpros.com'
$TeamsAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'michelle.dubois@office365itpros.com'
$AdminAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'Marty.King@office365itpros.com'
$SharePointAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'Rene.artois@office365itpros.com'
$PowerAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'hans.flick@office365itpros.com'
$PlannerAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'mimie@office365exchangebook.com'
$EntraAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'brian.weakliam@office365itpros.com'
$OfficeAppsAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'hans.geering@office365itpros.com'
$GeneralAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'kim.Akers@office365itpros.com'
$PurviewAssignee = $PlanMembers.additionalProperties | Where-Object mail -match 'James.A.Abrahams@office365itpros.com'

# Planner likes to assign people via user id, so we make sure that we have it
$PowerAssignee = (Get-MgUser -UserId $PowerAssignee.userPrincipalName)
$TeamsAssignee = (Get-MgUser -UserId $TeamsAssignee.userPrincipalName)
$SharePointAssignee = (Get-MgUser -UserId $SharePointAssignee.userPrincipalName)
$AdminAssignee = (Get-MgUser -UserId $AdminAssignee.userPrincipalName)
$ExchangeAssignee = (Get-MgUser -UserId $ExchangeAssignee.userPrincipalName)
$PlannerAssignee = (Get-MgUser -UserId $PlannerAssignee.userPrincipalName)
$EntraAssignee = (Get-MgUser -UserId $EntraAssignee.userPrincipalName)
$OfficeAppsAssignee = (Get-MgUser -UserId $OfficeAppsAssignee.userPrincipalName)
$GeneralAssignee = (Get-MgUser -UserId $GeneralAssignee.userPrincipalName)
$PurviewAssignee = (Get-MgUser -UserId $PurviewAssignee.userPrincipalName)

# Some generic settings needed for an assignment
$GenericTaskData = @{}
$GenericTaskData.Add("@odata.type", "#microsoft.graph.plannerAssignment")
$GenericTaskData.Add("orderHint"," !")

$GenericCategoryData = @{}
$GenericCategoryData.Add("@odata.type", "#microsoft.graph.plannerAppliedCategories")

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Announcement in $Announcements) {

    [array]$Services = $Announcement.Services
    $TaskTitle = ("[{0}] {1} [{2}]" -f ($Services -join ","), $Announcement.Title, $Announcement.Id)
    Write-Host ("Processing task {0}" -f $Announcement.Title)

    $AssignedUserId = $AdminAssignee.Id
    
    # Figure out who should be assigned
    Switch -wildcard ($Services) {
        "Microsoft Teams" {
            $AssignedUserId = $TeamsAssignee.Id
            $AssignedName = $TeamsAssignee.DisplayName
            $TargetBucket = $TeamsBucket.Id
            $BucketName = "Teams"
            $Category = "category4"
        }
        "*Power*" {
            $AssignedUserId = $PowerAssignee.Id
            $AssignedName = $PowerAssignee.DisplayName
            $TargetBucket = $PowerBucket.Id
            $BucketName = "Power Platform"
            $Category = "category1"
        }
        "Exchange*" {
            $AssignedUserId = $ExchangeAssignee.Id
            $AssignedName = $ExchangeAssignee.DisplayName
            $TargetBucket = $ExchangeBucket.Id
            $BucketName = "Exchange Online"
            $Category = "category2"
        }
        "SharePoint*"{
            $AssignedUserId = $SharePointAssignee.Id
            $AssignedName = $SharePointAssignee.DisplayName
            $TargetBucket = $SharePointBucket.Id
            $BucketName = "SharePoint Online"
            $Category = "category7"
        }
        "Plan*" {
            $AssignedUserId = $PlannerAssignee.Id
            $AssignedName = $PlannerAssignee.DisplayName
            $TargetBucket = $PlannerBucket.Id
            $BucketName = "Planner"
            $Category = "category8"
        }
        "Microsoft 365*" {
            $AssignedUserId = $AdminAssignee.Id
            $AssignedName = $AdminAssignee.DisplayName
            $TargetBucket = $AdminBucket.Id
            $BucketName = "Administration"
            $Category = "category3"
        }
        "Entra*" {
            $AssignedUserId = $EntraAssignee.Id
            $AssignedName = $EntraAssignee.DisplayName
            $TargetBucket = $EntraBucket.Id
            $BucketName  = "Entra"
            $Category = "category5"
        }
        "*Office*" {
            $AssignedUserId = $OfficeAppsAssignee.Id
            $AssignedName = $OfficeAppsAssignee.DisplayName
            $TargetBucket = $OfficeAppsBucket.Id
            $BucketName   = "Office Apps"
            $Category = "category6"
        }
        "*Purview*" {
            $AssignedUserId = $PurviewAssignee.Id
            $AssignedName = $PurviewAssignee.DisplayName
            $TargetBucket = $PurviewBucket.Id
            $BucketName   = "Purview"
            $Category = "category11"
        }
        default {
            $AssignedUserId = $GeneralAssignee.Id
            $AssignedName = $GeneralAssignee.DisplayName
            $TargetBucket = $GeneralBucket.Id
            $BucketName   = "For assignment"
            $Category = "category10" 
        }
    }

    $TaskAssignments = @{}
    $TaskAssignments.Add($AssignedUserId, $GenericTaskData)

    # Create a text only version of the notes
    $Body = $Announcement | Select-Object -ExpandProperty Body
    $HTML = New-Object -Com "HTMLFile"
    $HTML.write([ref]$body.content)
    $TextOnly = $HTML.body.innerText
    
    # Add announcement properties to the top of the notes
    $AnnouncementStartDate = Get-Date $Announcement.StartDateTime -format "dd-MMM-yyyy"
    $AnnouncementLastModifiedDate = Get-Date $Announcement.LastmodifiedDateTime -format "dd-MMM-yyyy"
    $TextHeading = ("Message ID: {0} `nPublished date: {1} `nLast updated date: {2} `nCategory: {3} `n" -f `
        $Announcement.Id, $AnnouncementStartDate, $AnnouncementLastModifiedDate, $Announcement.Category)
    [array]$Tags = $Announcement.Tags | Sort-Object
    $TextBody = ("{0}Tags: {1} `n`n{2}" -f $TextHeading, ($Tags -join ", "), $TextOnly)

    $TaskDescription = @{}
    $TaskDescription.Add('description', $TextBody)

    $TaskCategoryParams = @{
        AppliedCategories = @{
            ($Category) = $true
        }
    }

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
        Update-MgPlannerTask -PlannerTaskId $NewTask.Id -BodyParameter $TaskCategoryParams -IfMatch $NewTask.additionalProperties.'@odata.etag'
        Update-MgPlannerTaskDetail -PlannerTaskId $TaskDetails.Id -BodyParameter $TaskDescription -IfMatch $TaskDetails.additionalProperties.'@odata.etag'
        Write-Host ("Task {0} assigned to {1}" -f $TaskTitle, $AssignedName)
        $DataLine = [PSCustomObject][Ordered]@{
            TaskId          = $Announcement.ID
            Title           = $TaskTitle
            Assignedto      = $AssignedName
            Bucketid        = $BucketName
        }
        $Report.Add($DataLine)
    }
}

$Report | Out-GridView
$AllAnnouncements | Export-CSV -NoTypeInformation $ProcessedDataFile

# Export the data that we have processed

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
