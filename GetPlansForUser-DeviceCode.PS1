# GetPlansForUser-DeviceCode.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/GetPlansForUser-DeviceCode.PS1
# Sample script to show how to use the Graph to create a report of the Planner plans a user has access to using delegated permissions with a device code

# Get the user we're going to use to get plans
$UserCredential = Get-Credential -Message "Enter credentials for the account you want to use to fetch Plan information"
$User = $UserCredential.UserName

# Application (client) ID, tenant ID, resource and scope
$ClientID = "ded88173-911c-42a5-892b-26d7bea4c788" #GetPlansV2
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3478"
$Resource = "https://graph.microsoft.com/"
$Scope = "Group.Read.All Group.ReadWrite.All User.Read User.Read.All"

$CodeBody = @{ 
    resource  = $resource
    client_id = $clientId
    scope     = $scope }

# Get OAuth Code
$CodeRequest = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/devicecode" -Body $CodeBody

# Print Code to console
Write-Host "`n$($codeRequest.message)"

$Body = @{
    grant_type = "urn:ietf:params:oauth:grant-type:device_code"
    code       = $codeRequest.device_code
    client_id  = $clientId
}

# Get OAuth Token
$Token = $Null; $TokenRequest = $Null
while ([string]::IsNullOrEmpty($tokenRequest.access_token)) {
    $tokenRequest = try {
        Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/token" -Body $Body }
    catch {
        $errorMessage = $_.ErrorDetails.Message | ConvertFrom-Json
        # If not waiting for auth, throw error
        if ($errorMessage.error -ne "authorization_pending") {
            throw
        }
}}

$Token = $tokenRequest.access_token
$Headers = @{Authorization = "Bearer $Token"}

# Find all Microsoft 365 Groups the user belongs to
$Uri = "https://graph.microsoft.com/beta/users/$User/transitiveMemberOf"
$MemberOf = Invoke-WebRequest -Headers $Headers -Uri $Uri | ConvertFrom-Json
# Put the result in a list of groups we can process ;ater
$GroupsMemberOf = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $MemberOf.Value) {
   If ($M.GroupTypes -eq "Unified") { # Only select Microsoft 365 Groups
       $ReportLine = [PSCustomObject][Ordered]@{
            GroupId   = $M.Id
            Name      = $M.DisplayName } 
        $GroupsMemberOf.Add($ReportLine) }
}
# If there are any more groups to get, fetch them using the Nextlink given by the Graph and add them to the list
$NextLink = $MemberOf.'@Odata.NextLink'
While ($NextLink -ne $Null) { 
   Write-Host "Still processing..."
   $MemberOf = Invoke-WebRequest -Method GET -Uri $NextLink -ContentType "application/json" -Headers $Headers | ConvertFrom-JSon
   ForEach ($M in $MemberOf.Value) {
   If ($M.GroupTypes -eq "Unified") { # Only select Microsoft 365 Groups
       $ReportLine = [PSCustomObject][Ordered]@{
            GroupId   = $M.Id
            Name      = $M.DisplayName } 
        $GroupsMemberOf.Add($ReportLine) }
    }
   $NextLink = $MemberOf.'@Odata.NextLink'
} #End While

CLS
# We now have a list of Microsoft 365 Groups that the user belongs to, so we can check
# the groups to find out which have plans and report details of the plans we find.
$Activity = "Checking Plans for " + $User
$Report = [System.Collections.Generic.List[Object]]::new(); $PlanNumber = 0
$i = 0; $GroupCount = $GroupsMemberOf.Count
ForEach ($Group in $GroupsMemberOf) {
  $i++
  $ProgressBar = "Processing group " + $Group.Name + " (" + $i + " of " + $GroupCount + ")"
  Write-Progress -Activity $Activity -Status $ProgressBar -PercentComplete ($i/$GroupCount*100)
  $PlanURI = 'https://graph.microsoft.com/V1.0/groups/' + $Group.GroupId + '/planner/plans'
  $Plans = Invoke-WebRequest -Method GET -Uri $PlanURI -ContentType "application/json" -Headers $Headers  | ConvertFrom-Json
 
  ForEach ($Plan in $Plans.Value) {
        $PlanId = $Plan.Id
        $PlanNumber++
        $PlanCreated = Get-Date($Plan.CreatedDateTime) -format g
        $PlanOwner = $Plan.Owner # Microsoft 365 Group
        $PlanTitle = $Plan.Title 
        $BucketURI = 'https://graph.microsoft.com/v1.0/planner/plans/' + $PlanId + '/buckets/'
        $Buckets = Invoke-RestMethod -Method GET -Uri $BucketURI -ContentType "application/json" -Headers $Headers
        $NumberBuckets = $Buckets.Value.Count
        $TasksURI = 'https://graph.microsoft.com/v1.0/planner/plans/' + $PlanId + '/tasks/'
        $Tasks = Invoke-RestMethod -Method GET -Uri $TasksURI -ContentType "application/json" -Headers $Headers
        $NumberTasks = $Tasks.'@odata.count'
        [DateTime]$LastTask = "1-Jan-1999"
        # Grab some data about tasks like the date of the latest task and task completion stats
        $TasksNotStarted = 0; $TasksInProgress = 0; $TasksComplete = 0
        ForEach ($Task in $Tasks.Value) {
         If (-not [string]::IsNullOrEmpty($Task.CreatedDateTime)) { 
            [DateTime]$TaskCreated = Get-Date($Task.CreatedDateTime) -format g }
        If ($TaskCreated -gt $LastTask) {
            $LastTask = $TaskCreated; $LastTaskTitle = $Task.Title}
        Switch ($Task.PercentComplete) { #Generate stats for task completion status
            0   {$TasksNotStarted++}
           50   {$TasksInProgress++}
          100  {$TasksComplete++}
         } #End switch
        }  # End For
        If ($LastTask -eq "1-Jan-1999") { # Check how long it's been since a task was created in the plan
            $LastTaskDate = "No tasks"; $DaysSinceTask = "N/A"}
        Else {
           $LastTaskDate = Get-Date($LastTask) -format g
           $DaysSinceTask = (New-TimeSpan($LastTask)).Days  }
        # Write out information about plan
        $ReportLine = [PSCustomObject][Ordered]@{
            GroupId       = $Group.GroupId
            Name          = $Group.Name
            PlanId        = $PlanId
            Title         = $PlanTitle
            Created       = $PlanCreated
            Buckets       = $NumberBuckets
            Tasks         = $NumberTasks
            NotStarted    = $TasksNotStarted
            InProgress    = $TasksInProgress
            Complete      = $TasksComplete
            LastTaskDate  = $LastTaskDate
            DaysSinceTask = $DaysSinceTask
            LastTaskTitle = $LastTaskTitle }
        $Report.Add($ReportLine) } # End processing plan
 } # End Groups

Write-Host "All done. " $PlanNumber "plans found in" $GroupCount "Microsoft 365 Groups for user" $User
$Report | Select Name, Title, Created, Tasks, NotStarted, InProgress, Complete, LastTaskDate, DaysSinceTask, Buckets | Out-GridView 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
