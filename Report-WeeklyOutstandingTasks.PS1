# Report-WeeklyOutstandingTasks.PS1
# This script generates a weekly report of incomplete tasks for a user and sends the list to them via email

# V1.0 1-Apr-2026
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-WeeklyOutstandingTasks.PS1

# Permissions needed to run the script:
# - User.Read.All (to read user details)
# - Group.Read.All (to read group details and membership)
# - GroupMember.Read.All (to read group membership)
# - Tasks.Read.All (to read the tasks assigned to users by an app). If running interactively, Tasks.Read is sufficient.

# The script is intended to run as a scheduled Azure Automation runbook. Make sure that the Azure automation account has the
# necessary permissions to run the script and access the Microsoft Graph API.

# Modules needed for runtime environment:
# - Microsoft.Graph.Authentication (to connect to Microsoft Graph)
# - Microsoft.Graph.Users (to read user details)
# - Microsoft.Graph.Groups (to read group details and membership)   
# - Microsoft.Graph.Planner (to read the tasks assigned to users)
# - Microsoft.Graph.Users.Actions (to send the email with the report)
# - ImportExcel (optional, to generate Excel worksheets for the report attachments)

[array]$RequiredScopes = "Group.Read.All", "GroupMember.Read.All", "Tasks.Read", "Mail.Send", "User.Read.All"
$Interactive = $false

# Determine if we're interactive or not
If ([Environment]::UserInteractive) { 
    # We're running interactively...
    Clear-Host
    Write-Host "Script running interactively... connecting to the Graph" -ForegroundColor Yellow
    Connect-MgGraph -NoWelcome
    $Interactive = $true
    # Email address to use when sending email from interactive session
    $MsgFrom = (Get-MgContext).Account    
} Else { 
    # We're not, so likely in Azure Automation
    Write-Output "Executing the runbook to send email about incomplete tasks..." 
    Connect-MgGraph -Identity -NoWelcome
    # Email address to use when sending email from Azure Automation - update for your tenant
    $MsgFrom = "no-reply@office365itpros.com"
}

# Check that we have the right permissions - in Azure Automation, we assume that the automation account has the right permissions
If ($Interactive) {
    [int]$RequiredScopesCount = $RequiredScopes.Count
    [string[]]$CurrentScopes = (Get-MgContext).Scopes
    [string[]]$RequiredScopes = $RequiredScopes

    $CheckScopes =[object[]][Linq.Enumerable]::Intersect($RequiredScopes,$CurrentScopes)
    If ($CheckScopes.Count -ne $RequiredScopesCount ) { 
        Write-Host ("To run this script, you need to connect to Microsoft Graph with the following scopes: {0}" -f $RequiredScopes) -ForegroundColor Red
        Break
    }
}

Write-Output "Outstanding Tasks Check starting up..."
# Find the set of users to process - we use a group to define the target set. Update this in your tenant to point to a group to use.
Try {
    $Group = Get-MgGroup -Filter "displayName eq 'Weekly Task Updates'" -ErrorAction Stop
} Catch {
    Write-Error "Failed to retrieve the group: $_"
    break
}

# A hash table to store details of users the script has processed so we don't need to keep calling Graph for details of the same user multiple times (e.g. when they are assigned multiple tasks)
$UsersHash = @{}

# Populate the hash table with the members of the group we are using to define the target set for the report.
Try {
    [array]$Members = Get-MgGroupMember -GroupId $Group.Id | Select-Object -ExpandProperty Id -ErrorAction Stop
    ForEach ($Member in $Members) {
        $MemberData = Get-MgUser -UserId $Member -Property Id, DisplayName, Mail
        $UsersHash.Add($Member, $MemberData)
    }
} Catch {
    Write-Error "Failed to retrieve group members: $_"
    break
}

$Report = [System.Collections.Generic.List[Object]]::new()
$GroupsHash = @{}

ForEach ($Member in $Members) {

    # Get the groups the user is a member of - we need this to determine which tasks to report on. Reduce the fetched set
    # to find just the Microsoft 365 groups, which is where plans are located
    [array]$Groups = Get-MgUserMemberOfAsGroup -UserId $Member -All -PageSize 500 -Property Id, DisplayName, GroupTypes `
        | Where-Object {$_.groupTypes -contains "unified"}
    
    Write-Output "Processing plans for user: $($UsersHash[$Member].DisplayName) ($($UsersHash[$Member].Mail)) - Groups: $($Groups.Count)"
    Write-Output "----------------------------------------------------------------------------------------------------------------------"

    ForEach ($Group in $Groups) {

        # Make sure that we don't process the plans for a group multiple times
        If ($GroupsHash.ContainsKey($Group.Id)) {
            Write-Output "  Already processed plans for group: $($Group.DisplayName) - skipping"
            continue
        } Else {
            Write-Output "  Processing plans for group: $($Group.DisplayName)"
            $GroupsHash[$Group.Id] = $Group.DisplayName
        }  

        Try {
            [array]$Plans = Get-MgGroupPlannerPlan -GroupId $Group.Id -ErrorAction Stop
        } Catch {
            Write-Error ("Failed to retrieve plans for group {0}: " -f $Group.DisplayName)
            continue
        }

        ForEach ($Plan in $Plans) {
            Write-Output "    Processing plan: $($Plan.Title)"
            Try {
                [array]$Tasks = Get-MgPlannerPlanTask -PlannerPlanId $Plan.Id -All -PageSize 500 `
                    -Property Id, Title, CompletedDateTime, StartDateTime, DueDateTime, BucketId, Assignments -ErrorAction Stop
                [array]$Buckets = Get-MgPlannerPlanBucket -PlannerPlanId $Plan.Id -All -PageSize 500 -ErrorAction Stop
            } Catch {
                Write-Error "Failed to retrieve tasks for plan $($Plan.Title)"
                continue
            }

            # Create hash table of buckets to translate bucket IDs to bucket names in the report
            $BucketsHash = @{}
            ForEach ($Bucket in $Buckets) {
                $BucketsHash.Add($Bucket.Id, $Bucket.Name)
            }
            # Filter the tasks to find those that are assigned to the user and are not completed
            $IncompleteTasks = $Tasks | Where-Object { $null -eq $_.CompletedDateTime }
            
            If ($IncompleteTasks.Count -gt 0) {
                Write-Output "      Found $($IncompleteTasks.Count) incomplete tasks in plan $($Plan.Title)"
                # Here you would add code to format the list of incomplete tasks and send it via email to the user
                # For example, you could create an HTML table of the tasks and use Send-MailMessage to send it
                ForEach ($Task in $IncompleteTasks) {
                    # Get the assignees
                    [array]$Assignees = @($Task.Assignments.AdditionalProperties.Keys)
                    If ($Assignees) {
                        ForEach ($Assignee in $Assignees) {
                            $Assignment = $Task.Assignments.AdditionalProperties[$Assignee]
                            $AssignedDateTime = $Assignment['assignedDateTime']
                            # Try to resolve the assignee identifier to a user using the user hash table
                            If ($UsersHash.ContainsKey($Assignee)) {
                                $AssignedUser = $UsersHash[$Assignee].DisplayName
                                $AssignedUserEmail = $UsersHash[$Assignee].Mail
                            } Else {
                                # Can't find the user, so create an entry in the hash table for them.
                                Try {
                                    $UserAssigned = Get-MgUser -UserId $Assignee -Property Id, DisplayName, Mail -ErrorAction Stop
                                    $AssignedUser = $UserAssigned.DisplayName
                                    $AssignedUserEmail = $UserAssigned.Mail
                                    $UsersHash.Add($Assignee, $UserAssigned)
                                } Catch {
                                    Write-Error ("Failed to retrieve user details for assignee {0} in task {1}:" -f $Assignee, $Task.Title)
                                    continue
                                }
                              
                            }
                            $ReportLine = [PSCustomObject]@{
                                User                = $AssignedUser
                                UserEmail           = $AssignedUserEmail
                                Group               = $Group.DisplayName
                                Plan                = $Plan.Title
                                Task                = $Task.Title
                                Bucket              = $BucketsHash[$Task.BucketId]
                                'Assigned Date'     = Get-Date $AssignedDateTime -format 'dd-MMM-yyyy HH:mm'
                                'Days Outstanding'  = (New-TimeSpan -Start $AssignedDateTime -End (Get-Date)).Days
                                'Start Date'        = If ($Task.StartDateTime) { Get-Date $Task.StartDateTime -format 'dd-MMM-yyyy' } Else { $null }
                                'Due Date'          = If ($Task.DueDateTime) { Get-Date $Task.DueDateTime -format 'dd-MMM-yyyy' } Else { $null }
                                Id                  = $Assignee
                            }
                            $Report.Add($ReportLine)
                        }
                    }   
                }
            } Else {
                Write-Output "      No incomplete tasks found in plan $($Plan.Title)"
            }
        }
    }   
}

$Report = $Report | Where-Object { $_.UserEmail -ne $null } | Sort-Object User, Group, Plan, 'Assigned Date'
# Get the unique set of users in the report so we can send them an email with their incomplete tasks
[array]$UniqueUsers = $Report | Select-Object -Property Id -Unique
# Reduce the set to the folks we are interested in (i.e. those in our target group)
$UniqueUsers = $UniqueUsers | Where-Object { $Members -contains $_.Id } | Select-Object -ExpandProperty Id

# Set up to use the ImportExcel module to generate Excel worksheets if it is available. If not, the report will be generated in CSV format instead.
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
}

[int]$UpdatesSent = 0
# Send mail to $usersHash[$member].mail
ForEach ($Id in $UniqueUsers) {
    $DestinationEmailAddress = $UsersHash[$Id].Mail
    $UserName = $UsersHash[$Id].DisplayName
    Write-Output "Sending email to $UserName ($DestinationEmailAddress) with their incomplete tasks"
    # Get list of incomplete tasks for the user
    [array]$UserTasks = $Report | Where-Object { $_.Id -eq $Id }

    # Generate the attachment in either Excel worksheet or CSV format, depending on if the ImportExcel module is available
    If ($ExcelGenerated) {
        If ($Interactive) {
            $AttachmentFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\TasksTemporary_$($Id).xlsx"
        } Else {
            $AttachmentFile = "TasksTemporary_$($Id).xlsx"
        }
            $UserTasks | Export-Excel -Path $AttachmentFile -WorksheetName "OutstandingTasks" -Title ("Outstanding Tasks {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "Tasks" 
    } Else {
        If ($Interactive) {
            $AttachmentFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\TasksTemporary_$($Id).csv"
        } Else {
            $AttachmentFile = "TasksTemporary_$($Id).csv"
        }
        $UserTasks | Export-Csv -Path $AttachmentFile -NoTypeInformation -Encoding Utf8
    }

    # prepare the attachment for sending in the email
    $EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($AttachmentFile))
    $MsgAttachments = @(
    @{
        '@odata.type' = '#microsoft.graph.fileAttachment'
        Name = (Split-Path $AttachmentFile -Leaf)
        ContentBytes = $EncodedAttachmentFile
        ContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
    )
    $AverageDaysOutstanding = [math]::Round(($UserTasks | Measure-Object -Property 'Days Outstanding' -Average).Average, 2)
    $LongestOutstandingTask = $UserTasks | Sort-Object -Property 'Days Outstanding' -Descending | Select-Object -First 1
    $MostTasksinBucket = $UserTasks | Group-Object -Property Bucket | Sort-Object -Property Count -Descending | Select-Object -First 1
    # Generate the HTML body for the email
    $HtmlHeading = "<h2 style='background-color:#CC0000;color:white;padding:10px;'>Incomplete Planner Tasks for $UserName</h2>"
    $HtmlContent = "<p>Dear $UserName</p>"
    $HtmlContent += "<p>Here is a list of your <b>$($UserTasks.count)</b> incomplete Planner tasks as of $(Get-Date -Format 'dd-MMM-yyyy').</p>"
    $HtmlContent += "<p>On average, these tasks have been incomplete for <b>$AverageDaysOutstanding</b> days, with the longest incomplete task being <b>$($LongestOutstandingTask.Task)</b> from the plan <b>$($LongestOutstandingTask.Plan)</b>, which has been incomplete for <b>$($LongestOutstandingTask.'Days Outstanding')</b> days.</p>"
    $HtmlContent += "<p>The bucket with the most tasks is <b>$($MostTasksinBucket.Name)</b> in the <b>$($MostTasksinBucket.Group.Plan[0])</b> plan with <b>$($MostTasksinBucket.Count)</b> tasks.</p>"
    $HtmlContent += "<p>Please try to get these tasks completed as soon as possible!</p>"

    $HtmlTable = $UserTasks | Select-Object Group, Plan, Task, Bucket, 'Assigned Date', 'Start Date', 'Due Date' | `
        Sort-Object {'Assigned Date' -as [datetime]} | ConvertTo-Html -Fragment
    $HtmlMessage = $HtmlHeading + $HtmlContent + $HtmlTable

    # Add message details
    $ToRecipient = @{}
    $ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
    [array]$MsgTo = $ToRecipient
    $MsgSubject = "Your incomplete Planner Tasks - $(Get-Date -Format 'dd-MMM-yyyy')"
    # Construct the message body 	
    $MsgBody = @{}
    $MsgBody.Add('Content', "$($HtmlMessage)")
    $MsgBody.Add('ContentType','html')
    # Build the message parameters
    $Message = @{}
    $Message.Add('subject', $MsgSubject)
    $Message.Add('toRecipients', $MsgTo)
    $Message.Add('body', $MsgBody)
    $Message.Add("attachments", $MsgAttachments)

    $EmailParameters = @{}
    $EmailParameters.Add('message', $Message)
    $EmailParameters.Add('saveToSentItems', $true)
    $EmailParameters.Add('isDeliveryReceiptRequested', $true)
    # Send the message
    Try {
        Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters -ErrorAction Stop
        Write-Output ("Incomplete Planner Tasks reminder emailed to {0}" -f $DestinationEmailAddress)
        $UpdatesSent++
    } Catch {
        Write-Output "Unable to send email"
        Write-Output $_.Exception.Message
    }
    # Clean up temporary CSV file
    Remove-Item -Path $AttachmentFile -ErrorAction SilentlyContinue
}

Write-Output ("All done! {0} email(s) sent to users with information about their incomplete tasks." -f $UpdatesSent)    

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.