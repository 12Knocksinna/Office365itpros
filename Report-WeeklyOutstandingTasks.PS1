# Report-WeeklyOutstandingTasks.PS1
# This script generates a weekly report of outstanding tasks for a user and sends the list to them via email

# V1.0 1-Apr-2026
# GitHub link: 

# Permissions needed to run the script:
# - User.Read.All (to read user details)
# - Group.Read.All (to read group details and membership)
# - GroupMember.Read.All (to read group membership)
# - Tasks.Read.All (to read the tasks assigned to users)

# The script is intended to run as a scheduled Azure Automation runbook. Make sure that the Azure automation account has the
# necessary permissions to run the script and access the Microsoft Graph API.

# Modules needed for runtime environment:
# - Microsoft.Graph.Authentication (to connect to Microsoft Graph)
# - Microsoft.Graph.Users (to read user details)
# - Microsoft.Graph.Groups (to read group details and membership)   
# - Microsoft.Graph.Planner (to read the tasks assigned to users)
# - Microsoft.Graph.Users.Actions (to send the email with the report)
# - ImportExcel (optional, to generate Excel worksheets for the report attachments)


Connect-MgGraph -Identity 

# Find the set of users to process - we use a group to define the target set
Try {
    $Group = Get-MgGroup -Filter "displayName eq 'Ultimate Guide to Office 365'" -ErrorAction Stop
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
    # to find just the Mkicrosoft 365 groups, which is where plans are located
    [array]$Groups = Get-MgUserMemberOfAsGroup -UserId $Member -All -PageSize 500 -Property Id, DisplayName, GroupTypes `
        | Where-Object {$_.groupTypes -contains "unified"}
    
    Write-Output "Processing plans for user: $($UsersHash[$Member].DisplayName) ($($UsersHash[$Member].Mail)) - Groups: $($Groups.Count)"

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
                [array]$Tasks = Get-MgPlannerPlanTask -PlannerPlanId $Plan.Id -All -PageSize 500 -ErrorAction Stop
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
            $OutstandingTasks = $Tasks | Where-Object { $_.PercentComplete -lt 100 }
            
            If ($OutstandingTasks.Count -gt 0) {
                Write-Output "      Found $($OutstandingTasks.Count) outstanding tasks in plan $($Plan.Title)"
                # Here you would add code to format the list of outstanding tasks and send it via email to the user
                # For example, you could create an HTML table of the tasks and use Send-MailMessage to send it
                ForEach ($Task in $OutstandingTasks) {
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
                                AssignedDateTime    = Get-Date $AssignedDateTime -format 'dd-MMM-yyyy HH:mm'
                                StartDate           = If ($Task.StartDateTime) { Get-Date $Task.StartDateTime -format 'dd-MMM-yyyy' } Else { $null }
                                DueDate             = If ($Task.DueDateTime) { Get-Date $Task.DueDateTime -format 'dd-MMM-yyyy' } Else { $null }
                                Id                  = $Assignee
                            }
                            $Report.Add($ReportLine)
                        }
                    }   
                }
            } Else {
                Write-Output "      No outstanding tasks found in plan $($Plan.Title)"
            }
        }
    }   
}

$Report = $Report | Where-Object { $_.UserEmail -ne $null } | Sort-Object User, Group, Plan, AssignedDateTime
# Get the unique set of users in the report so we can send them an email with their outstanding tasks
[array]$UniqueUsers = $Report | Select-Object -Property Id -Unique
# Reduce the set to the folks we are interested in (i.e. those in our target group)
$UniqueUsers = $UniqueUsers | Where-Object { $Members -contains $_.Id }

# The message is sent from the signed in user.
$MsgFrom = (Get-MgContext).Account

# Set up to use the ImportExcel module to generate Excel worksheets if it is available. If not, the report will be generated in CSV format instead.
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
}

# Send mail to $usersHash[$member].mail
ForEach ($Id in $UniqueUsers) {
    $DestinationEmailAddress = $UsersHash[$Id].Mail
    $UserName = $UsersHash[$Id].DisplayName
    Write-Output "Sending email to $UserName ($DestinationEmailAddress) with their outstanding tasks"
    [array]$UserTasks = $Report | Where-Object { $_.Id -eq $Id }

    # Generate the attachment in either Excel worksheet or CSV format, depending on if the ImportExcel module is available
    If ($ExcelGenerated) {
        $AttachmentFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\TasksTemporary_$($UserName).xlsx"
        $UserTasks | Export-Excel -Path $AttachmentFile -WorksheetName "OutstandingTasks" -Title ("Outstanding Tasks {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "Tasks" 
    } Else {
        $AttachmentFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\TasksTemporary_$($UserName).CSV"
        $UserTasks | Export-Csv -Path $AttachmentFile -NoTypeInformation -Encoding Utf8
    }

    $EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($AttachmentFile))
    $MsgAttachments = @(
    @{
        '@odata.type' = '#microsoft.graph.fileAttachment'
        Name = (Split-Path $AttachmentFile -Leaf)
        ContentBytes = $EncodedAttachmentFile
        ContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
    )
    $HtmlHeading = "<h2>Outstanding Planner Tasks for $UserName</h2>"
    $HtmlContent = "<p>Dear $UserName,</p><p>Here is a list of your <b>$($UserTasks.count)</b> outstanding Planner tasks as of $(Get-Date -Format 'dd-MMM-yyyy').</p>"
    $HtmlContent += "<p>Keep up the good work and try to get these tasks completed as soon as possible!</p>"
    $HtmlTable = $UserTasks | Select-Object Group, Plan, Task, Bucket, AssignedDateTime, StartDate, DueDate | `
        Sort-Object {$_.AssignedDateTime -as [datetime]} | ConvertTo-Html -Fragment
    $HtmlMessage = $HtmlHeading + $HtmlContent + $HtmlTable
    $ToRecipient = @{}
    $ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
    [array]$MsgTo = $ToRecipient
    $MsgSubject = "Your outstanding Planner Tasks - $(Get-Date -Format 'dd-MMM-yyyy')"
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
        Write-Output ("Outstanding Planner Tasks reminder emailed to {0}" -f $DestinationEmailAddress)
        Write-Output "All done!"
    } Catch {
        Write-Output "Unable to send email"
        Write-Output $_.Exception.Message
    }
    # Clean up temporary CSV file
    Remove-Item -Path $AttachmentFile -ErrorAction SilentlyContinue
}




# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.