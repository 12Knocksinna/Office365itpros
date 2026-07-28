# Remove-OldPlannerTasks.PS1
# This script removes old completed tasks from Microsoft Planner based on a specified age threshold.

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Remove-OldPlannerTasks.PS1
# V1.0 25-Jul-2025

# This script must be run in app-only mode to have full access to all Microsoft 365 groups and plans in the tenant.
# Alternatively, run it as an Azure Automation runbook with the appropriate permissions.

# GroupMember.Read.all - Read group information 
# Tasks.ReadWrite.All - Read and write tasks in Microsoft Planner
# Mail.Send - send mail with report

Connect-MgGraph -NoWelcome

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
                Remove-MgPlannerTask -PlannerTaskId $Task.Id -IfMatch $Task.additionalProperties.'@odata.etag' -ErrorAction Stop
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
$UniquePlans = $Report | Select-Object -ExpandProperty PlanId -Unique
$Report2 = [System.Collections.Generic.List[object]]::new()
ForEach ($PlanId in $UniquePlans) {
    $PlanReport = $Report | Where-Object { $_.PlanId -eq $PlanId }
    $Report2.Add([PSCustomObject]@{
        'Group Name'    = $PlanReport[0].GroupName
        'Plan title'    = $PlanReport[0].PlanTitle
        'Tasks Removed' = $PlanReport.Count
    })
}

Write-Host ""
Write-Host "Summary of removed tasks by Plan:"
$Report2 | Sort-Object 'Group Name', 'Plan title' | Format-Table -AutoSize
$Report2 = $Report2 | Sort-Object 'Tasks Removed' -Descending

Write-Host ""
Write-Host "Task removal process completed. $($Report.Count) tasks removed using a threshold date of $($ThresholdDate.ToString('dd-MMM-yyyy'))."
Write-Host "Number of plans scanned: $PlansScanned" 
Write-Host "Number of groups scanned: $($Groups.count)"
Write-Host ("Tasks removed from unique plans: {0}" -f $UniquePlans.Count)
Write-Host "Duration: $($Duration.ToString())"

# Generate report file
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $OutputXLSXFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Deleted Planner Tasks.xlsx"
    Remove-Item -Path $OutputXLSXFile -ErrorAction SilentlyContinue
    $Report | Export-Excel -Path $OutputXLSXFile -WorksheetName "Deleted Planner Task" -Title ("Deleted Planner Tasks {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "RemovedPlannerTasks" 
} Else {
    $OutputCSVFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Deleted Planner Tasks.csv"
    Remove-Item -Path $OutputCSVFile -ErrorAction SilentlyContinue
    $Report | Export-Csv -Path $OutputCSVFile -NoTypeInformation -Encoding Utf8
}
  
If ($ExcelGenerated) {
  Write-Host ("An Excel workbook containing details of removed Planner tasks is available in {0}" -f $OutputXLSXFile)
  $AttachmentFile = $OutputXLSXFile
} Else {
  Write-Host ("A CSV file containing details of removed Planner tasks is available in {0}" -f $OutputCSVFile)
  $AttachmentFile = $OutputCSVFile
}

$EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($AttachmentFile))
$MsgAttachments = @(
    @{
	"@odata.type" = "#microsoft.graph.fileAttachment"
	Name = ($AttachmentFile -split '\\')[-1]
	ContentType = if ($AttachmentFile -like "*.csv") {"application/csv"} else {"application/vnd.ms-excel"}  
	ContentBytes = $EncodedAttachmentFile
	}
)

# Send email with the results - define your own email address here and the address to receive the report.
$MsgFrom = "PowerShell.Reports@office365itpros.com"
$DestinationEmailAddress = "User.Administration@office365itpros.com"
# Build the array of a single TO recipient detailed in a hash table - change this to the appropriate recipient for your tenant
$ToRecipient = @{}
$ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
[array]$MsgTo = $ToRecipient
# Define the message subject
$MsgSubject = "Deleted Planner Task Report - $($Report.Count) tasks removed from $($UniquePlans.Count) plans"
# Create the HTML content
$HtmlMsg = ("</body></html><p>The attached report lists Planner tasks that were removed based on the threshold date of $($ThresholdDate.ToString('dd-MMM-yyyy')).</p>"  )
$HtmlMsg += "<p>Please review the information at your convenience</p>"
# Construct the message body 	
$HtmlBody = $Report2 | ConvertTo-Html -As Table -PostContent "</body></html>" -Fragment
$HtmlMsg = $HtmlMsg + $HtmlBody
$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlMsg)")
$MsgBody.Add('ContentType','html')
# Build the parameters to submit the message
$Message = @{}
$Message.Add('subject', $MsgSubject)
$Message.Add('toRecipients', $MsgTo)
$Message.Add('body', $MsgBody)
$Message.Add("attachments", $MsgAttachments)

$EmailParameters = @{}
$EmailParameters.Add('message', $Message)
$EmailParameters.Add('saveToSentItems', $true)
$EmailParameters.Add('isDeliveryReceiptRequested', $true)
Write-Output ""
Write-Output "Sending email..."
Write-Output ""

# Send the message
Try {
    Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters -ErrorAction Stop
    Write-Output ("Planner Task deletion report emailed to {0}" -f $ToRecipient.emailAddress.address)
} Catch {
    Write-Output "Unable to send email"
    Write-Output $_.Exception.Message
}

Write-Host "Planner Task Cleanup complete!"

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.