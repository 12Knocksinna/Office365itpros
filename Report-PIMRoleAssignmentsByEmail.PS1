<# 
.SYNOPSIS
Report-PIMRoleAssignmentsByEmail.PS1
Searches the Entra ID audit log for PIM role assignment events for the last week and sends a report to specified email addresses.

.NOTES
Intended to run as a scheduled weekly task under Azure Automation

Requires Microsoft.Graph SDK and delegated scopes:
AuditLog.Read.All, RoleManagement.Read.Directory, Mail.Send

Version 1.0 14-Jan-2026
GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-PIMRoleAssignmentsByEmail.PS1

#>

[array]$RequiredScopes = "AuditLog.Read.All","RoleManagement.Read.Directory","Mail.Send"
$Interactive = $false

# Define the mailboxes that are used to send and receive the report. If run interactvely, the email goes from the signed-in account
$MsgFrom = 'Customer.Services@contoso.com'
$DestinationEmailAddress = "Admin.Info@contoso.com"

# Determine if we're interactive or not
If ([Environment]::UserInteractive) { 
    # We're running interactively...
    Clear-Host
    Write-Host "Script running interactively... connecting to the Graph" -ForegroundColor Yellow
    Connect-MgGraph -NoWelcome -Scopes $RequiredScopes
    $MsgFrom = (Get-MgContext).Account
    $Interactive = $true
} Else { 
    # We're not, so likely in Azure Automation
    Write-Output "Executing the runbook to create the last activity report for guest access outside the tenant..." 
    Connect-MgGraph -Identity -NoWelcome
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

# Retrieve role assignment events from the last week
$StartDate = (Get-Date).AddDays(-7).toString('yyyy-MM-dd')

Write-Output "Checking for role assignments in the last week since $StartDate..."
[array]$Records = Get-MgAuditLogDirectoryAudit -All -Filter `
     "(activityDisplayName eq 'Add member to role' or activityDisplayName eq 'Remove member from role') and ActivityDateTime gt $StartDate"

If ($Records.Count -eq 0) {
    Write-Host "No PIM role assignment events found in the last week." -ForegroundColor Yellow
    Break
} Else {
    Write-Output ("Found {0} PIM role assignment events in the last week." -f $Records.Count)
}

$CriticalRoles = @(
    "Global Administrator",
    "Privileged Role Administrator",
    "User Administrator",
    "Security Administrator",
    "Exchange Administrator",
    "SharePoint Administrator",
    "Teams Administrator"
)

Write-Output "Processing audit records to build report..."

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Record in $Records) {
    $CriticalFlag = $false
    If ($Record.ActivityDisplayName -eq 'Add member to role') {
        $AssignmentType = "Add role"
    } ElseIf ($Record.ActivityDisplayName -eq 'Remove member from role') {
        $AssignmentType = "Remove role"
    } Else {
        $AssignmentType = "Unknown"
    }
    If ($Record.initiatedBy.app.displayName -eq 'MS-PIM') {
        # Skip system initiated changes
        $AssignmentMethod = "Privileged Identity Management" 
        $ActionedBy = "MS-PIM"   
    }  Else {
        $ActionedBy =  $Record.initiatedBy.user.userPrincipalName
        $AssignmentMethod = "Direct Assignment"
    }
    $Role =  (Get-MgDirectoryRole -DirectoryRoleId $Record.TargetResources[1].Id).DisplayName

    If ($Role -in $CriticalRoles) {
        $CriticalFlag = $true
    }

    $TargetUser = Get-MgUser -UserId $Record.TargetResources[0].Id
    $ReportItem = [PSCustomObject]@{
        'Assigned Role'      = $Role
        'Assignee'           = $TargetUser.DisplayName
        'Assignee UPN'       = $TargetUser.UserPrincipalName
        'Actioned By'        = $ActionedBy
        'Assignment Method'  = $AssignmentMethod
        'Assignment Date'    = Get-Date $Record.ActivityDateTime -format 'dd-MMM-yyyy HH:mm'
        'Assignment Type'    = $AssignmentType
        'Is Critical Role'   = If ($CriticalFlag) { "🚨" } Else { $null }
    }
    $Report.Add($ReportItem)
}

# Build attachment for the email 
Write-Output "Generating output...  "
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $true
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\PIM RRole Assignments.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    } 
    $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "PIM Role Assignments" -Title ("PIM Role Assignments Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "PIMRoleAssignments" -AutoSize -AutoFilter
    $AttachmentFile = $ExcelOutputFile
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\PIM Role Assignments.CSV"
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
    $AttachmentFile = $CSVOutputFile
}

If ($ExcelGenerated) {
    Write-Output ("Excel worksheet output written to {0}" -f $ExcelOutputFile)
} Else {
    Write-Output ("CSV output file written to {0}" -f $CSVOutputFile)
} 

# Send the spreadsheet as an email attachment
$EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($AttachmentFile))
$MsgAttachments = @(
    @{
        '@odata.type' = '#microsoft.graph.fileAttachment'
        Name = (Split-Path $AttachmentFile -Leaf)
        ContentBytes = $EncodedAttachmentFile
        ContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
)

$HTMLAddRole = $Report | Where-Object {$_.'Assignment Type' -eq "Add role"} | Select-Object 'Assignment Date', 'Assigned Role','Assignee','Assignee UPN','Actioned By','Is Critical Role' | ConvertTo-Html -Fragment
$HTMLRemoveRole = $Report | Where-Object {$_.'Assignment Type' -eq "Remove role"} | Select-Object 'Assignment Date', 'Assigned Role','Assignee','Assignee UPN','Actioned By','Is Critical Role' | ConvertTo-Html -Fragment

# Build the array of a single TO recipient detailed in a hash table - change the sender and recipient to the appropriate recipient for your tenant
$ToRecipient = @{}
$ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
[array]$MsgTo = $ToRecipient
$MsgSUbject = "Weekly Role Assignments Report - $(Get-Date -Format 'dd-MMM-yyyy')"

$HtmlMsg = "</body></html><p>The output file for the <b>Weekly Role Assignments Report</b> is attached to this message. Please review the information at your convenience.</p>"
$HtmlMsg += "<h2>Weekly Role Assignment Report</h2>"
$HtmlMsg += "<p><h3>Role Assignments Made</h3></p>"
$HtmlMsg += $HTMLAddRole 
$HtmlMsg += "<p><h3>Role Assignments Removed</h3></p>"
$HtmlMsg += $HTMLRemoveRole
$htmlMsg += "</body></html>"

# Construct the message body 	
$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlMsg)")
$MsgBody.Add('ContentType','html')
# Build the pmessage parameters
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
    Write-Output ("Weekly role assignments report emailed to {0}" -f $DestinationEmailAddress)
    Write-Output "All done!"
} Catch {
    Write-Output "Unable to send email"
    Write-Output $_.Exception.Message
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.