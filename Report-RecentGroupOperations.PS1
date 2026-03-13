# Report-RecentGroypOperations.PS1
# A script to report recent additions, changes, and deletions of Microsoft 365 groups over the last 30 days. The script also generates
# a list of all the Microsoft 365 groups in the tenant, along with details of their membership and ownership. The report is generated in HTML format, and also in either Excel or CSV format depending on the availability of the ImportExcel module.

# V1.0 11-Mar-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-RecentGroupOperations.PS1

# Connect to Exchange Online PowerShell
[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Connecting to Exchange Online PowerShell..." -ForegroundColor Green
    Connect-ExchangeOnline -ShowBanner:$false
}   

# Connect to the Microsoft Graph with the necessary permissions to read group information. The script will prompt for consent to the required permissions if not already granted.
Connect-MgGraph -Scopes "GroupMember.Read.All", "Group.Read.All", "Organization.Read.All" -NoWelcome

# Search the audit log for group operations in the last 30 days
$StartDate = (Get-Date).AddDays(-30)
$EndDate = Get-Date

Write-Host "Finding Microsoft 365 groups in the tenant..." -ForegroundColor Green
# Find all Microsoft 365 groups in the tenant
[array]$Groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'unified')" -All -PageSize 500 `
     -ExpandProperty "Members(`$Select=displayName)" -Property Id, Members, displayName, Visibility, CreatedDateTime, ExpirationDateTime 

If ($Groups) {
    Write-Host ("Found {0} Microsoft 365 groups in the tenant." -f $Groups.Count) -ForegroundColor Green
    # Create a hash table of groups for quick lookup when processing the audit log records
    $GroupHash = @{}
    ForEach ($Group in $Groups) {
        $GroupHash.Add($Group.Id, $Group.displayName)
    }
} Else {
    Write-Host "No Microsoft 365 groups found in the tenant." -ForegroundColor Yellow
    Break
}

# Group administration operations include adding, updating, and deleting groups - https://learn.microsoft.com/purview/audit-log-activities?WT.mc_id=M365-MVP-9501#microsoft-entra-group-administration-activities
$Operations = "Add group.", "Update group.", "Delete group."
Write-Host "Searching the audit log for group operations from $($StartDate.ToShortDateString()) to $($EndDate.ToShortDateString())..." -ForegroundColor Green   
[array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations $Operations -Formatted -ResultSize 5000 -SessionCommand ReturnLargeSet

If ($Records) {
    # Remove duplicate records based on the identity field
    $Records = $Records | Sort-Object Identity -Unique
} Else {
    Write-Host "No group operations found in the audit log for the specified date range." -ForegroundColor Yellow
    Break
}

Write-Host "Analyzing audit log records..." -ForegroundColor Green
$AuditReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
    $AuditData = ConvertFrom-Json -InputObject $Rec.AuditData 
    $GroupId = $AuditData.Target.Id[1]
    # Only report if the group is a Microsoft 365 group - the audit log may include other group types such as security groups or distribution groups, but we are only interested in Microsoft 365 groups for this report
    If ($GroupHash.ContainsKey($GroupId)) {
        $ReportItem = [PSCustomObject]@{
            GroupName = $AuditData.Target.Id[3]
            GroupId = $GroupId
            Operation = $Rec.Operations
            User =  $Auditdata.Actor.id[0]
            Date = Get-Date $Rec.CreationDate -format 'dd/MM/yyyy HH:mm:ss'
        }
        $AuditReport.Add($ReportItem)
    } Else {
        Continue
    }
}

$AuditReport = $AuditReport | Sort-Object {$_.Date -as [datetime]} -Descending

Write-Host "Analyzing group details..." -ForegroundColor Green
$GroupReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Group in $Groups) {
    $MemberCount =  $Group.Members.additionalProperties.displayName.Count
    $MemberNames = $Group.Members.additionalProperties.displayName -join ", "
    $Owners = Get-MgGroupOwner -GroupId $Group.Id -All
    $OwnerCount = $Owners.Count
    $OwnerNames = $Owners.additionalProperties.displayName -join ", "
    $ReportItem = [PSCustomObject]@{
        Group               = $Group.displayName
        GroupId             = $Group.Id
        "Member Count"      = $MemberCount
        "Member Names"      = $MemberNames
        "Owner Count"       = $OwnerCount
        "Owner Names"       = $OwnerNames
        Visibility          = $Group.Visibility
        Created             = $Group.CreatedDateTime
        "Expiration Date"   = $Group.ExpirationDateTime
    }
    $GroupReport.Add($ReportItem)
}

$GroupReport = $GroupReport | Sort-Object Group

If ($AuditReport) {
    Write-Host ("Found {0} unique group operations in the audit log since {1}." -f $Records.Count, $StartDate.ToShortDateString()   ) -ForegroundColor Green
    [array]$GroupAdditions = $AuditReport | Where-Object { $_.Operation -eq "Add group." }
    [array]$GroupUpdates = $AuditReport | Where-Object { $_.Operation -eq "Update group." }
    [array]$GroupRemovals = $AuditReport | Where-Object { $_.Operation -eq "Delete group." }
    Write-Host ("Group Additions: {0}" -f $GroupAdditions.Count) -ForegroundColor Cyan
    Write-Host ("Group Updates: {0}" -f $GroupUpdates.Count) -ForegroundColor Cyan
    Write-Host ("Group Removals: {0}" -f $GroupRemovals.Count) -ForegroundColor Cyan  
} Else {
    Write-Host "No recent group operations found in the audit log." -ForegroundColor Yellow 
}

Write-Host "Generating HTML report..." -ForegroundColor Green
$CreationDate = Get-Date -format 'dd-MMM-yyyy HH:mm:ss'
$Version = "1.0"
$Organization = (Get-MgOrganization).DisplayName

# Create the HTML report
$htmlhead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 36px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 24px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Microsoft 365 Groups Report</h1></p>
           <p><h2><b>For the " + $Organization + " organization</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format g) + "</h3></p></div>"

$HTMLAnalysisSectionHead = "<p><h2>Analysis</h2></p>" +
                   "<p>The report shows all Microsoft 365 groups in the tenant, along with details of their membership and ownership. The report also includes a list of recent group operations in the tenant, including group additions, updates, and removals.</p>" +
                   "<p>Use this report to understand the current state of Microsoft 365 groups in the tenant, and to identify any recent changes to groups that may require attention.</p>"
$HTMLAnalysisAdditions = "<p><h3>Group Additions</h3></p>" +
                 "<p>The following groups have been added in the last 30 days:</p>" +
                 ($AuditReport | Where-Object { $_.Operation -eq "Add group." }| Select-Object -Property GroupName, User, Date | `
                 ConvertTo-Html -Fragment)
$HTMLAnalysisUpdates = "<p><h3>Group Updates</h3></p>" +
                 "<p>The following groups have been updated in the last 30 days:</p>" +
                 ($AuditReport | Where-Object { $_.Operation -eq "Update group." }| Select-Object -Property GroupName, User, Date | ConvertTo-Html -Fragment)
$HTMLAnalysisRemovals = "<p><h3>Group Removals</h3></p>" +
                 "<p>The following groups have been removed in the last 30 days:</p>" +
                 ($AuditReport | Where-Object { $_.Operation -eq "Delete group." }| Select-Object -Property GroupName, User, Date | ConvertTo-Html -Fragment)

$HTMLBody = $HTMLHead + $HTMLAnalysisSectionHead
If ($GroupAdditions.Count -gt 0) {
    $HTMLBody += $HTMLAnalysisAdditions
} Else {
    $HTMLBody += "<p><h3>Microsoft 365 Group Additions</h3></p><p>No additions of Microsoft 365 Groups in the last 30 days.</p>"
}
If ($GroupUpdates.Count -gt 0) {
    $HTMLBody += $HTMLAnalysisUpdates
} Else {
    $HTMLBody += "<p><h3>Microsoft 365 Group Updates</h3></p><p>No updates for Microsoft 365 Groups found for the last 30 days.</p>"
}
If ($GroupRemovals.Count -gt 0) {
    $HTMLBody += $HTMLAnalysisRemovals
} Else {
    $HTMLBody += "<p><h3>Microsoft 365 Group Removals</h3></p><p>No removals of Microsoft 365 Groups found for the last 30 days.</p>"
}

$GroupDetailsHeader = "<p><h2>Microsoft 365 Groups Details</h2></p>" +
                      "<p>The following table shows details of all Microsoft 365 groups in the tenant, including their membership and ownership information.</p>"
$GroupDetailsTable = $GroupReport | Select-Object Group, GroupId, "Member Count", "Owner Count", Visibility, Created, "Expiration Date" | ConvertTo-Html -Fragment
$HTMLBody = $HTMLBody + $GroupDetailsHeader + $GroupDetailsTable    

$HTMLTail = "<p><h3>Report created for: " + $Organization + "</h3></p>" +
             "<p>Created: " + $CreationDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of Microsoft 365 Groups found:                     " + $GroupReport.Count + "</p>" +
             "<p>Number of Microsoft 365 Groups without a manager:         " + ($GroupReport | Where-Object { $_."Owner Count" -eq 0 }).Count + "</p>" +
             "<p>Number of Microsoft 365 Groups with no members:           " + ($GroupReport | Where-Object { $_."Member Count" -eq 0 }).Count + "</p>" +
             "<p>Number of recent Microsoft 365 Groups audit events found: " + $AuditReport.Count + "</p>" +
             "<p>Number of recent Microsoft 365 Group additions:           " + $GroupAdditions.Count + "</p>" +
             "<p>Number of recent Microsoft 365 Group updates:             " + $GroupUpdates.Count + "</p>" +
             "<p>Number of recent Microsoft 365 Group removals:            " + $GroupRemovals.Count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Microsoft 365 Groups Report <b>" + $Version + "</b>"	

$HTMLReport = $HTMLBody + $HTMLTail
$HTMLReportFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Microsoft 365 Groups.html"
$HTMLReport | Out-File $HTMLReportFile -Encoding UTF8

# Generate the report in either Excel worksheet or CSV format, depending on if the ImportExcel module is available
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Microsoft 365 Groups.xlsx"
    $GroupReport | Export-Excel -Path $ExcelOutputFile -WorksheetName "Microsoft 365 Groups" -Title ("Microsoft 365 Groups {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "M365Groups" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Microsoft 365 Groups.CSV"
    $GroupReport | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}   

If ($ExcelGenerated -eq $true) {
    Write-Host ("Microsoft 365 groups report is available in Excel workbook {0}" -f $ExcelOutputFile)
    $AttachmentFile = $ExcelOutputFile
} Else {
    Write-Host ("Microsoft 365 groups report is available in CSV file {0}" -f $CSVOutputFile)
    $AttachmentFile = $CSVOutputFile
}

Write-Host ("HTML report of Microsoft 365 groups is available at {0}" -f $HTMLReportFile) -ForegroundColor Green

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

# The message is sent from the signed in user.
$MsgFrom = (Get-MgContext).Account
# The message is sent to the specified recipient - change the sender and recipient to the appropriate recipient for your tenant
$DestinationEmailAddress = "Customer.Services@Office365itpros.com"
# Build the array of a single TO recipient detailed in a hash table - change the sender and recipient to the appropriate recipient for your tenant
$ToRecipient = @{}
$ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
[array]$MsgTo = $ToRecipient
$MsgSubject = "Microsoft 365 Groups Report - $(Get-Date -Format 'dd-MMM-yyyy')"

# Construct the message body 	
$MsgBody = @{}
$MsgBody.Add('Content', "$($HTMLReport)")
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
    Write-Output ("Microsoft 365 groups report emailed to {0}" -f $DestinationEmailAddress)
    Write-Output "All done!"
} Catch {
    Write-Output "Unable to send email"
    Write-Output $_.Exception.Message
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.