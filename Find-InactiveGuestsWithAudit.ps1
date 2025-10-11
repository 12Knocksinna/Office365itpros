# Find-InactiveGuestsWithAudit.PS1
# A script to find inactive Entra ID guests and report what they've been doing

# V1.0.1 6-Oct-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-InactiveGuestsWithAudit.ps1

# Flag to let script code know if we're running interactively or within Azure Automation
$Interactive = $false

# Determine if we're interactive or not
If ([Environment]::UserInteractive) { 
    # We're running interactively...
    Write-Host "Script running interactively... connecting to the Graph" -ForegroundColor Yellow
    Connect-MgGraph -NoWelcome -Scopes User.Read.All, AuditLog.Read.All, Mail.Send, Organization.Read.All
    $Interactive = $true
    [array]$Modules = Get-Module | Select-Object -ExpandProperty Name
    If ("ExchangeOnlineManagement" -Notin $Modules) {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowBanner:$false 
    }
    $MsgFrom = (Get-MgContext).Account
} Else { 
    # We're not, so likely in Azure Automation
    Write-Host "Running the script to identify the last app accessed by Users" 
    Connect-MgGraph -Identity -NoWelcome
    $Tenant = Get-MgOrganization
    # Connect with a managed identity
    $TenantDomain = $Tenant.VerifiedDomains | Where-Object {$_.isDefault -eq $true} | Select-Object -ExpandProperty Name
    Connect-ExchangeOnline -ManagedIdentity -Organization $TenantDomain
    $CurrentFolder = (Get-Location).Path
    $MsgFrom = "no-reply@office365itpros.com"
}

# Check that we have the right permissions - in Azure Automation, we assume that the automation account has the right permissions
If ($Interactive) {
    [string[]]$CurrentScopes = (Get-MgContext).Scopes
    [string[]]$RequiredScopes = @('AuditLog.Read.All','User.Read.All','Mail.Send', 'Organization.Read.All')

    $CheckScopes =[object[]][Linq.Enumerable]::Intersect($RequiredScopes,$CurrentScopes)
    If ($CheckScopes.Count -ne 4) { 
        Write-Host ("To run this script, you need to connect to Microsoft Graph with the following scopes: {0}" -f $RequiredScopes) -ForegroundColor Red
        Disconnect-Graph
        Break
    }
}

# Change this to the email address of the recipient of the report
$DestinationEmailAddress = "customer.services@office365itpros.com"

# Find information about sharing events so that we know  when someone has been invited to the tenant or otherwise updated (like being added to a group)
$ShareDateStart = (Get-Date).AddDays(-365)
$EndDate = (Get-Date).AddDays(1)
# SharePoint Sharing Invitation
Write-Output "Searching for SharePoint sharing invitations..."
[array]$SharingRecords = Search-UnifiedAuditLog -StartDate $ShareDateStart -EndDate $EndDate -SessionCommand ReturnLargeSet -Formatted -ResultSize 5000 -Operations SharingInvitationCreated
$SharingRecords = $SharingRecords | Sort-Object Identity -Unique
$SharingData = [System.Collections.Generic.List[Object]]::new()
ForEach ($Record in $SharingRecords) {
    $AuditData = $Record.AuditData | ConvertFrom-Json
    If ($AuditData.TargetUserOrGroupType -eq "Guest") {
        $Guest = $AuditData.TargetUserOrGroupName
        $InvitationSource = $AuditData.UserId
        $Item = [PSCustomObject]@{
            Guest               = $Guest.toLower()
            InvitationSource    = $InvitationSource
            Timestamp           = $Record.CreationDate
            App                 = 'SharePoint Online'
        }
        $SharingData.Add($Item)
    }
}
# Guests added to Microsoft 365 Groups
Write-Output "Searching for Microsoft 365 Groups invitations..."
[array]$GroupData =  Search-UnifiedAuditLog -StartDate $ShareDateStart -EndDate $EndDate -SessionCommand ReturnLargeSet -Formatted `
    -ResultSize 5000 -Operations 'Add member to group.'
$GroupData = $GroupData | Sort-Object Identity -Unique
ForEach ($Record in $GroupData) {
    $AuditData = $Record.AuditData | ConvertFrom-Json
    If ($AuditData.ObjectId -Like "*#EXT#*" -and $AuditData.UserId -notlike "*ServicePrincipal*") {
        $Guest = $AuditData.ObjectId
        $InvitationSource = $AuditData.UserId
        $Item = [PSCustomObject]@{
            Guest               = $Guest.toLower()  
            InvitationSource    = $InvitationSource
            Timestamp           = $Record.CreationDate
            App                 = 'Groups'
        }
        $SharingData.Add($Item)
    }
}
# Users invited from the Entra ID portal
Write-Output "Searching for Entra ID invitations..."
[array]$InvitationData = Search-UnifiedAuditLog -StartDate $ShareDateStart -EndDate $EndDate -SessionCommand ReturnLargeSet -Formatted `
    -ResultSize 5000 -Operations 'Add user.'
$InvitationData = $InvitationData | Sort-Object Identity -Unique
ForEach ($Record in $InvitationData) {
    $AuditData = $Record.AuditData | ConvertFrom-Json
    If ($AuditData.ObjectId -Like "*#EXT#*" -and $AuditData.UserId -notlike "*ServicePrincipal*") {
        $Guest = $AuditData.ObjectId
        $InvitationSource = $AuditData.UserId
        $Item = [PSCustomObject]@{
            Guest               = $Guest.toLower()
            InvitationSource    = $InvitationSource
            Timestamp           = $Record.CreationDate
            App                 = 'Entra ID'
        }
        $SharingData.Add($Item)
    }
}

$SharingData  = $SharingData | Sort-Object {$_.TimeStamp -as [datetime]} -Descending

$StartDate = (Get-Date).AddDays(-30)
[datetime]$StartProcessing = Get-Date

# Define the audit records used to figure out the important events to indicate what guests have been doing
[array]$Operations = "FileAccessed","FileModified","FileUploaded","FileDeleted","FileDownloaded","MessageSent", "ReactedToMessage",`
     "MessageRead", "MessageDeleted", "TaskCompleted", "TaskRead", "TaskAssigned", "SensitivityLabeledFileOpened", "TeamsSessionStarted", "UserLoggedIn", "SignInEvent"

# Find all guests - a complex query is used to sort the retrieved results
Write-Output "Retrieving guest accounts..."
[array]$Guests = Get-MgUser -Filter "usertype eq 'Guest'" -PageSize 500 -All  `
    -Property DisplayName,UserPrincipalName,SignInActivity,Mail,Sponsors,Id,CreatedDateTime,AccountEnabled,EmployeeLeaveDateTime -ExpandProperty Sponsors `
    | Sort-Object displayName
If ($Guests.Count -eq 0) {
    Write-Host "No guest users found."
    break
} Else {
    Write-Host ("Found {0} guest users." -f $Guests.Count)
}

[int]$i = 0
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Guest in $Guests) {
    $i++
    $DaysSinceLastSignIn = $null; $DaysSinceLastSuccessfulSignIn = $null
    $GuestStatus = "Inactive"
    $TopActivities = $null
    Write-Host "Processing guest user $($Guest.DisplayName) <$($Guest.Mail)> ($($i)/$($Guests.Count))" -ForegroundColor Cyan
    [array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -UserIds $Guest.UserPrincipalName `
        -SessionCommand ReturnLargeSet -Formatted -ResultSize 5000 -Operations $Operations
    If ($Records.Count -eq 0) {
        Write-Host "No audit records found for guest user" -ForegroundColor Red
        $LastActivity = $null
    } Else {
        # Remove duplicate records
        $Records = $Records | Sort-Object Identity -Unique
        Write-Host ("Found {0} audit records for guest user" -f $Records.Count) -ForegroundColor Yellow
        $LastActivity = $Records | Sort-Object {$_.CreationDate -as [datetime]} -Descending | Select-Object -First 1
        [array]$GuestActivity = $Records | Group-Object Operations -NoElement | Sort-Object Count -Desc | Select-Object -First 3
        $TopActivities = ($GuestActivity | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ", "
        $GuestStatus = "Active"
    }
     
    # Can we find when the guest was invited to the tenant?
    [array]$Invitation = $SharingData | Where-Object { $_.Guest -eq $Guest.UserPrincipalName.tolower() } | `
        Sort-Object {$_.TimeStamp -as [datetime]} -Descending | Select-Object -First 1
    If ($Invitation) {
        $InvitedTimeStamp = Get-Date $Invitation.Timestamp -Format 'dd-MMMM-yyyy HH:mm'
        $InvitedSource = $Invitation.InvitationSource
    } Else {
        $InvitedTimeStamp = $null
        $InvitedSource = $null
    }

    If (!([string]::IsNullOrWhiteSpace($Guest.signInActivity.lastSuccessfulSignInDateTime))) {
        [datetime]$LastSuccessfulSignIn = $Guest.signInActivity.lastSuccessfulSignInDateTime
        $DaysSinceLastSuccessfulSignIn = (New-TimeSpan $LastSuccessfulSignIn).Days 
    }
    If (!([string]::IsNullOrWhiteSpace($Guest.signInActivity.lastSignInDateTime))) {
        [datetime]$LastSignIn = $Guest.signInActivity.lastSignInDateTime
        $DaysSinceLastSignIn = (New-TimeSpan $LastSignIn).Days
    }

    # Is there a photo for the guest?
    $Status = Get-MgUserPhoto -UserId $Guest.Id -ErrorAction SilentlyContinue
    If ($Status) {
        $HasPhoto = $true
    } Else {
        $HasPhoto = $false
    }
    # Is the guest a member of any groups?
    $MemberOf = Get-MgUserMemberOf -UserId $Guest.Id | Where-Object {$_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group'} -ErrorAction SilentlyContinue
    If ($MemberOf) {
        $GroupsCount = $MemberOf.Count
        $GroupsNames = $MemberOf.additionalProperties.displayName -join "; "
    } Else {
        $GroupsCount = 0
        $GroupsNames = $null
    }
    # Is the guest account disabled or is the employee leave date time property populated
    If ($Guest.AccountEnabled -eq $false -or $null -ne $Guest.EmployeeLeaveDateTime) {
        $GuestStatus = "Inactive"
    }

    $ReportItem = [PSCustomObject]@{
        Guest                           = $Guest.DisplayName
        UserPrincipalName               = $Guest.UserPrincipalName
        Email                           = $Guest.Mail
        Sponsors                        = ($Guest.Sponsors | ForEach-Object { Get-MgUser -UserId $_.Id | Select-Object -ExpandProperty DisplayName }) -join "; "
        'Creation Date'                 = Get-Date $Guest.CreatedDateTime -format 'dd-MMMM-yyyy HH:mm'
        'Days since creation'           = (New-TimeSpan $Guest.CreatedDateTime).Days
        'Date Last Audit Activity'      = If ($LastActivity) { Get-Date $LastActivity.CreationDate -format 'dd-MMMM-yyyy HH:mm'} else { $null }
        'Last Audit Activity'           = If ($LastActivity) { $LastActivity.Operations } else { $null }
        'Number of Audit Activities'    = $Records.Count
        'Top 3 activities'              = $TopActivities
        'Last administrator action'     = $InvitedTimeStamp
        'Administrator'                 = $InvitedSource
        'Last Signin'                   = Get-Date ($LastSignIn) -format 'dd-MMMM-yyyy HH:mm'
        'Days since last signin'        = $DaysSinceLastSignIn
        'Date of last successful signin'= Get-Date ($LastSuccessfulSignIn) -format 'dd-MMMM-yyyy HH:mm'
        'Days since last successful signin' = $DaysSinceLastSuccessfulSignin
        EmailDomain                     = ($Guest.Mail -split "@")[1]
        HasPhoto                        = $HasPhoto
        '# of groups guest is member of'= $GroupsCount
        'Groups guest is member of'     = $GroupsNames
        Id                              = $Guest.Id
        AccountEnabled                  = $Guest.AccountEnabled
        'Employee Leave Date'           = If ($Guest.EmployeeLeaveDateTime) { Get-Date $Guest.EmployeeLeaveDateTime -format 'dd-MMMM-yyyy HH:mm' } else { $null }
        'Guest status'                  = $GuestStatus
    }
    $Report.Add($ReportItem)

}

$Report = $Report | Sort-Object Guest

[datetime]$EndProcessing = Get-Date
$TimeRequired = $EndProcessing - $StartProcessing
$Minutes = [math]::Floor($TimeRequired.TotalSeconds / 60)
$Seconds = [math]::Round($TimeRequired.TotalSeconds % 60, 2)
If ($Interactive) {
    Write-Host ("Total processing time for {0} accounts: {1}m {2}s" -f $Guests.count, $Minutes, $Seconds) -ForegroundColor Cyan
    Write-Host ("Average required per user {0} seconds" -f [math]::Round($TimeRequired.TotalSeconds / $Guests.count, 2)) -ForegroundColor Cyan
} Else {
    Write-Output ("Total processing time for {0} accounts: {1}m {2}s" -f $Guests.count, $Minutes, $Seconds) 
    Write-Output ("Average required per user {0} seconds" -f [math]::Round($TimeRequired.TotalSeconds / $Guests.count, 2)) 
}

[array]$InactiveGuests = $Report | Where-Object { $_.'Guest status' -eq "Inactive" } | Sort-Object DisplayName
Write-Host ""
Write-Host ("Found {0} inactive guests ({1})" -f $InactiveGuests.Count,($InactiveGuests.Count/$Report.Count).toString("P")) -ForegroundColor Green

Write-Host ""
Write-Host "Inactive guests come from these domains"
$InactiveGuests | Group-Object EmailDomain | Sort-Object Count -Descending | Format-Table Name,Count -AutoSize

# Create a nice HTML report
# Generate sortable HTML table with type-aware sorting - use number as the type for numeric values, date for dates, and string for text
$HtmlHeader = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Inactive Guest Accounts Report</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; background: #f4f6f8; color: #222; }
h1 { background: #0078d4; color: #fff; padding: 16px; border-radius: 6px 6px 0 0; margin-bottom: 20px; }
table { width: 100%; background: #fff; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-collapse: collapse; }
th, td { padding: 12px; text-align: left; }
th { background: #e5eaf1; cursor: pointer; position: relative; }
th:hover { background: #d0e7fa; }
th::after { content: '↕'; position: absolute; right: 8px; opacity: 0.5; }
tr:nth-child(even) { background: #f0f4fa; }
tr:hover { background: #d0e7fa; }
</style>
<script>
function parseValue(val, type) {
    if(type === 'number') return parseFloat(val.replace(/,/g,'')) || 0;
    if(type === 'date') return new Date(val);
    return val.toLowerCase();
}
function sortTable(n, type) {
    var table = document.getElementById('GuestStats');
    var rows = Array.from(table.rows).slice(1);
    var dir = table.getAttribute('data-sortdir'+n) === 'asc' ? 'desc' : 'asc';
    rows.sort(function(a, b) {
        var x = parseValue(a.cells[n].innerText, type);
        var y = parseValue(b.cells[n].innerText, type);
        if(x < y) return dir === 'asc' ? -1 : 1;
        if(x > y) return dir === 'asc' ? 1 : -1;
        return 0;
    });
    rows.forEach(function(row) { table.tBodies[0].appendChild(row); });
    table.setAttribute('data-sortdir'+n, dir);
}
</script>
</head>
<body>
<h1>Inactive Guest Accounts Report</h1>
<table id="GuestStats">
<thead>
<tr>
<th onclick="sortTable(0,'string')">Guest</th>
<th onclick="sortTable(1,'string')">Email</th>
<th onclick="sortTable(2,'string')">Sponsors</th>
<th onclick="sortTable(3,'date')">Creation date</th>
<th onclick="sortTable(4,'date')">Date of last successful signin</th>
<th onclick="sortTable(5,'date')">Date last audit activity</th>
<th onclick="sortTable(6,'number')">Number of audit activities</th>
<th onclick="sortTable(7,'string')">Top 3 activities</th>
<th onclick="sortTable(8,'string')">Guest status</th>
</tr>
</thead>
<tbody>
"@

$HtmlRows = foreach ($Row in $Report ) {
    "<tr><td>$($row.Guest)</td><td>$($row.Email)</td><td>$($row.Sponsors)</td><td>$($row.'Creation date')</td><td>$($row.'Date of last successful signin')</td><td>$($row.'Date last audit activity')</td><td>$($row.'Number of audit activities')</td><td>$($row.'Top 3 activities')</td><td>$($row.'Guest status')</td></tr>"
}

$HtmlFooter = @"
</tbody>
</table>
</body>
</html>
"@

#Generate the full HTML content and save it to a file
If ($Interactive) {
    $HTMLReportFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveGuests.html"
} Else {
    $HTMLReportFile = $CurrentFolder + "\InactiveGuests.html"
}

$HTMLFile = $HtmlHeader + ($HtmlRows -join "`n") + $HtmlFooter
$HTMLFile | Out-File -FilePath $HTMLReportFile -Encoding utf8
Write-Host ("HTML report written to {0}" -f $HTMLReportFile) -ForegroundColor Green

# And generate an output file
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $true
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveGuests.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    } 
    $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Inactive Guests" -Title ("Inactive Guests Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "InactiveGuests" 
    $AttachmentFile = $ExcelOutputFile
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveGuests.CSV"
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
    $AttachmentFile = $CSVOutputFile
}

If ($ExcelGenerated) {
    Write-Host ("Excel worksheet output written to {0}" -f $ExcelOutputFile)
} Else {
    Write-Host ("CSV output file written to {0}" -f $CSVOutputFile)
}   

# Encode the output file to an email
$EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($AttachmentFile))
# Encode the HTML report too
$EncodedHTMLReportFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($HTMLReportFile))

$MsgAttachments = @(
    @{
        '@odata.type' = '#microsoft.graph.fileAttachment'
        Name = (Split-Path $AttachmentFile -Leaf)
        ContentBytes = $EncodedAttachmentFile
        ContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    },
    @{
        '@odata.type' = '#microsoft.graph.fileAttachment'
        Name = (Split-Path $HTMLReportFile -Leaf)
        ContentBytes = $EncodedHTMLReportFile
        ContentType = 'text/html'
    }
)

# Build the array of a single TO recipient detailed in a hash table - change this to the appropriate recipient for your tenant
$ToRecipient = @{}
$ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
[array]$MsgTo = $ToRecipient
# Define the message subject
$MsgSubject = "Important: Inactive Guests Report"
# Create the HTML content
$HtmlMsg = "</body></html><p>The output files for the <b>Inactive Guests Report</b> are attached to this message. Please review the information at your convenience</p>"
# Construct the message body 	
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

# Send the message
Try {
    Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters -ErrorAction Stop
    Write-Output ("Inactive guest account report emailed to {0}" -f $ToRecipient.emailAddress.address)
} Catch {
    Write-Output "Unable to send email"
    Write-Output $_.Exception.Message
}

Write-Output "All done"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.