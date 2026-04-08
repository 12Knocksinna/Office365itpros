# Report-RecentDLChanges.PS1
# A script to report recent additions, changes, and deletions of Exchange Online distribution lists over the last 30 days. The script also generates
# a list of all the Exchange Online distribution lists in the tenant, along with details of their membership and ownership. The report is generated in HTML format, and also in either Excel or CSV format depending on the availability of the ImportExcel module.

# V1.0 6-Apr-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-RecentDLChanges.PS1

# Connect to Exchange Online PowerShell
[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Connecting to Exchange Online PowerShell..." -ForegroundColor Green
    Connect-ExchangeOnline -ShowBanner:$false
}   

Write-Host "Finding Exchange Online distribution lists in the tenant..." -ForegroundColor Green
# Find all Exchange Online distribution lists in the tenant
[array]$DLs = Get-DistributionGroup -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -eq "MailUniversalDistributionGroup" }

If ($DLs) {
    Write-Host ("Found {0} Exchange Online distribution lists in the tenant." -f $DLs.Count) -ForegroundColor Green
} Else {
    Write-Host "No Exchange Online distribution lists found in the tenant." -ForegroundColor Yellow
    Break
}

# Search the audit log for distribution list operations in the last 30 days
$StartDate = (Get-Date).AddDays(-30)
$EndDate = Get-Date

# Distribution list administration operations include adding, updating, and deleting distribution lists - https://learn.microsoft.com/purview/audit-log-activities?WT.mc_id=M365-MVP-9501#microsoft-entra-group-administration-activities
$Operations = "New-DistributionGroup", "Set-DistributionGroup", "Remove-DistributionGroup", "Add-DistributionGroupMember", "Remove-DistributionGroupMember"
Write-Host "Searching the audit log for distribution list operations from $($StartDate.ToShortDateString()) to $($EndDate.ToShortDateString())..." -ForegroundColor Green   
[array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations $Operations -Formatted -ResultSize 5000 `
    -SessionCommand ReturnLargeSet -RecordType ExchangeAdmin

If ($Records) {
    # Remove duplicate records based on the identity field
    $Records = $Records | Sort-Object Identity -Unique
    Write-Host ("Found {0} distribution list operations in the audit log for the specified date range." -f $Records.Count) -ForegroundColor Green
} Else {
    Write-Host "No distribution list operations found in the audit log for the specified date range." -ForegroundColor Yellow
    Break
}

$DLHash = @{}
$DLMemberHash = @{}
Write-Host "Analyzing audit log records..." -ForegroundColor Green
$AuditReport = [System.Collections.Generic.List[Object]]::new()

$Records = $Records | Sort-Object {$_.CreationDate -as [datetime]}
ForEach ($Rec in $Records) {
    $DLMemberDisplayName = $null
    $AuditData = ConvertFrom-Json -InputObject $Rec.AuditData 
    $DLIdentity = $Auditdata.parameters | Where-Object {$_.Name -eq 'Identity'} | Select-Object -ExpandProperty Value
    $DLMemberId = $Auditdata.parameters | Where-Object {$_.Name -eq 'Member'} | Select-Object -ExpandProperty Value

    # If the action removes a distribution group, we won't be able to retrieve details of the group or its members, so we'll just use the identity and display name
    $DLMemberId = $null
    Switch ($Rec.Operations) {
        "Remove-DistributionGroup" {
            $DLIdentity = $Auditdata.parameters | Where-Object {$_.Name -eq 'Identity'} | Select-Object -ExpandProperty Value
            $DLDisplayName = $DLIdentity.Split("onmicrosoft.com/")[1] + " (Deleted)"
        }
        "New-DistributionGroup" {
            # For new distribution groups, get the display name and identity and store the data in the hash table
            $DLIdentity = $AuditData.ObjectId
            $DLDisplayName = $DLIdentity.Split("onmicrosoft.com/")[1]
            # Possibility exists that we process an audit record for the update of a new distribution group before we process the audit record for the creation of the distribution group, so we should only add the new distribution group to the hash table if it doesn't already exist
            If ($DLHash.Keys -NotContains $DLIdentity) {
                $DLHash.Add($DLIdentity, $DLDisplayName)
            }
        }
        Default {
            If ($DLHash.ContainsKey($DLIdentity)) {
                $DLDisplayName = $DLHash[$DLIdentity]
            } Else {
                Try {
                    $DLDisplayName = (Get-Recipient -Identity $DLIdentity -ErrorAction Stop).DisplayName 
                    If ($DLHash.Keys -NotContains $DLIdentity) {
                        $DLHash.Add($DLIdentity, $DLDisplayName)
                    }  
                } Catch {
                    Write-Host ("Error retrieving distribution list details for {0} (probably removed):" -f $DLIdentity) -ForegroundColor Red
                    $DLDisplayName = $DLIdentity.Split("onmicrosoft.com/")[1] + " (Deleted)"
                }    
            }
            $DLMemberId = $Auditdata.parameters | Where-Object {$_.Name -eq 'Member'} | Select-Object -ExpandProperty Value
        }
    }

    If ($DLMemberId) {
        If ($DLMemberHash.ContainsKey($DLMemberId)) {
            $DLMemberDisplayName = $DLMemberHash[$DLMemberId]
        } Else {
            Try {
                $DLMemberDisplayName = (Get-Recipient -Identity $DLMemberId -ErrorAction Stop).DisplayName 
            $DLMemberHash.Add($DLMemberId, $DLMemberDisplayName)
            } Catch {
                Write-Host ("Error retrieving member details for {0}:" -f $DLMemberId) -ForegroundColor Red
            }
        }
    }

    $ReportItem = [PSCustomObject]@{
        DLName      = $DLDisplayName
        DLId        = $DLIdentity
        Operation   = $Rec.Operations
        Actor       = $Rec.UserIds
        Member      = $DLMemberDisplayName
        Date        = Get-Date $Rec.CreationDate -format 'dd/MM/yyyy HH:mm:ss'
    }
    $AuditReport.Add($ReportItem)
}

$AuditReport = $AuditReport | Sort-Object {$_.Date -as [datetime]}

Write-Host "Analyzing distribution list details..." -ForegroundColor Green
$DLReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($DL in $DLs) {
    [array]$DLMembers = $null
    Write-Host "Analyzing distribution list $($DL.DisplayName)..." -ForegroundColor Green
    [array]$DLMembers = Get-DistributionGroupMember -Identity $DL.Alias -ResultSize Unlimited
    [array]$DLMemberMailboxes = ($DLMembers | Where-Object { $_.RecipientTypeDetails -eq "UserMailbox" }).DisplayName
    [array]$DLMemberContacts = ($DLMembers | Where-Object { $_.RecipientTypeDetails -eq "MailContact" }).DisplayName
    [array]$DLMemberPFs = ($DLMembers | Where-Object { $_.RecipientTypeDetails -eq "PublicFolder" }).DisplayName
    [array]$DLMemberGroups = ($DLMembers | Where-Object { $_.RecipientTypeDetails -eq "MailUniversalDistributionGroup" }).DisplayName
    [array]$DLMemberSharedMailboxes = ($DLMembers | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" }).DisplayName
    $MemberCount =  $DLMembers.Count
    $MemberNames = $DLMembers.DisplayName -join ", "
    [array]$DLOwners = $null
    # Figure out owner details
    ForEach ($Owner in $DL.ManagedBy) {
        Try {
            $OwnerDetails = Get-Recipient -Identity $Owner -ErrorAction Stop
            If ($OwnerDetails.RecipientTypeDetails -eq "UserMailbox") {
                $DLOwners += $OwnerDetails.DisplayName
            } ElseIf ($OwnerDetails.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
                $DLOwners += $OwnerDetails.DisplayName
            } Else {
                $DLOwners += $Owner
            }
        } Catch {
            Write-Host ("Error retrieving owner details for {0}: (probably no owners)" -f $DL.DisplayName) -ForegroundColor Red
            $DLOwners = $null
            Continue
        }
    }
    $OwnerNames = $DLOwners -join ", "
    $OwnerCount = $DLOwners.Count
    Switch ($DL.HiddenFromAddressListsEnabled) {
        "False" { $DLVisibility = "Public" }
        "True" { $DLVisibility = "Private" }
        Default { $DLVisibility = "Unknown" }
    }
    $ReportItem = [PSCustomObject]@{
        DL                  = $DL.displayName
        DLId                = $DL.Identity
        "Member Count"      = $MemberCount
        "Member Names"      = $MemberNames
        "Owner Count"       = $OwnerCount
        "Owner Names"       = $OwnerNames
        "User Members"      = $DLMemberMailboxes.Count
        "Group Members"     = $DLMemberGroups.Count
        "Contact Members"   = $DLMemberContacts.Count
        "Shared Mailbox Members" = $DLMemberSharedMailboxes.Count   
        "Public Folder Members" = $DLMemberPFs.Count
        Visibility          = $DLVisibility
        Created             = $DL.WhenCreated
    }
    $DLReport.Add($ReportItem)
}

$DLReport = $DLReport | Sort-Object DL

If ($AuditReport) {
    Write-Host ("Found {0} unique distribution list operations in the audit log since {1}." -f $Records.Count, $StartDate.ToShortDateString()   ) -ForegroundColor Green
    [array]$DLAdditions = $AuditReport | Where-Object { $_.Operation -eq "New-DistributionGroup" }
    [array]$DLUpdates = $AuditReport | Where-Object { $_.Operation -eq "Set-DistributionGroup" }
    [array]$DLRemovals = $AuditReport | Where-Object { $_.Operation -eq "Remove-DistributionGroup" }
    [array]$DLMemberAdditions = $AuditReport | Where-Object { $_.Operation -eq "Add-DistributionGroupMember" }
    [array]$DLMemberRemovals = $AuditReport | Where-Object { $_.Operation -eq "Remove-DistributionGroupMember" }
    Write-Host ("Distribution List Additions: {0}" -f $DLAdditions.Count) -ForegroundColor Cyan
    Write-Host ("Distribution List Member Additions: {0}" -f $DLMemberAdditions.Count) -ForegroundColor Cyan
    Write-Host ("Distribution List Member Removals: {0}" -f $DLMemberRemovals.Count) -ForegroundColor Cyan
    Write-Host ("Distribution Lists Updates: {0}" -f $DLUpdates.Count) -ForegroundColor Cyan
    Write-Host ("Distribution List Removals: {0}" -f $DLRemovals.Count) -ForegroundColor Cyan  
} Else {
    Write-Host "No recent distribution list operations found in the audit log." -ForegroundColor Yellow 
}

Write-Host "Generating HTML report..." -ForegroundColor Green
$CreationDate = Get-Date -format 'dd-MMM-yyyy HH:mm:ss'
$Version = "1.0"
$Organization = (Get-OrganizationConfig).DisplayName

# Create the HTML report
$htmlhead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 48px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 30px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Distribution List Operations Report</h1></p>
           <p><h2><b>For the " + $Organization + " organization</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format g) + "</h3></p></div>"

$HTMLAnalysisSectionHead = "<p><h2><u>Analysid of Distribution List Activity</u></h2></p>" +
                   "<p>The report shows all distribution lists in the tenant, along with details of their membership and ownership. The report also includes a list of recent distribution list operations in the tenant, including additions, updates, and removals.</p>" +
                   "<p>Use this report to understand the current state of distribution lists in the tenant, and to identify any recent changes to distribution lists that may require attention.</p>"
$HTMLAnalysisAdditions = "<p><h3>Distribution List Additions</h3></p>" +
                 "<p>The following distribution lists have been added in the last 30 days:</p>" +
                 ($AuditReport | Where-Object { $_.Operation -eq "New-DistributionGroup" }| Select-Object -Property DLName, Actor, Date | `
                 ConvertTo-Html -Fragment)
$HTMLAnalysisUpdates = "<p><h3>Distribution List Updates</h3></p>" +
                 "<p>The following distribution lists have been updated in the last 30 days:</p>" +
                 ($AuditReport | Where-Object { $_.Operation -eq "Set-DistributionGroup" }| Select-Object -Property DLName, Actor, Date | ConvertTo-Html -Fragment)
$HTMLAnalysisRemovals = "<p><h3>Distribution List Removals</h3></p>" +
                 "<p>The following distribution lists have been removed in the last 30 days:</p>" +
                 ($AuditReport | Where-Object { $_.Operation -eq "Remove-DistributionGroup" }| Select-Object -Property DLName, Actor, Date | ConvertTo-Html -Fragment)
$HTMLAnalysisMemberAdditions = "<p><h3>Distribution List Member Additions</h3></p>" +
                 "<p>The following distribution list members have been added in the last 30 days:</p>" +
                 ($AuditReport | Where-Object { $_.Operation -eq "Add-DistributionGroupMember" }| Select-Object -Property DLName, Actor, Member, Date | ConvertTo-Html -Fragment)
$HTMLAnalysisMemberRemovals = "<p><h3>Distribution List Member Removals</h3></p>" +
                 "<p>The following distribution list members have been removed in the last 30 days:</p>" +
                 ($AuditReport | Where-Object { $_.Operation -eq "Remove-DistributionGroupMember" }| Select-Object -Property DLName, Actor, Member,Date | ConvertTo-Html -Fragment)

$HTMLBody = $HTMLHead + $HTMLAnalysisSectionHead
If ($DLAdditions.Count -gt 0) {
    $HTMLBody += $HTMLAnalysisAdditions
} Else {
    $HTMLBody += "<p><h3>Distribution List Additions</h3></p><p>No additions of distribution lists found for the last 30 days.</p>"
}
If ($DLUpdates.Count -gt 0) {
    $HTMLBody += $HTMLAnalysisUpdates
} Else {
    $HTMLBody += "<p><h3>Distribution List Updates</h3></p><p>No updates for distribution lists found for the last 30 days.</p>"
}
If ($DLRemovals.Count -gt 0) {
    $HTMLBody += $HTMLAnalysisRemovals
} Else {
    $HTMLBody += "<p><h3>Distribution List Removals</h3></p><p>No removals of distribution lists found for the last 30 days.</p>"
}
If ($DLMemberAdditions.Count -gt 0) {
    $HTMLBody += $HTMLAnalysisMemberAdditions
} Else {
    $HTMLBody += "<p><h3>Additions of Members to Distribution List</h3></p><p>No additions of distribution list members found for the last 30 days.</p>"
}
If ($DLMemberRemovals.Count -gt 0) {
    $HTMLBody += $HTMLAnalysisMemberRemovals
} Else {
    $HTMLBody += "<p><h3>Removals of Members from Distribution List</h3></p><p>No removals of distribution list members found for the last 30 days.</p>"
}

$DLDetailsHeader = "<p><h2>Details of Distribution Lists</h2></p>" +
                      "<p>The following table shows details of all distribution lists in the tenant, including their membership and ownership information.</p>"
$DLDetailsTable = $DLReport | Select-Object DL, Created, "Member Count", "Owner Count", "Owner Names", "Member Names", "User Members", "Group Members", "Contact Members", "Shared Mailbox Members", "Public Folder Members"| ConvertTo-Html -Fragment
$HTMLBody = $HTMLBody + $DLDetailsHeader + $DLDetailsTable    

$HTMLTail = "<p><h3>Report created for: " + $Organization + "</h3></p>" +
             "<p>Created: " + $CreationDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of distribution lists found:                     " + $DLReport.Count + "</p>" +
             "<p>Number of distribution lists without a manager:         " + ($DLReport | Where-Object { $_."Owner Count" -eq 0 }).Count + "</p>" +
             "<p>Number of distribution lists with no members:           " + ($DLReport | Where-Object { $_."Member Count" -eq 0 }).Count + "</p>" +
             "<p>Number of recent distribution list audit events found:  " + $AuditReport.Count + "</p>" +
             "<p>Number of recent distribution list additions:           " + $GroupAdditions.Count + "</p>" +
             "<p>Number of recent distribution list updates:             " + $GroupUpdates.Count + "</p>" +
             "<p>Number of recent distribution list removals:            " + $GroupRemovals.Count + "</p>" +
             "<p>Number of recent distribution list member additions:    " + $DLMemberAdditions.Count + "</p>" +
             "<p>Number of recent distribution list member removals:     " + $DLMemberRemovals.Count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Distribution Lists Report <b>" + $Version + "</b>"	

$HTMLReport = $HTMLBody + $HTMLTail
$HTMLReportFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Distribution Lists.html"
$HTMLReport | Out-File $HTMLReportFile -Encoding UTF8

# Generate the report in either Excel worksheet or CSV format, depending on if the ImportExcel module is available
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Distribution Lists.xlsx"
    $DLReport | Export-Excel -Path $ExcelOutputFile -WorksheetName "Distribution Lists" -Title ("Distribution Lists {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "DistributionLists" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Distribution Lists.CSV"
    $DLReport | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}   

If ($ExcelGenerated -eq $true) {
    Write-Host ("Distribution lists report is available in Excel workbook {0}" -f $ExcelOutputFile)
    $AttachmentFile = $ExcelOutputFile
} Else {
    Write-Host ("Distribution lists report is available in CSV file {0}" -f $CSVOutputFile)
    $AttachmentFile = $CSVOutputFile
}

Write-Host ("HTML report of distribution lists is available at {0}" -f $HTMLReportFile) -ForegroundColor Green

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
$MsgSubject = "Distribution Lists Report - $(Get-Date -Format 'dd-MMM-yyyy')"

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
    Write-Output ("Distribution lists report emailed to {0}" -f $DestinationEmailAddress)
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