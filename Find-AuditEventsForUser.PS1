# Find-AuditEventsForUser.PS1
# Find and report audit events for actions taken by a user over a specified period.

# Demo for ESPC 2024 in Stockholm
# V1.0 28-Nov-2024
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-AuditEventsForUser.PS1

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ($Modules -notcontains "ExchangeOnlineManagement") {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$False
}   

$User = Read-Host "Enter the user's UPN"
$StartDate = Read-Host "Enter the start date (dd-mm-yyyy)"

If (!(Get-ExoMailbox -Identity $User)) {
    Write-Host "User $User not found"
    Exit
}

$EndDate = (Get-Date $StartDate).AddDays(1)

Write-Host ("Searching the audit log for actions for {0} between {1} and {2}" -f $User, $StartDate, (Get-Date $EndDate -format 'dd-MMM-yyyy'))
[array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -UserIds $User -ResultSize 5000 -SessionCommand ReturnLargeSet
If ($Records.Count -eq 0) {
    Write-Host "No audit records found for $User between $StartDate and $EndDate"
    Exit
}
$Records = $Records | Sort-Object Identity -Unique
$Records = $Records | Sort-Object {$_.CreationDate -as [datetime]} 
Write-Host ("Found {0} audit records for {1}. First record from {2}. Last record at {3}" -f $Records.Count, $User, $Records[0].CreationDate, $Records[-1].CreationDate) 

# Overall summary
$RecordsSummary = $Records | Group-Object Operations -NoElement | Sort-Object Count -Descending
Write-Host "Summary of actions taken by $User"
$RecordsSummary | Format-Table Name, Count -AutoSize

# Remove UserLoggedIn events
[array]$Records1 = $Records | Where-Object {$_.Operations -ne "UserLoggedIn"}
$Report = [System.Collections.Generic.List[Object]]::new()

Write-Host ("Processing {0} audit records for {1} between {2} and {3}" -f $Records1.Count, $User, $StartDate, $EndDate)
# Process each audit record and attempt to make sense of the audit data payload
ForEach ($Rec in $Records1) {
    $Action = $null; $Object = $null; $Location = $null; $Workload = $null
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    Switch ($Rec.Operations) {
        "Create" {
            $Action = 'Item created in mailbox'
            $Object = $AuditData.Item.Subject
            $Location = ($AuditData.MailboxOwnerUPN + $AuditData.Item.ParentFolder.Path)
            $Workload = "Exchange Online"
        }
        "CopilotInteraction" {
            $Action = 'Copilot interaction'
            $Object = $AuditData.CopilotEventData.ThreadId
            $Location = $AuditData.CopilotEventData.AppHost
            $Workload = "Copilot"
        }
        "FileAccessed" {
            $Action = 'File accessed'
            $Object = $AuditData.SourceFileName
            $Location = $AuditData.ObjectId
            $Workload = "SharePoint Online"
        }
        "FileModified" {
            $Action = 'File modified'
            $Object = $AuditData.SourceFileName
            $Location = $AuditData.ObjectId
            $Workload = "SharePoint Online"
        }
        "FileRenamed" {
            $Action = 'File renamed'
            $Object = $AuditData.SourceFileName
            $Location = $AuditData.ObjectId
            $Workload = "SharePoint Online"
        }
        "FolderCreated" {
            $Action = 'Folder created'
            $Object = $AuditData.SourceFileName
            $Location = $AuditData.ObjectId
            $Workload = "SharePoint Online"
        }
        "FolderModified" {
            $Action = 'Folder modified'
            $Object = $AuditData.SourceFileName
            $Location = $AuditData.ObjectId
            $Workload = "SharePoint Online"
        }
        "FileModifiedExtended" {
            $Action = 'File modified'
            $Object = $AuditData.SourceFileName
            $Location = $AuditData.ObjectId
            $Workload = "SharePoint Online"
        }
        "FileSensitivityLabelApplied" {
            $Action = 'Sensitivity label applied to file'
            $Object = $AuditData.SourceFileName
            $Location = $AuditData.ObjectId
            $Workload = "SharePoint Online"
        }
        "MailItemsAccessed" {
            $Action = 'Mail item accessed'
            If ($AuditData[0].Item.ParentFolder.Name) {
                $Object = $AuditData[0].Item.ParentFolder.Name
            } Else {
                $Object = 'Message Bind'
            }
            $Location = ($AuditData[0].MailboxOwnerUPN + $AuditData.Item.ParentFolder.Path)
            $Workload = "Exchange Online"
        }
        "MoveToDeletedItems" {
            $Action = 'Item moved to Deleted Items'
            $Object = $Auditdata[0].AffectedItems.Subject
            $Location = ($AuditData[0].MailboxOwnerUPN + $AuditData[0].AffectedItems.ParentFolder.Path)
            $Workload = "Exchange Online"
        }
        "HardDelete" {
            $Action = 'Item purged from mailbox'
            $Object = $Auditdata.AffectedItems[0].Subject
            $Location = ($AuditData.MailboxOwnerUPN + $AuditData.AffectedItems[0].ParentFolder.Path)
            $Workload = "Exchange Online"
        }
        "SoftDelete" {
            $Action = 'Item removed to Recoverable Items'
            $Object = $Auditdata.AffectedItems[0].Subject
            $Location = ($AuditData.MailboxOwnerUPN + $AuditData.AffectedItems[0].ParentFolder.Path)
            $Workload = "Exchange Online"
        }
        "Send" {
            $Action = 'Message Sent'
            $Object = $AuditData[0].Item.Subject
            $Location = ($AuditData[0].MailboxOwnerUPN + $AuditData[0].Item.ParentFolder.Path)
            $Workload = "Exchange Online"
        }
        "SendAs" {
            $Action = 'Message sent as another user'
            $Object = $AuditData[0].Item.Subject
            $Location = ($AuditData[0].MailboxOwnerUPN + $AuditData[0].Item.ParentFolder.Path)
            $Workload = "Exchange Online"
        }
        "MessageSent" {
            $Action = 'Teams message sent'  
            $Object =   $AuditData.MessageId
            $Location = ("Posted to {0} in {1}" -f $AuditData.ChannelName, $AuditData.TeamName) 
            $Workload = "Microsoft Teams" 
        }
        "MessageCreatedHasLink" {
            $Action = 'Teams message created with link'
            $Object = $AuditData.MessageId
            $Location = ("Posted to {0} in {1}" -f $AuditData.ChannelName, $AuditData.TeamName) 
            $Workload = "Microsoft Teams" 
        }
        "TaskCreated" {
            $Action = 'Task created'
            $Object = $AuditData.ObjectId
            $Location = $AuditData.Workload
            $Workload = "Planner"
        }
        "TaskModified" {
            $Action = 'Task modified'
            $Object = $AuditData.ObjectId
            $Location = $AuditData.Workload
            $Workload = "Planner"
        }
        "TaskRead" {
            $Action = 'Task read'
            $Object = $AuditData.ObjectId
            $Location = $AuditData.Workload
            $Workload = "Planner"
        }
        "TaskListRead" {
            $Action = 'Task list read'
            $Object = $AuditData.ObjectId
            $Location = $AuditData.Workload
            $Workload = "Planner"
        }
        "PlanRead" {
            $Action = 'Plan read'
            $Object = $AuditData.ObjectId
            $Location = $AuditData.Workload
            $Workload = "Planner"
        }
        "NewRetentionCompliancePolicy" {
            $Action = 'Retention policy created'
            $Object = $Auditdata.Parameters.Value.Split("-A")[0].Split('"')[1]
            $Location = $AuditData.Workload
            $Workload = "Data lifecycle management"
        }
        "NewRetentionComplianceRule" {
            $Action = 'Retention rule created'
            $Object = $Auditdata.Parameters.value[0].Split("-Expiration")[1]
            $Location = $Auditdata.Parameters.Value.Split("-A")[0].Split('"')[1]
            $Workload = "Data lifecycle management"
        }
        "Get-ComplianceSearch" {
            $Action = 'Compliance search retrieved'
            If ($AuditData.Parameters -eq '-ResultSize "Unlimited"' ) {
                $Object = "All results"
            } Else {
                Try {
                    $Object = $AuditData.Parameters.Split('"')[1]               
                } Catch {
                    $Object = "Unknown"
                }
            }
            $Location = $AuditData.Workload
            $Workload = 'eDiscovery'
        }
        "Set-ComplianceSearch" {
            $Action = 'Compliance search updated'
            $Object = $Auditdata.Parameters.Split("-D")[0].Split('"')[1]
            $Location = $AuditData.Workload
            $Workload = 'eDiscovery'
        }
        "Start-ComplianceSearch" {
            $Action = 'Compliance search started'
            $Object =  [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String( $AuditData.Parameters.Split('"')[1].SubString(0,32)))
            $Location = $AuditData.Workload
            $Workload = 'eDiscovery'
        }
        "New-ComplianceSearchAction" {
            $Action = 'Compliance search action created'
            $Object = $AuditData.Parameters
            $Location = $AuditData.Workload
            $Workload = 'eDiscovery'
        }
        "Get-ComplianceCase" {
            $Action = 'Compliance case retrieved'
            $Object = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String( $AuditData.Parameters.Split('"')[1].SubString(0,48)))
            $Location = $AuditData.Workload
            $Workload = "eDiscovery"
        }
        "Get-ComplianceCaseMember" {
            $Action = 'Compliance case membership retrieved'
            $Object = $AuditData.Parameters
            $Location = $AuditData.Workload
            $Workload = "eDiscovery"
        }
        "Set-Mailbox" {
            $Action = 'Mailbox settings updated'
            $Object = $AuditData.Parameters.value[0]
            $Location = $AuditData.Workload
            $Workload = 'Exchange Online PowerShell'
        }
        "Search" {
            $Action = 'Search performed'
            $Object = $AuditData.DataType
            $Location = $AuditData.DatabaseType
            $Workload = "Data Insights"
        }
        "SearchQueryInitiatedSharePoint" {
            $Action = 'SharePoint Search performed'
            $Object = $AuditData.QueryText
            $Location = $AuditData.ScenarioName
            $Workload = "SharePoint Online"
        }
        "SearchQueryInitiatedExchange" {
            $Action = 'Exchange Search performed'
            $Object = $AuditData.QueryText
            $Location = $AuditData.ScenarioName
            $Workload = "Exchange Online"
        }
        "SearchViewed" {
            $Action = 'Search results viewed'
            $Object = $AuditData.ObjectId
            $Location = $AuditData.Query
            $Workload = "eDiscovery"
        }
        "QueryUpdate" {
            $Action = 'Search query updated'
            $Object = $AuditData.QueryText
            $Location = $AuditData.CaseId
            $Workload = "eDiscovery"
        }
        "Validate" {
            $Action = 'Search query validated'
            $Object = $AuditData.DataType
            $Location = $AuditData.RelativeURL
            $Workload = "eDiscovery"
        }
        "GATFRTokenIssue" {
            $Action = 'Get Access Token for Resource'
            $Object = $AuditData.ResourceURL
            $Location = "Outlook Web Access"
            $Workload = "Exchange Online"
        }
        Default {
            $Action = $Rec.Operations
            $Object = $AuditData.ObjectId
            $Location = $AuditData.Workload
            $Workload = $AuditData.Workload
        }
    }

    $ReportLine = [pscustomobject]@{
        CreationDate = $Rec.CreationDate
        Operation    = $Rec.Operations
        User         = $Rec.UserIds
        Record       = $Rec.RecordType
        Action       = $Action
        Object       = $Object
        Location     = $Location
        Workload     = $Workload
    }
    $Report.Add($ReportLine)
}

# Show what we've done...
$Report | Out-GridView -Title "Audit events for $User between $StartDate and $EndDate"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.