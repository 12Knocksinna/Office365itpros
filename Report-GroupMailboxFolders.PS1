# Report-GroupMailboxFolders.PS1
# This script is to demonstrate how to use the Exchange Import-Export Graph API to access information in group mailboxes.
# It also serves how to use the API in general

# V1.0 15-Sept-2026
# GitHub Link: 

function Get-GraphMailboxFolders {
    param(
        [string]$MailboxId
    )

    # A function to retrieve folders from a group mailbox. The function only handles up to 500 top-level folders. It does not
    # pagination child folders, and it only goes down one level for child folders.

    #Get top level folders
    $GroupMailboxFolders = [System.Collections.Generic.List[Object]]::new()
    $Uri = "https://graph.microsoft.com/v1.0/admin/exchange/mailboxes/$MailboxId/folders?`$top=500"
    $Data = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop | Select-Object -ExpandProperty Value
    $GroupMailboxFolders = $Data

    # Are there any child folders to process
    [array]$FoldersWithChildFolders = $GroupMailboxFolders | Where-Object {$_.childFolderCount -gt 0}
    
    If (!$FoldersWithChildFolders) {
       return $GroupMailboxFolders
    }

    ForEach ($FolderWithChildren in $FoldersWithChildFolders) {
      $ChildFoldersUri = ("https://graph.microsoft.com/v1.0/admin/exchange/mailboxes/{0}/folders/{1}/childFolders" -f $MailboxId,$FolderWithChildren.Id)
      $Data = Invoke-MgGraphRequest -Method GET -Uri $ChildFoldersUri -OutputType PSObject -ErrorAction Stop | Select-Object -ExpandProperty Value
      If ($Data) {
         $GroupMailboxFolders += $Data
      }   
    }
    return $GroupMailboxFolders
}

$AppId = "ecf80ca6-303d-4ce3-ac80-113ca4405746"
$TenantId = "b662313f-14fc-43a2-9a7a-d2e27f4f3478"
$Thumbprint = "A43876D6692CF5366FE872175B5CE27482774C26"
# Not interested in Outlook groups with just a few items in the Inbox
[int]$ThresholdInboxItems = 25
# Tenant name to use with Connect-ExchangeOnline - requires the service domain
$TenantName = 'office365itpros.onmicrosoft.com'
# Define filter to expand single value extension properties fetched using the Import-Export API
$PropertyFilter = "id eq 'String 0x0037' or id eq 'String 0x1035' or id eq 'String 0x1000' or id eq 'Binary 0x1013' or id eq 'Boolean 0x0E1B' or id eq 'Binary 0x3013' or id eq 'String 0x0C1A' or id eq 'String 0x0C1F'"
$EncodedPropertyFilter = [uri]::EscapeDataString($PropertyFilter)
[array]$FoldersNotToProcess = "Archive", "Drafts", "ExternalContacts", "Files", "Journal", "Outbox", "Synx Issues", "Tasks", "Versions", `
    "Conversation Action Settings", "Conversation History", "Team Chat", "Junk Email", "Notes", "PersonMetadata", "Yammer Root", "SubstrateHolds",`
    "Inbound", "Outbound", "Calendar Logging", "Audits", "Feeds", "Recoverable Items", "Deletions", "Purges", "DiscoveryHolds", "MigratedMessages", `
    "Contacts", "Companies", "GAL Contacts", "Organizational Contacts", "Recipient Cache", "PeopleCentricConversation Buddies", "Clutter", "Calendar", `
    "{06967759-274D-40B2-A3EB-D7F9E73727D7}", "{A9E2BC46-B3A0-4243-B315-60D991004455}", "Legacy Archive Journals", "MeContact", "Sync Issues", `
    "Top of Information Store", "SearchDiscoveryHoldsFolder"

Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome

Write-Output "Finding Outlook groyups..."
# Find Microsoft 365 Groups and Teams, and discard any group that is team-enabled
[array]$M365Groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All -Property id, displayName
[array]$Teams = Get-MgTeam -All -Property Id, displayName
[array]$OutlookGroups = $M365Groups | Where-Object { $_.GroupId -notIn $Teams.Id} | Sort-Object DisplayName

If ($OutlookGroups.Count -eq 0) {
    Write-Output "No group mailboxes found."
    Return
} Else {
    Write-Output "Found $($OutlookGroups.Count) group mailboxes."
}

# Connect to Exchange Online in app-only mode to fetch the statistics for the group mailboxes
Write-Output "Connecting to Exchange Online..."
Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $Thumbprint -Organization $TenantName -ShowBanner:$false
# [array]$Mbx = Get-Mailbox -GroupMailbox -ResultSize Unlimited

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Mailbox in $OutlookGroups) {
    Write-Output "Processing mailbox $($Mailbox.DisplayName)"
    [Array]$MbxStats = Get-MailboxFolderStatistics -Identity $Mailbox.Id
    $InboxData = $MbxStats | Where-Object {$_.FolderType -eq "Inbox"}
    If ($InboxData.ItemsInFolder -lt $ThresholdInboxItems) {
        Write-Output "Mailbox $($Mailbox.DisplayName) has less than $ThresholdInboxItems items in the Inbox."
        Continue
    } Else {
        # Drop all the folders we don't care about
        $FoldersToProcess = $MbxStats | Where-Object { $_.Name -notin $FoldersNotToProcess }
        ForEach ($Folder in $FoldersToProcess) {
            If ($Folder.ItemsInFolder -gt 0) { 
                # If a folder has items, process it
                $ConvertedFolderId = [Convert]::FromBase64String($Folder.FolderId)
                $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
                $nibbler = $encoding.GetBytes("0123456789ABCDEF")
                $indexIdBytes = New-Object byte[] 48; $indexIdIdx = 0
                $ConvertedFolderId | Select-Object -Skip 23 -First 24 | `
                    ForEach-Object { $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]; $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0x0F]
                }
                $FolderIdentifier = $encoding.GetString($indexIdBytes)

                # Write details about the folder to the report  
                $ReportLine = [PSCustomObject]@{
                    DisplayName      = $Mailbox.DisplayName
                    FolderItems      = $Folder.ItemsInFolder
                    FolderName       = $Folder.Name
                    GroupId          = $Mailbox.Id
                    FolderId         = $Folder.FolderId
                    FolderIdentifier = $FolderIdentifier
                    MailboxId        = $InboxData.ContentMailboxGuid.Guid
                }
                $Report.Add($ReportLine)
            }
        }
    }           
}

If ($Report.Count -eq 0) {
    Write-Output "No group mailboxes have more than $ThresholdInboxItems items in the Inbox."
} Else {
    Write-Output "Group mailboxes with more than $ThresholdInboxItems items in the Inbox:"
    $Report | Format-Table -AutoSize
}

Write-Output "Proceeding to process items from group mailboxes"
$MailboxReport = [System.Collections.Generic.List[Object]]::new()
$MailboxSummaryReport = [System.Collections.Generic.List[Object]]::new()
[array]$Mailboxes = $Report | Sort-Object -Property MailboxId -Unique
ForEach ($Mailbox in $Mailboxes) {
    Write-Output ""
    Write-Output "Processing mailbox $($Mailbox.DisplayName)"
    # Construct MailboxURI for use with Import-Export API
    $MailboxId = ("MBX:{0}@{1}" -f $Mailbox.MailboxId, $TenantId)
    # Replace unsafe characters in the MailboxId for use in a URI
    $MailboxId = [uri]::EscapeDataString($MailboxId)
    # Find folders in the group mailbox
    $GroupMailboxFolders = Get-GraphMailboxFolders -MailboxId $MailboxId
    If ($GroupMailboxFolders.Count -eq 0) {
        Write-Output "No folders found for mailbox $($Mailbox.DisplayName)"
        Continue
    } Else {
        $GroupMailboxFolders = $GroupMailboxFolders | Sort-Object displayName
    }

    # Select the folders of interest from the report
    [array]$FoldersToProcess = $Report | Where-Object { $_.MailboxId -eq $Mailbox.MailboxId } | Select-Object DisplayName, FolderIdentifier, FolderName
    ForEach ($Folder in $FoldersToProcess) {
        $MailboxFolder = $GroupMailboxFolders | Where-Object {$_.displayName -eq $Folder.FolderName} | Select-Object -First 1
        If ($null -eq $MailboxFolder) {
            Write-Output "Folder $($Folder.FolderName) not found in mailbox $($Mailbox.DisplayName)"
            Continue
        }
        Write-Output ""
        Write-Output "Processing folder $($Folder.FolderName) in mailbox $($Mailbox.DisplayName)"
        $MailboxItemsUri = ("https://graph.microsoft.com/v1.0/admin/exchange/mailboxes/{0}/folders/{1}/items?`$top=500&`$expand=singleValueExtendedProperties(`$filter=$EncodedPropertyFilter)" -f $MailboxId, $MailboxFolder.Id)
        $NextLink = $null; [array]$ItemData = @()
        Try {
            [array]$Data = Invoke-MgGraphRequest -Uri $MailboxItemsUri -Method GET -OutputType PsObject -ErrorAction Stop 
            $NextLink = $Data.'@Odata.NextLink'
            $ItemData = $Data.Value
        } Catch {
            Write-Output "Failed to retrieve items for folder $($Folder.FolderName) in mailbox $($Mailbox.DisplayName): $_"
            Continue
        }
        # Pagination to pick up additional pages
        If ($NextLink) {
            While ($null -ne $NextLink) {
                [array]$Data = Invoke-MgGraphRequest -Uri $NextLink -Method GET -OutputType PsObject -ErrorAction Stop 
                $ItemData += $Data.Value 
                $NextLink = $Data.'@Odata.NextLink'  
            }
        }
        # Display the first few items in the folder for inspection
        # $Data | Select-Object -first 5 | Format-Table CreatedDateTime, Size, @{n="Subject"; e= {($_.singleValueExtendedProperties | Where-Object Id -eq 'String 0x37').Value }}
        
        [datetime]$LastPostDateTime = ($ItemData | Sort-Object CreatedDateTime -Descending | Select-Object -First 1).CreatedDateTime
        $DaysSinceLastPost = (Get-Date) - $LastPostDateTime
        $DaysSinceLastPost = $DaysSinceLastPost.Days    
        [datetime]$FirstPostDateTime = ($Data | Sort-Object CreatedDateTime | Select-Object -First 1).CreatedDateTime
        $DaysSinceFirstPost = (Get-Date) - $FirstPostDateTime
        $DaysSinceFirstPost = $DaysSinceFirstPost.Days    
        # Who sends most posts to the group mailbox?
        $ItemData | Select-Object @{ Name = 'Sender'; Expression = { ($_.singleValueExtendedProperties | Where-Object Id -eq 'String 0xc1a').Value }} |
            Group-Object -NoElement Sender | Sort-Object Count -Descending | ForEach-Object {
                $MailboxSummaryReportLine = [PSCustomObject]@{
                    Mailbox = $Mailbox.DisplayName
                    Folder  = $Folder.FolderName
                    Sender  = $_.Name
                    Count   = $_.Count
                    LastPostDateTime = Get-Date $LastPostDateTime -format "dd-MMM-yyyy HH:mm:ss"
                    DaysSinceLastPost = $DaysSinceLastPost
                    FirstPostDateTime = Get-Date $FirstPostDateTime -format "dd-MMM-yyyy HH:mm:ss"
                    DaysSinceFirstPost = $DaysSinceFirstPost
                }
            $MailboxSummaryReport.Add($MailboxSummaryReportLine)
        }
        # Populate the report
        $ItemData | Select-Object -first 10 | ForEach-Object {
            $MailboxReportLine = [PSCustomObject]@{
                Mailbox = $Mailbox.DisplayName
                Folder = $Folder.FolderName
                CreatedDateTime = Get-Date $_.CreatedDateTime -format "dd-MMM-yyyy HH:mm:ss"
                'Size (MB)' = [math]::Round($_.Size/1MB, 2)
                Subject = ($_.singleValueExtendedProperties | Where-Object Id -eq 'String 0x37').Value
                Sender = ($_.singleValueExtendedProperties | Where-Object Id -eq 'String 0xc1a').Value
            }
            $MailboxReport.Add($MailboxReportLine)
        } 
    }
  
}


# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 