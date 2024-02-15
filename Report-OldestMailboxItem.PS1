# Feport-OldestMailboxItems.PS1
# Find mailboxes larger than 50 GB and report the folder name where the oldest item is found
#

function Format-FolderSize {
    # Format File Size nicely
    param (
            [parameter(Mandatory = $true)]
            $InFolderSize
        ) 
    
    If ($InFolderSize -lt 1KB) { # Format the size of a folder
        $OutFolderSize = $InFolderSize.ToString() + " B" 
    } ElseIf ($InFolderSize -lt 1MB) {
        $OutFolderSize = $InFolderSize / 1KB
        $OutFolderSize = ("{0:n2}" -f $OutFolderSize) + " KB"
    } Elseif ($InFolderSize -lt 1GB) {
        $OutFolderSize = $InFoldersize / 1MB
        $OutFolderSize = ("{0:n2}" -f $OutFolderSize) + " MB" 
    } Elseif ($InFolderSize -ge 1GB) {
        $OutFolderSize = $InFolderSize / 1GB
        $OutFolderSize = ("{0:n2}" -f $OutFolderSize) + " GB" 
    }
    Return $OutFolderSize
} 

Connect-ExchangeOnline -SkipLoadingCmdletHelp

[array]$UnwantedFolders = "Top of Information Store", "Quick Step Settings", "Recipient Cache", "Team Chat", "Conversation History", `
   "LinkedIn", "PersonMetadata", "Suggested Contacts", "Snoozed", "Conflicts", "Calendar Logging", "Audits"

Write-Host "Searching for mailboxes..."
[array]$Mailboxes = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -PropertySets Archive -Properties DisplayName | Sort-Object DisplayName
If (!$Mailboxes) {
    Write-Host "Can't find any mailboxes to process - exiting"
    Break
} Else {
    Write-Host ("Found {0} mailboxes to process..." -f $Mailboxes.count)
}

$MbxReport = [System.Collections.Generic.List[Object]]::new()
$FolderReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Mbx in $Mailboxes) {
    Write-Host ("Processing mailbox {0}..." -f $Mbx.DisplayName)
    # Process primary mailbox
    [array]$Stats = Get-ExoMailboxStatistics -Identity $Mbx.ExternalDirectoryObjectId 
    [float]$MbxSizeGB = [math]::Round(($Stats.TotalItemSize.value.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1024MB),2)
        $FolderStats = Get-MailboxFolderStatistics -Identity $Mbx.ExternalDirectoryObjectId -IncludeOldestAndNewestItems | `
          Where-Object {$_.ItemsInFolder -gt 0 -and $_.Name -notin $UnwantedFolders} | Sort-Object Name
        ForEach ($Folder in $FolderStats) {
            [float]$FolderSize = [math]::Round($Folder.FolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",",""),2)
            $FolderSizeReport = Format-FolderSize -InFolderSize $FolderSize
            #[float]$FolderSize = [math]::Round(($Folder.FolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1024),2)
            If (($Folder.Name.Substring(0,1) -ne "{")) {
                $OldestItemDate = $null; $NewestItemDate = $null
                If ($Folder.OldestItemReceivedDate) {
                    $OldestItemDate = Get-Date $Folder.OldestItemReceivedDate -format 'dd-MMM-yyy HH:mm'
                }
                If ( $Folder.NewestItemReceivedDate) {
                    $NewestItemDate = Get-Date $Folder.NewestItemReceivedDate -format 'dd-MMM-yyy HH:mm'
                }
                $FolderReportLine = [PSCustomObject][Ordered]@{ 
                    Mailbox         = $Mbx.displayName
                    UPN             = $Mbx.UserPrincipalName
                    Folder          = $Folder.Name
                    Size            = $FolderSizeReport
                    Items           = $Folder.ItemsInFolder
                    'Oldest item'   = $OldestItemDate
                    'Newest item'   = $NewestItemDate
                    Path            = $Folder.FolderPath
                    Type            = $Folder.FolderType
                }  
                $FolderReport.Add($FolderReportLine)       
            }      
        }
    If ($Mbx.ArchiveStatus -eq 'Active') {
        [array]$ArchiveStats = Get-ExoMailboxStatistics -Identity $Mbx.ExternalDirectoryObjectId -Archive | Sort-Object Name
        [float]$ArchiveMbxSizeGB = [math]::Round(($ArchiveStats.TotalItemSize.value.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1024MB),2)
    } Else {
        $ArchiveStats = $null
    }

    $MbxReportLine = [PSCustomObject][Ordered]@{ 
        Mailbox                         = $Mbx.displayName
        'Primary mailbox size GB'       = ("{0} GB" -f $MbxSizeGB.toString())
        'Primary mailbox total items'   = $Stats.ItemCount
        'Archive status'                = $Mbx.ArchiveStatus
        'Archive mailbox size GB'       = ("{0} GB" -f $ArchiveMbxSizeGB.toString())
        'Archive mailbox total items'   = $ArchiveStats.ItemCount
    }  
    $MbxReport.Add($MbxReportLine)  
  }


 $MbxReport | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.