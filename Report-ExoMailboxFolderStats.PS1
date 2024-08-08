# Feport-OldestMailboxItems.PS1
# Report user mailboxes (including archives) with statistics for each folder
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ExoMailboxFolderStats.PS1
# V1.0 16-Feb-2024

function Format-FolderSize {
    # Format File Size nicely
    param (
            [parameter(Mandatory = $true)]
            $InFolderSize
        ) 
    # Format a size in bytes into KB, MB, or GB
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

# Make sure that we're connected to Exchange Online
[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

# Variables used in the report
$OrgName = Get-OrganizationConfig | Select-Object -ExpandProperty DisplayName
$RunDate = Get-date -format 'dd-MMM-yyyy HH:mm'
$Version = "V1.0"
$HtmlReportFile = "c:\temp\MailboxFolderReport.html"
$MailboxStatsCSV = "c:\temp\MailboxStats.Csv"

# Define folders that we don't want to include in the report
[array]$UnwantedFolders = "Top of Information Store", "Quick Step Settings", "Recipient Cache", "Team Chat", "Conversation History", `
   "LinkedIn", "PersonMetadata", "Suggested Contacts", "Snoozed", "Conflicts", "Calendar Logging", "Audits"
# Define Recoverable Items folders that the report shows in a separate section
[array]$RIFolders = "RecoverableItemsVersions", "RecoverableItemsPurges", "RecoverableItemsDeletions", `
    "RecoverableItemsDiscoveryHolds", "RecoverableItemsSubstrateHolds"

Write-Host "Searching for mailboxes..."
[array]$Mailboxes = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited `
    -PropertySets Archive -Properties DisplayName, RecipientTypeDetails | Sort-Object DisplayName
If (!$Mailboxes) {
    Write-Host "Can't find any mailboxes to process - exiting"
    Break
} Else {
    Write-Host ("Found {0} mailboxes to process..." -f $Mailboxes.count)
}

$MbxReport = [System.Collections.Generic.List[Object]]::new()
$FolderReport = [System.Collections.Generic.List[Object]]::new()
[int]$i=0
ForEach ($Mbx in $Mailboxes) {
    $i++
    Write-Host ("Processing mailbox {0} ({1}/{2})..." -f $Mbx.DisplayName, $i, $Mailboxes.count)
    # Process primary mailbox
    [array]$Stats = Get-ExoMailboxStatistics -Identity $Mbx.ExternalDirectoryObjectId 
    # Get the mailbox size in bytes and in as a formatted version in GB
    $MbxBytes = ($Stats.TotalItemSize.value.ToString().Split("(")[1].Split(" ")[0].Replace(",",""))  
    [float]$MbxSizeGB = [math]::Round($MbxBytes/1024MB,2) 
    $FolderStats = Get-ExoMailboxFolderStatistics -Identity $Mbx.ExternalDirectoryObjectId -IncludeOldestAndNewestItems | `
          Where-Object {$_.ItemsInFolder -gt 0 -and $_.Name -notin $UnwantedFolders} | Sort-Object Name
        ForEach ($Folder in $FolderStats) {
            [float]$FolderSize = [math]::Round($Folder.FolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",",""),2)
            # Format the folder size (in bytes) in KB, MB, or GB
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
                    User            = $Mbx.displayName
                    UPN             = $Mbx.UserPrincipalName
                    Folder          = $Folder.Name
                    Size            = $FolderSizeReport
                    Items           = $Folder.ItemsInFolder
                    'Oldest item'   = $OldestItemDate
                    'Newest item'   = $NewestItemDate
                    Path            = $Folder.FolderPath
                    Type            = $Folder.FolderType
                    Mailbox         = "Primary"
                }  
                $FolderReport.Add($FolderReportLine)       
            }      
        }
    If ($Mbx.ArchiveStatus -eq 'Active') {
        [array]$ArchiveStats = Get-ExoMailboxStatistics -Identity $Mbx.ExternalDirectoryObjectId -Archive | Sort-Object Name
        $ArchiveMbxBytes = ($ArchiveStats.TotalItemSize.value.ToString().Split("(")[1].Split(" ")[0].Replace(",",""))  
        [float]$ArchiveMbxSizeGB = [math]::Round($ArchiveMbxBytes/1024MB,2) 
        $ArchiveFolderStats = Get-ExoMailboxFolderStatistics -Identity $Mbx.ExternalDirectoryObjectId -Archive -IncludeOldestAndNewestItems | `
            Where-Object {$_.ItemsInFolder -gt 0 -and $_.Name -notin $UnwantedFolders} | Sort-Object Name
        ForEach ($Folder in $ArchiveFolderStats) {
            [float]$FolderSize = [math]::Round($Folder.FolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",",""),2)
            $FolderSizeReport = Format-FolderSize -InFolderSize $FolderSize
            If (($Folder.Name.Substring(0,1) -ne "{")) {
                $OldestItemDate = $null; $NewestItemDate = $null
                If ($Folder.OldestItemReceivedDate) {
                  $OldestItemDate = Get-Date $Folder.OldestItemReceivedDate -format 'dd-MMM-yyy HH:mm'
                }
                If ( $Folder.NewestItemReceivedDate) {
                    $NewestItemDate = Get-Date $Folder.NewestItemReceivedDate -format 'dd-MMM-yyy HH:mm'
                }
            $FolderReportLine = [PSCustomObject][Ordered]@{ 
                User            = $Mbx.displayName
                UPN             = $Mbx.UserPrincipalName
                Folder          = $Folder.Name
                Size            = $FolderSizeReport
                Items           = $Folder.ItemsInFolder
                'Oldest item'   = $OldestItemDate
                'Newest item'   = $NewestItemDate
                Path            = $Folder.FolderPath
                Type            = $Folder.FolderType
                Mailbox         = "Archive"
              }  
            $FolderReport.Add($FolderReportLine)       
          } 
        }     
    } Else {
        $ArchiveStats = $null
        $ArchiveMbxSizeGB = 0
    }

    # Get some other details for the user account
    $UserDetails = Get-User -Identity $Mbx.ExternalDirectoryObjectId
    $MbxReportLine = [PSCustomObject][Ordered]@{ 
        Mailbox                         = $Mbx.displayName
        UPN                             = $Mbx.UserPrincipalName
        'Primary mailbox size GB'       = ("{0} GB" -f $MbxSizeGB.toString())
        'Primary mailbox size bytes'    = $MbxBytes
        'Primary mailbox total items'   = $Stats.ItemCount
        'Archive status'                = $Mbx.ArchiveStatus
        'Archive mailbox size GB'       = ("{0} GB" -f $ArchiveMbxSizeGB.toString())
        'Archive mailbox size bytes'    = $ArchiveMbxBytes
        'Archive mailbox total items'   = $ArchiveStats.ItemCount
        'Mailbox type'                  = $Mbx.RecipientTypeDetails
        City                            = $UserDetails.City
        Country                         = $UserDetails.CountryOrRegion
        Department                      = $UserDetails.Department
        Office                          = $UserDetails.Office
        Title                           = $UserDetails.Title
    }  
    $MbxReport.Add($MbxReportLine)  
} # End processing mailboxes

# Uncomment these lines if you want to see the data captured for folders and mailboxes
# $FolderReport | Out-GridView
# $MbxReport | Out-GridView

$HtmlPrimarySeparator = "<p></h3>Folders in the <b>primary mailbox</b></h3></p>"
$HtmlArchiveSeparator = "<p></h3>Folders in the <b>archive mailbox</b></h3></p>"
$HtmlRecoverableItemsSeparator = "<p></h3><b>Recoverable Items</b> folders</h3></p>"

# Create the HTML report
$HtmlHead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 10pt;}
	   H1{font-size: 32px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 24px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 20px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.warn{background: #FFF275;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Exchange Online Mailbox Folders Report</h1></p>
           <p><h2><b>For the <b>" + $Orgname + "</b> tenant</b></h2></p>"

$HtmlBody = $null
ForEach ($Mbx in $Mailboxes) {
    # Find the folders etc. for each user that we will report
    $UserFolders = $null
    [array]$UserFolders = $FolderReport | Where-Object {$_.UPN -eq $Mbx.UserPrincipalName}
    [array]$PrimaryFolders = $UserFolders | Where-Object Mailbox -match 'Primary'
    [array]$ArchiveFolders = $UserFolders | Where-Object Mailbox -match 'Archive'
    [array]$MailboxStats = $MbxReport | Where-Object {$_.Mailbox -eq $Mbx.DisplayName}

    $PrimaryFolderHTML = $PrimaryFolders | Where-Object {$_.Type -notin $RIFolders} | `
        Select-Object Folder, Size, Items, 'Newest item', 'Oldest item', Path | ConvertTo-Html -Fragment
    $PrimaryFolderHTML = $HtmlPrimarySeparator + $PrimaryFolderHTML
    $PrimaryFolderRIHTML = $PrimaryFolders | Where-Object {$_.Type -in $RIFolders} | `
        Select-Object Folder, Size, Items, 'Newest item', 'Oldest item', Path | ConvertTo-Html -Fragment
    $PrimaryFolderRIHTML = $HtmlRecoverableItemsSeparator + $PrimaryFolderRIHTML

    If ($ArchiveFolders) {
        $ArchiveFolderHTML = $ArchiveFolders | Where-Object {$_.Type -notin $RIFolders} | `
            Select-Object Folder, Size, Items, 'Newest item', 'Oldest item', Path | ConvertTo-Html -Fragment
        $ArchiveFolderHTML = $HtmlArchiveSeparator + $ArchiveFolderHTML
        $ArchiveFolderRIHTML = $ArchiveFolders | Where-Object {$_.Type -in $RIFolders} | `
            Select-Object Folder, Size, Items, 'Newest item', 'Oldest item', Path | ConvertTo-Html -Fragment
        $ArchiveFolderRIHTML = $HtmlRecoverableItemsSeparator + $ArchiveFolderRIHTML
    }
    # Add mailbox statistics to the output
    If ($ArchiveFolders) {
        $MailboxStatsHTML = $MailboxStats | Select-Object 'Primary mailbox size GB', 'Primary mailbox total items', `
        'Archive status',  'Archive mailbox size GB', 'Archive mailbox total items' | ConvertTo-Html -Fragment
    } Else {
        $MailboxStatsHTML = $MailboxStats | Select-Object 'Primary mailbox size GB', 'Primary mailbox total items', `
        'Archive status' | ConvertTo-Html -Fragment
    }

    $MailboxHTMLHeader = ("<h3>Mailbox: <b>{0}</b>  ({1})</h3>" -f $Mbx.DisplayName, $Mbx.UserPrincipalName)
    $MailboxHTMLHeader = "<p>" + $MailboxHTMLHeader + $MailboxStatsHTML + "<p></p>"
    $MailboxHTML = $MailboxHTMLHeader + "<p>" + $PrimaryFolderHTML + "</p><p>" + $PrimaryFolderRIHTML + "</p>"
    If ($ArchiveFolders) {
        $MailboxHTML = $MailboxHTML + "<p>" + $ArchiveFolderHTML + "</p><p>" + $ArchiveFolderRIHTML + "</p>"
    }
    $HtmlBody = $HtmlBody + $MailboxHTML
}

$HtmlBody = $HtmlBody + "<p>Report created for: " + $OrgName + "</p>" +
                        "<p>Created: " + $RunDate + "<p>" 
                                  
$HtmlTail = "<p>Exchange Online Mailbox Folder Report<b> " + $Version + "</b></p>"	

$HtmlReport = $Htmlhead + $Htmlbody + $Htmltail
$HtmlReport | Out-File $HtmlReportFile  -Encoding UTF8
Write-Host ("Report generated and available at {0}" -f $HtmlReportFile)
$MbxReport | Export-CSV -NoTypeInformation $MailboxStatsCSV
Write-Host ("A CSV file with mailbox statistics is available in {0}" -f $MailboxStatsCSV)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
