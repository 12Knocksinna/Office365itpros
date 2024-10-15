# Purge-ComplianceSearchItems.PS1
# An example script to show how to purge items found by a compliance search
# V1.0 15-Oct-2024

# The script will ask some questions to perform a compliance search against one or all mailboxes, do the search, report what it finds
# and ask the administrator if they want to proceed with purging the items found by the search. The script will then purge the items.

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Purge-ComplianceSearchItems.PS1

function Get-RecoverableItemsFolderId ([string]$MailboxId) {
    # Function to return folder identifiers for the Recoverable Items folders in a mailbox
    # The types of folders we want to exclude from the search
    $FolderExclusions = [System.Collections.Generic.List[Object]]::new()
    [array]$RecoverableFolders = 'RecoverableItemsVersions', 'RecoverableItemsDeletions', 'RecoverableItemsPurges'
    
    # Process Recoverable Items folders in the primary mailbox
    $Folders = Get-ExoMailboxFolderStatistics $MailboxId
    Foreach ($Folder in $Folders) {
        If ($Folder.FolderType -in $RecoverableFolders) {           
            $encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
            $nibbler= $encoding.GetBytes("0123456789ABCDEF")
            $folderIdBytes = [Convert]::FromBase64String($Folder.FolderId)
            $indexIdBytes = New-Object byte[] 48
            $indexIdIdx=0
            $folderIdBytes | Select-Object -skip 23 -First 24 | ForEach-Object{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
            $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))"
            $ReportLine = [PSCustomObject][Ordered]@{ 
                FolderPath = $Folder.FolderPath
                FolderQuery = $folderQuery
            }
            $FolderExclusions.Add($ReportLine)
        }
    }
    
    # Process Recoverable Items folders in the archive mailbox (if enabled)
    If ((Get-Mailbox -Identity $MailboxId).ArchiveGuid) {
        $Folders = Get-ExoMailboxFolderStatistics $MailboxId
        Foreach ($Folder in $Folders) {
            If ($Folder.FolderType -in $RecoverableFolders) {           
                $encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
                $nibbler= $encoding.GetBytes("0123456789ABCDEF")
                $folderIdBytes = [Convert]::FromBase64String($Folder.FolderId)
                $indexIdBytes = New-Object byte[] 48
                $indexIdIdx=0
                $folderIdBytes | Select-Object -skip 23 -First 24 | ForEach-Object{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
                $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))"
                $ReportLine = [PSCustomObject][Ordered]@{ 
                    FolderPath = $Folder.FolderPath
                    FolderQuery = $folderQuery
                }
                $FolderExclusions.Add($ReportLine)
            }
        }
    }
    
        Return $FolderExclusions
}

function Get-EmailDetailsFromPreviewResults {
    param (
        [string]$Line
    )

    # function to extract email details from the format used by the preview results

    $Pattern = "Location: (?<Location>.*?); Sender: (?<Sender>.*?); Subject: (?<Subject>.*?); Type: .*?; Size: .*?; Received Time: (?<ReceivedTime>.*?); Data Link: .*"
    if ($Line -match $Pattern) {
        return [PSCustomObject]@{
            Location     = $matches['Location']
            Sender       = $matches['Sender']
            Subject      = $matches['Subject']
            ReceivedTime = $matches['ReceivedTime']
        }
    } else {
        Write-Host "Line does not match the expected format."
        return $null
    }
}

function Get-ComplianceSearchPreview {
    Param (
        [string]$SearchName
    )

    # function to return the set of preview items for a compliance search

    Write-Host "Finding preview items for review..."
    $SearchNamePreview = ("{0}_Preview" -f $SearchName)
    Remove-ComplianceSearchAction -Identity $SearchNamePreview -Confirm:$False -ErrorAction SilentlyContinue | Out-Null
    New-ComplianceSearchAction -SearchName $SearchName -Preview -ErrorAction SilentlyContinue -Confirm:$false | Out-Null

    $PreviewState = Get-ComplianceSearchAction -Identity $SearchNamePreview

    If ($null -ne $PreviewState){
        While ($PreviewState.Status -ne "Completed") {
            Start-Sleep -Seconds 5
            $PreviewState = Get-ComplianceSearchAction -Identity $SearchNamePreview
        }
    }
    If ($PreviewState.Results) {
        Write-Host "Formatting preview results..."
        $PreviewItems = $PreviewState.Results -replace "{" -replace "}" -split ".eml," -replace "`n"
        ForEach ($Item in $PreviewItems) {
            $PreviewItem = Get-EmailDetailsFromPreviewResults -Line $Item
            $ReportLine = [PSCustomObject][Ordered]@{ 
                Recipient       = $PreviewItem.Location
                Sender          = $PreviewItem.Sender
                Subject         = $PreviewItem.Subject
                'Received Date' = $PreviewItem.ReceivedTime
            }
            $PreviewReport.Add($ReportLine)
        }
        $PreviewReport = $PreviewReport | Sort-Object {$_.'Received Date' -as [datetime]} -Descending
    }
        
        Return $PreviewReport
   
}

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp -ShowBanner:$false
}

Write-Host "Connecting to Security & Compliance Center endpoint..."
Connect-IPPSSession -ShowBanner:$false

# Start of processing - clear the screen and ask which mailbox to search, what to look for, and so on.
Clear-Host
$SearchAllMailboxes = $false
[string]$Mbx = Read-Host "Enter the mailbox to search or type All to search all mailboxes"
If (Get-Mailbox -Identity $Mbx -ErrorAction SilentlyContinue) {
    $MailboxId = Get-Mailbox -Identity $Mbx
    $MailboxesToSearch = $MailboxId.UserPrincipalName
    $MailboxDisplayName = ("the {0} mailbox" -f $MailboxId.DisplayName)
} ElseIf ($Mbx.ToLower() -eq "all") {
    # Checking all mailboxes...
    $MailboxId = $null
    $SearchAllMailboxes = $true
    $RecoverableItemsOption = "N"
    $MailboxesToSearch = "All"
    $MailboxDisplayName = "All mailboxes"
} Else {    
    Write-Host "Mailbox $Mbx not found" -ForegroundColor Red
    Break
}

[string]$SearchString = Read-Host "Enter the message subject to search for"

# If we search all mailboxes, we won't exclude Recoverable Items, but if the flag is false, ask if we should exclude the RI folders
If ($SearchAllMailboxes -eq $false) {
    $RecoverableItemsOption = Read-Host "Exclude Recoverable Items folders [Y] Yes [N] No (default is `"N`")"
    If ($RecoverableItemsOption.ToUpper() -eq "Y") {
        Write-Host "Recoverable Items folders will be excluded..."
    }
}

Write-Host "Computing query criteria"
# start and end date for the query. The default search goes back 60 days from today. Obviously, it can be any value you want.
$StartDate =  Get-Date (Get-Date).AddDays(-60) -format yyyy-MM-dd
$EndDate = Get-Date -format yyyy-MM-dd

# Construct the KQL query for the compliance search
$KQLQuery = ("(c:c)(subjecttitle: '{0}')" -f $SearchString)
$DateRange = ("(sent={0}..{1})" -f $StartDate, $EndDate)

# add the folder identifiers for the Recoverable Items folders to the query if the option is set
$FormattedFolderIds = $null
If ($RecoverableItemsOption.ToUpper() -eq "Y") {
    Write-Host "Fetching folder identifiers to exclude Recoverable Items folders from the search"
    $FoldersToExclude = Get-RecoverableItemsFolderId -MailboxId $MailboxId.ExternalDirectoryObjectId
    $FormattedFolderIds = " NOT ((" + ($FoldersToExclude.FolderQuery -join ") OR (") + "))"
}
$KQLQuery = ("{0}{1}{2}" -f $KQLQuery, $DateRange, $FormattedFolderIds)

# Tell the administrator what we are going to do
Write-Host ("Compliance search will run against {0} using query {1}" -f $MailboxDisplayName, $KQLQuery)
$SearchName = ("Compliance Purge {0} {1}" -f $MailboxesToSearch, $EndDate)

# remove any old search of the same name and create and start a new search
Remove-ComplianceSearch -Identity $SearchName -Confirm:$False -ErrorAction SilentlyContinue
Write-Host ("Creating and running new compliance search ({0})" -f $SearchName) -ForegroundColor Yellow 
New-ComplianceSearch -Name $SearchName -ExchangeLocation $MailboxesToSearch -ContentMatchQuery $KQLQuery -Description 'Compliance Search Test' | Out-Null
Start-ComplianceSearch -Identity $SearchName

Write-Host "Waiting for compliance search to finish..."
Do { 
    $ComplianceSearch = Get-ComplianceSearch $SearchName -ErrorAction SilentlyContinue
    $Status = $ComplianceSearch.Status 
    Start-Sleep -Seconds 3 
} While ($Status -ne 'Completed') 
[int]$Count = 0 

[int]$ItemsFound = $ComplianceSearch.Items
If (!($ItemsFound)) { 
    Write-Host "No items found in compliance search" -ForegroundColor Red 
    break 
}

# Use regex to find all instances of the search results where item count is greater than zero
$LocationsWithItemCount = [regex]::Matches($ComplianceSearch.SuccessResults, "Item count: (\d+)")
[array]$Locations = $LocationsWithItemCount | Where-Object { $_.Groups[1].Value -gt 0 }

# Use regex to extract email addresses, item counts, and total sizes from the search results
$LocationsWithEmail = [regex]::Matches($ComplianceSearch.SuccessResults, "Location: (\S+@\S+\.\S+), Item count: (\d+), Total size: (\d+)")
# Extract email addresses and item counts where item count is greater than 0
[array]$LocationsWithItems = $null
foreach ($Match in $LocationsWithEmail) {
    $Email = $match.Groups[1].Value
    $ItemCount = [int]$match.Groups[2].Value
    if ($ItemCount -gt 0) {
        $LocationsWithItems += [PSCustomObject]@{
            Email = $email
            ItemCount = $itemCount
        }
    }
}

# Figure out how many loops might be needed to remove all items
[int]$LocationsGT10 = 0; [int]$HighestValue = 0
If ($LocationsWithItems.Count -eq 1) {
    $HighestLocation = $LocationsWithItems[0].Email
    $HighestValue = $LocationsWithItems[0].ItemCount
} Else {
    ForEach ($Item in $LocationsWithItems) {
       $ItemCount = $Item.ItemCount
       If ($ItemCount -gt 10) {
            $LocationsGT10++
       }
       If ($ItemCount -gt $HighestValue) {
            $HighestValue = $ItemCount
            $HighestLocation = $Item.Email
       }
    }
}

$LoopsNeeded = [math]::ceiling($HighestValue/10)
$TotalResults = 0
# Let the administrator know what we have found 
Write-Host ("Compliance search completed. {0} items found in {1} locations. {2} iterations are required to remove these items. The mailbox with most items is {3} with {4}." -f $ItemsFound, $Locations.Count, $LoopsNeeded, $HighestLocation, $HighestValue) -ForeGroundColor Yellow
Write-Host ""
If ($LocationsWithItems.count -gt 1) {
    Write-Host "Locations with items found by the search" -ForegroundColor Yellow
    Write-Host "----------------------------------------" 
    $LocationsWithItems | Sort-Object -Property ItemCount -Descending | Format-Table -AutoSize
    Write-Host ""
}

# Section to find preview and show some times to user to check
$Global:PreviewReport = [System.Collections.Generic.List[Object]]::new()
$PreviewReport = Get-ComplianceSearchPreview -SearchName $SearchName

If ($PreviewReport) {
    Write-Host "Preview of items found by the search" -ForegroundColor Yellow
    Write-Host "------------------------------------" 
    $PreviewReport | Select-Object -First 10 | Format-Table -AutoSize
    Write-Host ""
} Else {
    Write-Host "No preview items found" -ForegroundColor Red
}

$OKToProceed = Read-Host "Do you want to proceed with the purge? [Y] Yes [N] No"
If ($OKToProceed.ToUpper() -eq "N") {
    Write-Host "Exiting without purging items" -ForegroundColor Red
    Break
}

Write-Host "Setting up to purge items..."
$SearchNameAction = $ComplianceSearch.Name + "_Purge"

While ($Count -lt $LoopsNeeded) {
    $Count++ 
    # Remove old purge action
    Write-Host "Removing previous compliance search action (if one exists)" -ForegroundColor Yellow 
    Remove-ComplianceSearchAction $SearchNameAction -Confirm:$False -ErrorAction SilentlyContinue
    Write-Host ("Submitting search purge action for run {0}" -f $Count)
    New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType SoftDelete -Confirm:$False | Out-Null
    Write-Host "Waiting for compliance purge action to finish..."
    Do { 
        $Action = Get-ComplianceSearchAction $SearchNameAction 
        $Status = $Action.Status    
        Start-Sleep -Seconds 3
    } While ($Status -ne 'Completed') 
    # Figure out what happened during the purge action
    $Action = Get-ComplianceSearchAction $SearchNameAction 
    
    # Parse the purge action results to extract the results of all locations where items were found
    If ($Action.Results -match "Item count: (\d+);") {
        $Results = [int]$matches[1]
    } Else {
        Write-Host "No items were removed by the purge"
    }
    $TotalResults = $TotalResults + $Results
    Write-Host ("Loop {0} purged {1} items" -f $Count, $Results) -ForegroundColor Green
    $Action = $null; $Status = $null; $Results = $null
}

Write-Host ("Purge completed. {0} items removed from {1} locations" -f $TotalResults, $Locations.Count) -ForegroundColor Green


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
