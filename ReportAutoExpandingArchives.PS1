# ReportAutoExpandingArchives.PS1
# A script to find mailboxes with auto-expanding archives and report how near each mailbox is to the 1.5 TB limit imposed by Microsoft, the growth rate per day, 
# and how many days it will take a mailbox to hit the limit
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportAutoExpandingArchives.PS1
# V1.0 29-Sept-2021
$Modules = Get-Module
If ("ExchangeOnlineManagement" -notin  $Modules.Name) {Write-Host "Please connect to Exchange Online Management before continuing...";break}
If (((Get-OrganizationConfig).AutoExpandingArchiveEnabled) -ne $True) { 
     Write-Host "Expanding archives are not enabled in this tenant" ; break}

# Find set of archive-enabled mailboxes 
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails SharedMailbox, UserMailbox -Filter {ArchiveStatus -ne $Null} -ResultSize Unlimited -Properties ArchiveQuota, ArchiveStatus, AutoExpandingArchiveEnabled, RecipientTypeDetails, ArchiveName, WhenCreated
If ($Mbx -eq 0) { Write-Host "No archive-enabled mailboxes found" ; break }

# Number of bytes in 1.5 TB and 90% of that figure to generate warning
[long]$TBBytes = "1649267441664"
[long]$TBBytesWarning = $TBBytes * 0.9  
$Now = Get-Date
# Reduce to the set of mailboxes with expanding archives
[array]$ExMbx = $Mbx | Where-Object {$_.AutoExpandingArchiveEnabled -eq $True}
If ($ExMbx -eq 0) { Write-Host "No mailboxes with expandable archives found" ; break }

# Process each mailbox and identify how many days they can expand before hitting the limit
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $ExMbx) {
   $Status = $Null
   Write-Host "Processing mailbox" $M.DisplayName
   [int]$DaysSinceCreation = ((New-TimeSpan -Start ($M.WhenCreated) -End ($Now)).Days)
   $Stats = Get-ExoMailboxStatistics -Archive -Identity $M.UserPrincipalName
   [string]$ArchiveSize = $Stats.TotalItemSize.Value
   [string]$DeletedArchiveItems = $Stats.TotalDeletedItemSize.Value 
   [long]$BytesInArchive = $Stats.TotalItemSize.Value.ToBytes()
   [long]$BytesInRecoverableItems = $Stats.TotalDeletedItemSize.Value.ToBytes()
   [long]$TotalBytesInArchive = $BytesInArchive + $BytesInRecoverableItems
   # Check if archive size is within 10% of the 1.5 TB limit - the size that counts is the combination of Recoverable Items and normal folders
   If ($TotalBytesInArchive -ge $TBBytesWarning) 
       { Write-Host ("Archive size {0} for {1} is within 10% of 1.5 TB limit" -f $ArchiveSize, $M.DisplayName ) 
         $Status = "Archive within 10% of 1.5 TB limit" }
   [long]$BytesPerDay = $TotalBytesInArchive/$DaysSinceCreation
   [long]$NumberDaysLeft = (($TBBytes - $TotalBytesInArchive)/$BytesPerDay)
   $BytesPerDayMB = $BytesPerDay/1MB
   $GrowthRateDay = [math]::Round($BytesPerDayMB,4)
   $TotalArchiveSizeGB = [math]::Round(($TotalBytesInArchive/1GB),2) 
   
   $ReportLine = [PSCustomObject][Ordered]@{  
       Mailbox                   = $M.DisplayName
       UPN                       = $M.UserPrincipalName
       Created                   = $M.WhenCreated
       Days                      = $DaysSinceCreation
       Type                      = $M.RecipientTypeDetails
       "Archive Quota"           = $M.ArchiveQuota.Split("(")[0] 
       "Archive Status"          = $M.ArchiveStatus
       "Archive Size"            = $ArchiveSize.Split("(")[0] 
       "Archive Items"           = $Stats.ItemCount
       "Deleted Archive Items Size" = $DeletedArchiveItems.Split("(")[0] 
       "Deleted Items"           = $Stats.DeletedItemCount
       "Total Archive Size (GB)" = $TotalArchiveSizeGB
       "Daily Growth Rate (MB)"  = $GrowthRateDay
       "Days Left to Limit"      = $NumberDaysLeft
       Status                    = $Status   
    }
    $Report.Add($ReportLine) 
} #End ForEach

Write-Host ("{0} mailboxes processed" -f $Mbx.count)
$Report | Select Mailbox, Type, "Archive Size", "Deleted Archive Items Size", "Total Archive Size (GB)", "Daily Growth Rate (MB)",  "Days Left to Limit"  | Out-GridView 
$Report | Export-CSV -NoTypeInformation c:\temp\ArchiveMailboxes.csv

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
