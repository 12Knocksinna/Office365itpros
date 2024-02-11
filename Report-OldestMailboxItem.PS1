# Feport-OldestMailboxItems.PS1
# Find mailboxes larger than 50 GB and report the folder name where the oldest item is found
#

Connect-ExchangeOnline -SkipLoadingCmdletHelp

$Report = [System.Collections.Generic.List[Object]]::new()
[array]$Mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Sort-Object DisplayName
ForEach ($Mbx in $Mailboxes) {
  $Stats = Get-MailboxStatistics -Identity $Mbx.Alias
  [float]$MbxSizeGB = [math]::Round(($Stats.TotalItemSize.value.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1024MB),2)
  If ($MbxSizeGB -gt 50 ) {
        $FolderStats = Get-MailboxFolderStatistics -Identity $Mbx.Alias -IncludeOldestAndNewestItems | `
             Select-Object Name, OldestItemReceivedDate | Sort-Object OldestItemReceivedDate | `
            Where-Object OldestItemReceiveddate -ne $null | Select-Object -First 1
        $ReportLine = [PSCustomObject][Ordered]@{ 
            Mailbox         = $Mbx.displayName
            'Size GB'       = $MbxSizeGB
            Folder          = $FolderStats.name
            'Oldest item'   = $FolderStats.OldestItemReceivedDate
        }  
        $Report.Add($ReportLine)  
  }
 }

 $Report | Out-GridView

 # An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.