# ReportMailboxQuotaUsed.Ps1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportMailboxQuotaUsed.Ps1
# Script to report mailbox quota assigned and percentage used and to signal warning if quota used exceeds set threshold
# Set threshold % of quota to use as warning level
$Threshold = 85
# Get all user mailboxes
Cls
Write-Host "Finding mailboxes..."
$Mbx = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox -Properties ProhibitSendReceiveQuota | Select DisplayName, ProhibitSendReceiveQuota, DistinguishedName
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
ForEach ($M in $Mbx) {
    # Find current usage
    Write-Host "Processing" $M.DisplayName
    $Mailbox = $M.DisplayName
    $ErrorText = $Null
    $MbxStats = Get-ExoMailboxStatistics $M.DistinguishedName | Select ItemCount, TotalItemSize
    # Return byte count of quota used
    [INT64]$QuotaUsed = [convert]::ToInt64(((($MbxStats.TotalItemSize.ToString().split("(")[-1]).split(")")[0]).split(" ")[0]-replace '[,]',''))
    # Byte count for mailbox quota
    [INT64]$MbxQuota = [convert]::ToInt64(((($M.ProhibitSendReceiveQuota.ToString().split("(")[-1]).split(")")[0]).split(" ")[0]-replace '[,]',''))
    $MbxQuotaGB = [math]::Round(($MbxQuota/1GB),2)
    $QuotaPercentUsed = [math]::Round(($QuotaUsed/$MbxQuota)*100,2)
    $QuotaUsedGB = [math]::Round(($QuotaUsed/1GB),2)
    If ($QuotaPercentUsed -gt $Threshold) {
       Write-Host $M.DisplayName "current mailbox use is above threshold at" $QuotaPercentUsed -Foregroundcolor Red
       $ErrorText = "Mailbox quota over threshold" }
    # Generate report line for the mailbox
    $ReportLine = [PSCustomObject][Ordered]@{ 
        Mailbox          = $M.DisplayName 
        MbxQuotaGB       = $MbxQuotaGB
        Items            = $MbxStats.ItemCount
        MbxSizeGB        = $QuotaUsedGB
        QuotaPercentUsed = $QuotaPercentUsed
        ErrorText        = $ErrorText} 
   $Report.Add($ReportLine)
} 
# Export to CSV
$Report | Sort Mailbox | Export-csv -NoTypeInformation MailboxQuotaReport.csv

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
