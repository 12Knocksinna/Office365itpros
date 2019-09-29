# ReportMailboxQuotaUsed.Ps1
# Script to report mailbox quota assigned and percentage used and to signal warning if quota used exceeds set threshold
# Set threshold % of quota to use as warning level
$Threshold = 85
# Get all user mailboxes
Cls
Write-Host "Finding mailboxes..."
$Mbx = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Select DisplayName, ProhibitSendReceiveQuota, DistinguishedName
$Report = @() 
ForEach ($M in $Mbx) {
    # Find current usage
    Write-Host "Processing" $M.DisplayName
    $Mailbox = $M.DisplayName
    $ErrorText = $Null
    $MbxStats = Get-MailboxStatistics $M.DistinguishedName | Select ItemCount, TotalItemSize
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
   $Report += $ReportLine 
} 
# Export to CSV
$Report | Sort Mailbox | Export-csv -NoTypeInformation MailboxQuotaReport.csv
