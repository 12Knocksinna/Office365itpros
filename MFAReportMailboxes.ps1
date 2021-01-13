# MFAReportMailboxes.ps1
# List mailboxes and the last time the Mailbox Folder Assistant processed each mailbox
# https://github.com/12Knocksinna/Office365itpros/blob/master/MFAReportMailboxes.ps1
$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
Write-Host "Fetching details of user mailboxes..."
ForEach ($M in $Mbx) {
   $LastProcessed = $Null
   Write-Host "Processing" $M.DisplayName
   $Log = Export-MailboxDiagnosticLogs -Identity $M.Alias -ExtendedProperties
   $xml = [xml]($Log.MailboxLog)  
   $LastProcessed = ($xml.Properties.MailboxTable.Property | ? {$_.Name -like "*ELCLastSuccessTimestamp*"}).Value   
   $ItemsDeleted  = $xml.Properties.MailboxTable.Property | ? {$_.Name -like "*ElcLastRunDeletedFromRootItemCount*"}
   If ($LastProcessed -eq $Null) {
      $LastProcessed = "Not processed"}
   $ReportLine = [PSCustomObject][Ordered]@{
           User          = $M.DisplayName
           LastProcessed = $LastProcessed
           ItemsDeleted  = $ItemsDeleted.Value}      
    $Report.Add($ReportLine)
  }
$Report | Select User, LastProcessed, ItemsDeleted
$Report | Out-GridView
