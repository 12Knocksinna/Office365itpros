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

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
