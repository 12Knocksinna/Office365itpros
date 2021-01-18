# GetLastActiveTimeMailboxes.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/GetLastActiveTimeMailboxes.PS1
# Report last active time for mailboxes
# Example script from Chapter 5 of Office 365 for IT Pros

$Mbx = (Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Select DisplayName, DistinguishedName) 
  $Report = [System.Collections.Generic.List[Object]]::new()
  ForEach ($M in $Mbx) { 
   Write-Host "Processing" $M.DisplayName 
   $Log = Export-MailboxDiagnosticLogs -Identity $M.DistinguishedName -ExtendedProperties 
   $xml = [xml]($Log.MailboxLog) 
   $LastEmail = ($xml.Properties.MailboxTable.Property | ? {$_.Name -like "LastEmailTimeCurrentValue"}).Value
   $LastCalendar = ($xml.Properties.MailboxTable.Property | ? {$_.Name -like "LastCalendarTimeCurrentValue"}).Value
   $LastContacts = ($xml.Properties.MailboxTable.Property | ? {$_.Name -like "LastContactsTimeCurrentValue"}).Value
   $LastFile = ($xml.Properties.MailboxTable.Property | ? {$_.Name -like "LastFileTimeCurrentValue"}).Value
   $Stats = (Get-MailboxStatistics -Identity $M.DistinguishedName) 
   $MbxSize = ($Stats.TotalItemSize.Value.ToString()).Split("(")[0] 
   $ReportLine = [PSCustomObject]@{ 
        Mailbox = $M.DisplayName 
        Items = $Stats.ItemCount 
        Size = $MbxSize 
        LastLogon = Get-Date($Stats.LastLogonTime) -Format g
        LastActive = Get-Date($Stats.LastInteractionTime) -Format g
        LastEmail  = Get-Date($LastEmail) -Format g
        LastCalendar = Get-Date($LastCalendar) -Format g
        LastContacts = Get-Date($LastContacts) -Format g
        LastFile = Get-Date($LastFile) -Format g} 
   $Report.Add($ReportLine)} 
$Report | Export-csv -NoTypeInformation Users.csv
$Report | Sort {$_.LastLogon -as [datetime] -descending | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
