# ReportQuarantinedMessages.PS1
# Showing how to download details of quarantined messages, do some analysis, and create a CSV file that can be edited
# and then used to release good messages
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportQuarantinedMessages.PS1
# 
$ModuleCheck = Get-Module -Name ExchangeOnlineManagement
If ($ModuleCheck -eq $Null) {
     Write-Host "Your PowerShell session is not connected to Exchange Online."
     Write-Host "Please connect to Exchange Online using an administrative account and retry."; Break }

Write-Host "Finding messages in quarantine"
# Check https://docs.microsoft.com/en-us/powershell/module/exchange/get-quarantinemessage?view=exchange-ps for other 
# parameters that can be used to refine the set of quarantined messages
$QMessages = Get-QuarantineMessage 
$Report = [System.Collections.Generic.List[Object]]::new(); $Now = Get-Date

# Extract the data we want to report abouit each quarantined message
ForEach ($Message in $QMessages) {
   
   $RemainingTime = (New-TimeSpan -Start $Now -End $Message.Expires)
   $Remaining = $RemainingTime.Days.toString() + " days " + $RemainingTime.Hours.toString() + " hours"
   [String]$Recipient = $Null; $i = 0
   ForEach ($Address in $Message.RecipientAddress) {
        If ($i -eq 0) {
           $i++
           $Recipient = $Address} 
        Else 
           {$Recipient = "; " + $Address }
   }
   $ReportLine = [PSCustomObject]@{  #Update with details of what we have done
           Identity         = $Message.Identity
           Received         = Get-Date($Message.ReceivedTime) -format g
           Recipient        = $Recipient
           Sender           = $Message.SenderAddress
           Subject          = $Message.Subject
           SenderDomain     = $Message.SenderAddress.Split("@")[1]
           Type             = $Message.QuarantineTypes
           Expires          = Get-Date($Message.Expires) -format g
           "Time Remaining" = $Remaining } 
    $Report.Add($ReportLine) 
}

CLS
Write-Host "Type of Quarantined messages"
$Report | Group Type | Sort Count -Descending | Format-Table Name, Count

Write-Host "Messages quarantined per recipient address"
$Report | Group Recipient | Sort Count -Descending | Format-Table Name, Count

Write-Host "Problem domains"
$Report | Group SenderDomain |Sort Count -Descending | Format-Table Name, Count

Write-Host "High confidence Phishing Messages"
$Report | ? {$_.Type -eq "HighConfPhish"} | Format-Table Received, Recipient, Sender, Subject

$Report | Export-CSV -NoTypeInformation c:\Temp\QuarantinedMessages.CSV
# After editing the list, you can release the messages with:
# Import-CSV c:\temp\QuarantinedMessages.csv | Release-QuarantineMessage -ReleaseToAll 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
