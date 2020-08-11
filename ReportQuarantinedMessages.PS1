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
