# MailTrafficStatistics.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/MailTrafficStatistics.PS1
# Illustration of how to use the data produced by the Get-MailTrafficSummaryReport cmdlet to create a useful per-mailbox report
# of messages sent and received.
# Uses the Exchange Online PowerShell module

CLS
# Check that we are connected to Exchange Online
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}

$StartDate = (Get-Date).AddDays(-92); $EndDate = (Get-Date).AddDays(+1)
[array]$SenderData = Get-MailTrafficSummaryReport -Category TopMailSender -StartDate $StartDate -EndDate $EndDate | Select-Object C1, C2 
[array]$RecipientData = Get-MailTrafficSummaryReport -Category TopMailRecipient -StartDate $StartDate -EndDate $EndDate | Select-Object C1, C2 
$MbxReport = [System.Collections.Generic.List[Object]]::new()

$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited
ForEach ($M in $Mbx) {
Write-Host "Processing" $M.DisplayName   
# Check each email proxy address to see if it was used to send email
[int]$TotalSentMessages = 0 ; [int]$Messages = 0
ForEach ($A in $M.EmailAddresses) {
      If ($A.Substring(0,4) -eq "smtp") {
         $Messages = $SenderData | ? {$_.C1 -eq $A.Split(":")[1] } | Select -ExpandProperty C2
         # Write-Host "Messages found for" $A " " $Messages
         $TotalSentMessages = ($TotalSentMessages + $Messages) }
}
# Check each email proxy address to see if it was used to receive email
[int]$TotalReceivedMessages = 0 ; [int]$Messages = 0
ForEach ($A in $M.EmailAddresses) {
      If ($A.Substring(0,4) -eq "smtp") {
         $Messages = $RecipientData | ? {$_.C1 -eq $A.Split(":")[1] } | Select -ExpandProperty C2
         Write-Host "Messages found for" $A " " $Messages
         $TotalReceivedMessages = ($TotalReceivedMessages + $Messages) }
}
     $ReportLine = [PSCustomObject] @{
        User                = $M.DisplayName
        Address             = $M.UserPrincipalName
        "Sent messages"     = $TotalSentMessages
        "Received messages" = $TotalReceivedMessages }    
      $MbxReport.Add($ReportLine) 
}
$MbxReport | Out-GridView


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
