# FindAccountsWithForwarding.PS1
# Script to find accounts with mail forwarding addresses enanled or with rules to forward messages
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindAccountsWithForwarding.PS1

# Check that the right modules are loaded
$Modules = Get-Module
If ("ExchangeOnlineManagement" -notin  $Modules.Name) {Write-Host "Please connect to Exchange Online Management  before continuing...";break}

$Mbx = (Get-ExoMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -Properties ForwardingSmtpAddress -ResultSize Unlimited)
Write-Host $Mbx.Count "user and shared mailboxes found. Now checking..."
$Report = [System.Collections.Generic.List[Object]]::new()
$Count = 0; CLS; $MbxNumber = 0
ForEach ($M in $Mbx) {
    $MbxNumber++
    $ProgressBar = "Checking mailbox " + $M.DisplayName + " (" + $MbxNumber + " of " + $Mbx.Count + ")" 
    Write-Progress -Activity "Looking for forwarding settings and inbox rules" -Status $ProgressBar -PercentComplete ($MbxNumber/$Mbx.Count*100)
    Write-Host "Processing" $M.DisplayName
    $Rule = $Null
    If ($M.ForwardingSmtpAddress -ne $Null) { # Mailbox has a forwarding address
         $ReportLine = [PSCustomObject]@{
            Mailbox           = $M.DisplayName
            UPN               = $M.UserPrincipalName
            "Mailbox type"    = $M.RecipientTypeDetails
            ForwardingAddress = $M.ForwardingSmtpAddress.Split(":")[1]
            InboxRule         = "N/A" 
            "Rule Removed"    = "N/A" 
            Enabled           = "N/A"}
       $Report.Add($ReportLine)
    } 
    $InboxRules = (Get-InboxRule -Mailbox $M.Alias | ? {$_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo})
    If ($InboxRules -ne $Null) {
       Write-Host "Processing inbox rules"
       ForEach ($Rule in $InboxRules) {
          $Ex = $Null
          $ForwardTo = @()
          $ForwardTo = ($Rule.ForwardTo | ? { ($_ -Match "SMTP") -or ($_ -Match "EX:") } )
          $ForwardTo += ($Rule.ForwardAsAttachmentTo | ? {($_ -Match "SMTP") -or ($_ -Match "EX:")})
          $ForwardTo += ($Rule.RedirectTo | ? {($_ -Match "SMTP") -or ($_ -Match "EX:")})
          If ($ForwardTo.Count -gt 0) {
             ForEach ($Recipient in $ForwardTo) {
                If ($Recipient -Match "EX:") {
                   # Recipient known in Exchange directory
                   $Ex = (Get-Recipient -Identity ($Recipient-Split "Ex:")[1].trim("]}")) 
                   $EmailAddress = $Ex.PrimarySmtpAddress }
                Else  {
                  # Simple SMTP address
                   $EmailAddress = ($Recipient -Split "SMTP:")[1].Trim("]") 
                   $Ex = (Get-Recipient -Identity $EmailAddress) }
             }
      
             Write-Host $M.RecipientTypeDetails $M.DisplayName "has a rule to forward email to" $EmailAddress -ForegroundColor Red
             # Remove the rule if the address is unknown to the directory
             If ($Ex -eq $Null) {
                 Remove-InboxRule -Identity $Rule.Identity -Confirm:$False; $RuleRemoved = "Yes"
                 Write-Host "Rule" $Rule.Name "removed from mailbox!" }
             Else {
                 Write-Host "Destination is known to the tenant directory. Please remove" $Rule.Name "manually if necessary"; $RuleRemoved = "No" }

             $ReportLine = [PSCustomObject]@{
                Mailbox           = $M.DisplayName
                UPN               = $M.UserPrincipalName
                "Mailbox type"    = $M.RecipientTypeDetails
                ForwardingAddress = $EmailAddress
                InboxRule         = $Rule.Name 
                "Rule Removed"    = $RuleRemoved 
                Enabled           = $Rule.Enabled}
             $Report.Add($ReportLine) }
       }
     }
}
[array]$InboxRulesFound = $Report |?{$_.InboxRule -ne "N/A"}
[array]$MailForwarding  = $Report |?{$_.InboxRule -eq "N/A"}
$MailboxesWithRules = $InboxRulesFound.Mailbox -join ", "
$MailboxesForwarding =  $Mailforwarding.Mailbox -join ", "
Write-Host ("{0} mailboxes found with forwarding addresses: {1}; {2} mailboxes found with forwarding inbox rules: {3}" -f $MailForwarding.Count, $MailboxesForwarding, $InboxRulesFound.Count, $MailboxesWithRules)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
