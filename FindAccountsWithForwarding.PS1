! FindAccountsWithForwarding.PS1
! Script to find accounts with mail forwarding addresses enanled or with rules to forward messages

$Modules = Get-Module
If ("ExchangeOnlineManagement" -notin  $Modules.Name) {Write-Host "Please connect to Exchange Online Management  before continuing...";break}
# Date to Check for new accounts - we use the last 7 days here, but that's easily changable.

Write-Host "Finding mailboxes to check..."
$Mbx = (Get-ExoMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -Properties ForwardingSmtpAddress -ResultSize Unlimited)
Write-Host $Mbx.Count "user and shared mailboxes found. Now checking..."
$NumberMbxWithRules = 0; $NumberForwards = 0
ForEach ($M in $Mbx) {
    Write-Host "Processing" $M.DisplayName
    $Rule = $Null
    If ($M.ForwardingSmtpAddress -ne $Null) { # Mailbox has a forwarding address
       $NumberForwards++
       Write-Host $M.DisplayName "is forwarding email to" $M.ForwardingSmtpAddress.Split(":")[1] } 
    $InboxRules = (Get-InboxRule -Mailbox $M.Alias | ? {$_.ForwardTo -or $_.ForwardAsAttachmentTo })
    If ($InboxRules -ne $Null) {
       Write-Host "Processing inbox rules"
       ForEach ($Rule in $InboxRules) {
          $Ex = $Null
          $ForwardTo = @()
          $ForwardTo = ($Rule.ForwardTo | ? { ($_ -Match "SMTP") -or ($_ -Match "EX:") } )
          $ForwardTo += ($Rule.ForwardAsAttachmentTo | ? {($_ -Match "SMTP") -or ($_ -Match "EX:")})
          If ($ForwardTo.Count -gt 0) {
             ForEach ($Recipient in $ForwardTo) {
                If ($Recipient -Match "EX:") {
                   # Recipient known in Exchange directory
                   $Ex = (Get-Recipient -Identity ($Recipient-Split "Ex:")[1].trim("]}")) 
                   $EmailAddress = $Ex.PrimarySmtpAddress }
                Else  {
                  # Simple SMTP address
                   $EmailAddress = ($Recipient -Split "SMTP:")[1].Trim("]") }
             }
             Write-Host $M.RecipientTypeDetails $M.DisplayName "has a rule to forward email to" $EmailAddress -ForegroundColor Red
             # Remove the rule if the address is unknown to the directory
             # If ($Ex -ne $Null) {
             # Remove-InboxRule -Identity $Rule.Identity -Confirm:$False
             # Write-Host "Rule" $Rule.Name "removed from mailbox!" }
             $NumberMbxWithRules++ }
       }
     }
}
Write-Host $NumberMbxWithRules "found with forwarding rules and" $NumberForwards "with forwarding addresses"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
