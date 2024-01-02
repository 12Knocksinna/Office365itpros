# Disable-JunkEmailOptions.PS1
# A runbook to look for recently added mailboxes and update their junk email configuration if necessary
# V1.0

# replace practical365.onmicrosoft.com with your Microsoft 365 service domain
Connect-ExchangeOnline -ManagedIdentity -Organization practical365.onmicrosoft.com

# Establish how far back we look
[string]$CheckDate = (Get-Date).AddDays(-7)
# Find matching mailboxes
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Filter "WhenCreated -gt '$CheckDate'"

If ($Mbx) {
    $Report = [System.Collections.Generic.List[Object]]::new()
    ForEach ($M in $Mbx) {
        $CurrentJunkMailConfiguration = Get-MailboxJunkEmailConfiguration -Identity $M.ExternalDirectoryObjectId
        If ($CurrentJunkMailConfiguration.Enabled -eq $True) {
            Set-MailboxJunkEmailConfiguration -Identity $M.ExternalDirectoryObjectId -Enabled $False `
            -BlockedSendersAndDomains $null -TrustedSendersAndDomains $null -TrustedRecipientsAndDomains $null 
    
            $DataLine = [PSCustomObject][Ordered]@{
                Mailbox                         = $M.DisplayName
                UPN                             = $M.UserPrincipalName
                'Trusted Senders and Domains'   = $CurrentJunkMailConfiguration.TrustedSendersAndDomains.Count
                'Blocked Senders and Domains'   = $CurrentJunkMailConfiguration.BlockedSendersAndDomains.Count
                'Contacts Trusted'              = $CurrentJunkMailConfiguration.ContactsTrusted
            }
            $Report.Add($DataLine)
        } Else {
            Write-Output ("Mailbox {0} already has the junk email rule disabled" -f $M.displayName)
        }
    }   
} Else {
    Write-Output "No mailboxes found to update..."
}

$Report