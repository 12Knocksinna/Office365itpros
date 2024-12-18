# Enable-AllExchangeAuditEvents.PS1
# A script to update all user and shared mailboxes with the maximum level of auditing
# V1.0 December 2024

# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Enalble-AllExchangeAuditEvents.PS1

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
}

Write-Host "Looking for Exchange Online user and shared mailboxes"
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -ResultSize Unlimited | Sort-Object UserPrincipalName
If ($Mbx) {
    Write-Host ("{0} mailboxes found" -f $Mbx.Count)
} Else {
    Write-Host "No mailboxes found"
    Break
}
[array]$AuditAdminEvents = "Create", "FolderBind", "SendAs", "SendOnBehalf", "SoftDelete", "HardDelete", "Update", "Move", "Copy", "MoveToDeletedItems","UpdateFolderPermissions", "UpdateFolderPermissions", "UpdateComplianceTag", "UpdateInboxRules", "ApplyRecord", "RecordDelete", "AttachmentAccess", "PriorityCleanupDelete", "ApplyPriorityCleanup", "PreservedMailItemProactively"
[array]$AuditDelegateEvents = "Create", "SendAs", "SendOnBehalf", "SoftDelete", "HardDelete", "Update", "Move", "MoveToDeletedItems", "UpdateFolderPermissions", "UpdateComplianceTag", "UpdateInboxRules", "ApplyRecord", "RecordDelete", "AttachmentAccess", "PriorityCleanupDelete", "ApplyPriorityCleanup", "PreservedMailItemProactively"
[array]$AuditOwnerEvents = "Update", "Move", "MoveToDeletedItems", "SoftDelete", "HardDelete", "Create", "UpdateFolderPermissions", "PreservedMailItemProactively", "ApplyPriorityCleanup", "UpdateComplianceTag", "SearchQueryInitiated" , "PriorityCleanupDelete", "AttachmentAccess", "ApplyRecord", "RecordDelete", "UpdateCalendarDelegation", "UpdateInboxRules", "RemoveFolderPermissions", "ModifyFolderPermissions", "AddFolderPermissions"

ForEach ($M in $Mbx) {
    Write-Host ("Updating mailbox auditing for {0}" -f $M.UserPrincipalName)
    Set-Mailbox -Identity $M.ExternalDirectoryObjectId -AuditEnabled $true -AuditLogAgeLimit 365 -AuditAdmin @{Add=$AuditAdminEvents} -AuditDelegate @{Add=$AuditDelegateEvents} -AuditOwner @{Add=$AuditOwnerEvents}
}

Write-Host "All done..."