# Enable-AllExchangeAuditEvents.PS1
# A script to update all user and shared mailboxes with the maximum level of auditing
# V1.0 December 2024 after addition of the PreservedMailItemProactively event

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

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.