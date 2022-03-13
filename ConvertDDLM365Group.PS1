# ConvertDDLM365Group.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ConvertDDLM365Group.PS1
# Convert Dynamic Distribution List to a Microsoft 365 group and then team-enable the group
# Requires both the Exchange Online PowerShell module and the Microsoft Graph PowerShell SDK (or the Azure AD module - in that case
# convert the call to Get-MgDirectoryRoleMember to Get-AzureADDirectoryRoleMember and Get-MgUser to Get-AzureADUser

$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
# OK, we seem to be fully connected to Exchange Online. 

# Connect to the Graph
Connect-MgGraph -Scope "Directory.Read.All, User.Read.All"

$InputDDL = Read-Host "Enter the name of the Dynamic Distribution List to convert to a Microsoft 365 Group"
[array]$SourceDDL = Get-DynamicDistributionGroup -Identity $InputDDL -ErrorAction SilentlyContinue

If (!($SourceDDL)) {Write-Host ("Sorry! We can't find the {0} dynamic distribution list" -f $InputDDL); break}
If ($SourceDDL.Count -gt 1) {
   CLS
   Write-Host "We found multiple matching dynamic distribution lists"
   Write-Host "-----------------------------------------------------"
   Write-Host " "
   $SourceDDL | Format-Table DisplayName, Alias, PrimarySMTPAddress
   Write-Host " "
   Write-Host "Please try again..."; break }

[string]$SourceDDLId = $SourceDDL.ExternalDirectoryObjectId

# Now that we have a source DDL, let's get its membership
[array]$SourceMembers = Get-Recipient -RecipientPreviewFilter (Get-DynamicDistributionGroup -Identity $SourceDDLId).RecipientFilter
# could also be 
# [array]$SourceMembers = Get-DynamicDistributionGroupMember -Identity $SourceDDL.Id
# Throw away anything but user mailboxes because that's all a Microsoft 365 group supports
[array]$ValidMembers = $SourceMembers | ? {$_.RecipientTypeDetails -eq "UserMailbox"}

# We've got to assign an owner to the new Microsoft 365 group, so we need to have a default in case the source DDL doesn't have an owner
# Find the set of accounts that are Exchange admins (you can also use Get-AzureADDirectoryRoleMember here)
[array]$ExoAdmins = Get-MgDirectoryRoleMember -DirectoryRoleId "53add08e-5b0c-4276-a582-9ce02fb6c947" | Select Id, AdditionalProperties 
# Throw away any service principals which might have the Exchange Admin role
$ExoAdmins = $ExoAdmins | ? {$_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user'} | Select -ExpandProperty Id
# Select the first and use them as the default owner
$ExoDefaultAdmin = Get-MgUser -UserId $ExoAdmins[0] | Select -ExpandProperty UserPrincipalName
# Check that the group owner is a mailbox
$GroupOwner = Get-ExoMailbox -Identity $SourceDDL.Managedby -ErrorAction SilentlyContinue
# If it's null or something weird like a shared mailbox, use the default owner
If (($GroupOwner -eq $Null) -or ($GroupOwner.RecipientTypeDetails -ne "UserMailbox")) {
   $GroupOwner = $ExoDefaultAdmin }
Else {
   $GroupOwner = $GroupOwner.PrimarySmtpAddress
  }

# Populate other group properties
$AliasDDL = $SourceDDL.Alias + "M365"
$GroupDisplayName = $SourceDDL.DisplayName + " (Group)"

# Create the new Microsoft 365 Group
Write-Host "Creating the new Microsoft 365 group..."
$Description = "Created from the " + $SourceDDL.DisplayName + " dynamic distribution list on " + (Get-Date -Format g)
# If you use senstivity labels, add an appropriate label by passing its identifier in the -SensitivityLabel parameter.
$NewGroup = New-UnifiedGroup -DisplayName $GroupDisplayName –AccessType Private -Alias $AliasDDL -RequireSenderAuthenticationEnabled $True -Owner $SourceDDL.ManagedBy -AutoSubscribeNewMembers -Notes $Description
# Add the members to the group
Write-Host "Adding members from the dynamic distribution list to the Microsoft 365 group..."
Add-UnifiedGroupLinks -Identity $NewGroup.ExternalDirectoryObjectId -LinkType Members -Links $ValidMembers.PrimarySmtpAddress
Write-Host "Enabing Microsoft Teams for the Microsoft 365 group..."
New-Team -Group $NewGroup.ExternalDirectoryObjectId 
Write-Host "All done - new group and team created"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
