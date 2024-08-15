# Disable-PowerShellAdmins.PS1
# Some code taken from https://thesysadminchannel.com/get-pim-role-assignment-status-for-azure-ad-using-powershell/
# The original article explaining how to disable access to Exchange PowerShell for all accounts except those holding the
# Exchange administrator and Global administrator roles is https://practical365.com/exchange-online-powershell-remove/
# This version shows how to use Azure AD Privileged Identity Management assignments instead of simple role memberships.
# V1.0   July 2023 Based on the AzureAD module
# V2.0   August 2024 Updated to use the Microsoft Graph PowerShell SDK. Tested with SDK V2.22.
# https://github.com/12Knocksinna/Office365itpros/blob/master/Disable-PowerShellAdmins.PS1

Connect-MgGraph -Scopes RoleAssignmentSchedule.Read.Directory, RoleEligibilitySchedule.Read.Directory, User.Read.All, Group.Read.All, GroupMember.Read.All -NoWelcome
$Modules = Get-Module | Select-Object -ExpandProperty Name
If ($Modules -notcontains "ExchangeOnlineManagement") {
   Connect-ExchangeOnline
}

# Find the identifiers for the Exchange administrator and Global administrator management roles
$ExoAdminRoleId = Get-MgDirectoryRoleTemplate | Where-Object {$_.displayName -eq "Exchange administrator"} | Select-Object -ExpandProperty Id
$GlobalAdminRoleId = Get-MgDirectoryRoleTemplate | Where-Object {$_.displayName -eq "Global administrator"} | Select-Object -ExpandProperty Id

# Output list to gather information about admin accounts
$Report = [System.Collections.Generic.List[Object]]::new()  
$OutputFile = "c:\temp\PIMAssignments.csv"  

Write-Output "Retrieving assignment information from Privileged Identity Management..."                    
# Get PIM assignments for accounts holding Exchange administrator or Global administrator roles
[array]$ActiveAssignments = Get-MgBetaRoleManagementDirectoryRoleAssignmentSchedule -Filter "(RoleDefinitionId eq '$($ExoAdminRoleId)') or (RoleDefinitionId eq '$($GlobalAdminRoleId)')" `
    -ExpandProperty RoleDefinition, Principal, DirectoryScope -All

# Filter out the Exchange administrators
[array]$ExoRoleMembers = $ActiveAssignments | Where-Object {$_.RoleDefinitionId -eq $ExoAdminRoleId} | Select-Object RoleDefinitionId, Principal, MemberType   
If (!($ExoRoleMembers)) { Write-Output "Can't find any Exchange administrators! Exiting..." ; break }                                                                                                

# Do the same for global administrators
[array]$GARoleMembers = $ActiveAssignments | Where-Object {$_.RoleDefinitionId -eq $GlobalAdminRoleId} | Select-Object RoleDefinitionId, Principal, MemberType
If (!($GARoleMembers)) { Write-Output "Can't find any global administrators! Exiting..." ; break }

Write-Output "Processing assignment information retrieved from Privileged Identity Management..."

# Process Exchange administrators to extract accounts with individual assignments and those who receive assignments
# through a group
ForEach ($Member in $ExoRoleMembers) {
  $User = Get-MgUser -UserId $Member.Principal.Id -ErrorAction SilentlyContinue -Property Id, displayName, userPrincipalName
  If ($User) {
      $ReportLine = [PSCustomObject]@{
        User            = $User.UserPrincipalName
        UserId          = $User.Id
        Name            = $User.DisplayName
        Role            = "Exchange administrator"
        MemberType      = $Member.MemberType }
      $Report.Add($ReportLine)
    } Else { # Must be a group
        [array]$GroupMembers = Get-MgGroupMember -GroupId $Member.Principal.id -ErrorAction SilentlyContinue
        If ($GroupMembers) {
           ForEach ($U in $GroupMembers) {
            $ReportLine = [PSCustomObject]@{
              User            = $U.additionalProperties.userPrincipalName
              UserId          = $U.Id
              Name            = $U.additionalProperties.displayName
              Role            = "Exchange administrator"
              MemberType      = $Member.MemberType }
           $Report.Add($ReportLine)
         }
      }
   } # End if
} #End ForEach

# Process global administrators
ForEach ($Member in $GARoleMembers) {
  $User = Get-MgUser -UserId $Member.Principal.Id -ErrorAction SilentlyContinue -Property Id, displayName, userPrincipalName
  If ($User) {
      $ReportLine = [PSCustomObject]@{
        User            = $User.UserPrincipalName
        UserId          = $User.Id
        Name            = $User.DisplayName
        Role            = "Global administrator"
        MemberType      = $Member.MemberType }
      $Report.Add($ReportLine)
   } Else { # Must be a group
        [array]$GroupMembers = Get-MgGroupMember -GroupId $Member.Principal.Id -ErrorAction SilentlyContinue
        If ($GroupMembers) {
           ForEach ($U in $GroupMembers) {
            $ReportLine = [PSCustomObject]@{
              User            = $U.additionalProperties.userPrincipalName
              UserId          = $U.Id
              Name            = $U.additionalProperties.displayName
              Role            = "Global administrator"
              MemberType      = $Member.MemberType }
           $Report.Add($ReportLine)
         }
      }
   } # End if
} #End ForEach

# Show what we found
$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation $OutputFile

# Create an array holding the user principal names of the two sets of administrator accounts
[array]$AdminAccounts = $Report.User | Sort-Object -Unique

# Find new Exchange user mailboxes that need to be processed
Write-Output "Checking Exchange Online user mailboxes to remove PowerShell access where needed..."
[array]$ExoMailboxes = Get-ExoMailbox -Filter {CustomAttribute5 -eq $Null} -ResultSize Unlimited -RecipientTypeDetails UserMailbox -Properties CustomAttribute5
ForEach ($Mbx in $ExoMailboxes) {
   # If not an admin holder, go ahead and block PowerShell
   If ($Mbx.userPrincipalName -notin $AdminAccounts) {
     Write-Output ("Blocking PowerShell access for mailbox {0}..." -f $Mbx.displayName)
     Try {
         Set-User -Identity $Mbx.userPrincipalName -RemotePowerShellEnabled $False -Confirm:$False
         $MessageText = "PowerShell disabled on " + (Get-Date -format s)
         Set-Mailbox -Identity $Mbx.userPrincipalName -CustomAttribute5 $MessageText
     }
     Catch {
         Write-Output ("Error disabling PowerShell for mailbox {0}" -f $Mbx.userPrincipalNane )
     }
   }
} # End ForEach

# And make sure that mailboxes belonging to an admin who's received a recent assignment has PowerShell access
Write-Output "Checking administrator mailboxes to make sure that they have PowerShell access..."
ForEach ($Mbx in $AdminAccounts) {
   [string]$mbx = $mbx
   $PSEnabled = (Get-User -Identity $Mbx  -ErrorAction SilentlyContinue).RemotePowerShellEnabled
   If (!($PsEnabled)) {
        Write-Output ("Resetting PowerShell access for admin account {0}" -f $Mbx)
        Set-User -Identity $Mbx -RemotePowerShellEnabled $True -Confirm:$False 
   }
}
   
Write-Host ("All done. CSV output is available in {0}." -f $OutputFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
