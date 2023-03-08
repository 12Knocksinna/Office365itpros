# ReportMFAStatusAdmins.PS1
# 
# Uses https://docs.microsoft.com/en-us/graph/api/userregistrationdetails-get
# Example of how to use the User Registration Details API to report admin accounts that are not MFA-enabled.
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportMFAStatusAdmins.PS1

Connect-MgGraph -Scopes UserAuthenticationMethod.Read.All, AuditLog.Read.All
Select-MgProfile Beta
Write-Host "Retrieving information about users holding Microsoft 365 administrative roles"
$AdminRoleHolders = [System.Collections.Generic.List[Object]]::new() 
[array]$AdminRoles = Get-MgDirectoryRole | Select-Object DisplayName, Id | Sort-Object DisplayName
ForEach ($Role in $AdminRoles) {
    [array]$RoleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $Role.Id | ? {$_.AdditionalProperties."@odata.type" -eq "#microsoft.graph.user"}
    ForEach ($Member in $RoleMembers) {
      $UserDetails = Get-MgUser -UserId $Member.Id
      $ReportLine  = [PSCustomObject] @{   
         User   = $UserDetails.UserPrincipalName
         Id     = $UserDetails.Id
         Role   = $Role.DisplayName
         RoleId = $Role.Id }
     $AdminRoleHolders.Add($ReportLine) }
}
$AdminRoleHolders = $AdminRoleHolders | Sort User
$Unique = $AdminRoleHolders | Sort-Object User -Unique

# Create a slightly different report where each user has their assigned roles in one record
$UniqueAdminRoleHolders = [System.Collections.Generic.List[Object]]::new() 
   ForEach($User in $Unique) {
     $Records = $AdminRoleHolders | Where-Object {$_.id -eq $User.Id}
     $AdminRoles = $Records.Role -join ", "
     $ReportLine  = [PSCustomObject] @{  
        Id      = $User.Id 
        User    = $User.User
        Roles   = $AdminRoles }
    $UniqueAdminRoleHolders.Add($ReportLine)
}
Write-Host ("There are {0} user accounts holding Microsoft 365 administrative roles." -f $UniqueAdminRoleHolders.count)

Write-Host "Scanning for MFA status"
# Retrieve member accounts that are licensed
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All

$UserRegistrationDetails = [System.Collections.Generic.List[Object]]::new() 
ForEach ($User in $Users) {
   Write-Host ("Checking admin roles and MFA status for {0}" -f $User.DisplayName)
   $Uri = "https://graph.microsoft.com/beta/reports/authenticationMethods/userRegistrationDetails/" + $User.Id
   $AccessMethodData = Invoke-MgGraphRequest -Uri $Uri -Method Get
   # Check if Admin
   $AdminAccount = $False; $AdminRolesHeld = $Null
   If ($user.id -in $UniqueAdminRoleHolders.Id) { 
      $AdminAccount = $True
      $AdminRolesHeld = ($UniqueAdminRoleHolders | ? {$_.Id -eq $User.Id} | Select -ExpandProperty Roles) }
   $ReportLine  = [PSCustomObject] @{ 
      User             = $User.Displayname
      Id               = $User.Id
      AdminAccount     = $AdminAccount
      AdminRoles       = $AdminRolesHeld
      MfaRegistered    = $AccessMethodData.isMfaRegistered
      defaultMfaMethod = $AccessMethodData.defaultMfaMethod
      isMfaCapable     = $AccessMethodData.isMfaCapable
      Methods          = $AccessMethodData.MethodsRegistered -join ", " }
  $UserRegistrationDetails.Add($ReportLine)
} #End ForEach

[Array]$ProblemAdminAccounts = $UserRegistrationDetails | Where-Object {$_.AdminAccount -eq $True -and $_.MfaRegistered -eq $False }
If ($ProblemAdminAccounts) {
   Cls
   Write-Host "The following accounts have administrative roles and are not enabled for MFA."
   Write-Host "-----------------------------------------------------------------------------"
   Write-Host ""
   $ProblemAdminAccounts | Select User, AdminRoles
}

$UserRegistrationDetails | Export-Excel -Path c:\temp\UserRegistrationDetails.xlsx

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
