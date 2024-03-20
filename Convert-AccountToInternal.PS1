# Convert-AccountToInternal.PS1
# Example of how to convert an Entra ID account with an external identity to have an identity in the local
# tenant
# https://github.com/12Knocksinna/Office365itpros/blob/master/Convert-AccountToInternal.PS1

# Connect to the Microsoft Graph (tested with SDK V2.15)
Connect-MgGraph -NoWelcome -Scopes Directory.Read.All, User-ConvertToInternal.ReadWrite.All
$SourceUser = $null; $NewAccount = $null

# Find the domains registered for the tenant
$Domains =  Get-MgDomain -All
# Extract the default domain, which is used to create a new UPN and mail address for the user
$Domain = $Domains | Where-Object isDefault -match $True | Select-Object -ExpandProperty Id
If (!($Domain)) {
    Write-Host "Can't figure out the default domain for the tenant"
    Break
}
# Who are we converting?
$User = Read-Host "Enter name of external user to convert"
Write-Host ("Checking for user {0}" -f $User)
$SourceUser = Get-MgUser -Filter "displayName eq '$User'" -Property Id, displayName, userType, mail, userPrincipalName, givenName, surname

If (!($SourceUser)) {
    Write-Host ("Can't find user account for {0}" -f $User)
    Break
}
If ($SourceUser.mail -in $Domains.id) {
    Write-Host ("The {0} account is already a member of this tenant... " -f $User)
    Break
}
Write-Host ("Converting {0} to be an internal account" -f $SourceUser.UserPrincipalName) -ForegroundColor Red

# Make sure that we have a nice user principal name
If ($SourceUser.givenName -and $SourceUser.surname) {
    $NewUserPrincipalName = ("{0}.{1}@{2}" -f $SourceUser.givenname, $SourceUser.surname, $Domain)
} Else {
    $FirstName = $SourceUser.displayName.Split(' ')[0]
    $Surname = $SourceUser.displayName.Split(' ')[1]
    $NewUserPrincipalName = ("{0}.{1}@{2}" -f $FirstName, $Surname, $Domain)
}

# Create a new password
$NewPassword = $(1..1 | ForEach-Object { [char[]]"!@#$%" | Get-Random })
$NewPassword += $(1..3 | ForEach-Object { [char[]]"ABCDEFGHJKMNPQRSTUVWXYZ" | Get-Random }) -join ""
$NewPassword += $(1..3 | ForEach-Object { [char[]]"abcdefghkmnopqrstuvwxyz" | Get-Random }) -join ""
$NewPassword += $(1..2 | ForEach-Object { [char[]]"0123456789" | Get-Random }) -join "" 
$NewPassword += $(1..1 | ForEach-Object { [char[]]"!@#$%" | Get-Random })
$PasswordProfile = @{}
$PasswordProfile.Add('password',$NewPassword)
$PasswordProfile.Add('forceChangePasswordNextSignIn', $true)

# Create the parameters to convert the account
$NewAccountParameters = @{}
$NewAccountParameters.Add('userPrincipalName', $NewUserPrincipalName)
$NewAccountParameters.Add('passwordProfile', $PasswordProfile)

Write-Host "Switching the account to be internal..."
# Switch the account to make it internal
$Uri = ("https://graph.microsoft.com/Beta/users/{0}/convertExternalToInternalMemberUser" -f $SourceUser.Id)
$NewAccount = Invoke-MgGraphRequest -Uri $Uri -Body $NewAccountParameters -Method POST -ContentType "application/json"

# If we get back some account details, check to make sure that they're what we expect
If ($NewAccount) {
    $CheckNewAccount = Get-MgUser -UserId $SourceUser.Id -Property id, displayName, userPrincipalName, UserType
    If ($CheckNewAccount.usertype -eq 'Member' -and $CheckNewAccount.UserPrincipalName -eq $NewUserPrincipalName) {
        Update-MgUser -UserId $CheckNewAccount.Id -Mail $NewUserPrincipalName
        $RevokeStatus = Revoke-MgUserSignInSession -UserId $CheckNewAccount.Id
        Write-Host ("{0} is now a {1} account" -f $CheckNewAccount.UserPrincipalName, $CheckNewAccount.userType)
        Write-Host ("The temporary password for the account is {0}" -f $NewPassword)
        Write-Host ("Remember to assign some licenses to the converted account and to remove it from the previous source.")
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.