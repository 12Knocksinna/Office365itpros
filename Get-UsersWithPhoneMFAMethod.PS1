# Get-UsersWithPhoneMFAMethod.PS1
# Find user accounts with a phone-based MFA method so that we can send them a nagging email
# V1.0 10-Jan-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-UsersWithPhoneMFAMethod.PS1

# Requires Users.Read.All and UserAuthenticationMethod.Read.All
Connect-MgGraph -NoWelcome

# Find licensed user accounts
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
    -ConsistencyLevel eventual -CountVariable UsersFound -Sort "displayName" `
    -Property Id, DisplayName, UserPrincipalName, UserType, AssignedLicenses
If (!$Users) {
    Write-Host "No user accounts found"
    Exit
}
Write-Output ("Checking MFA methods for {0} user accounts" -f $UsersFound)

ForEach ($User in $Users) {

}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.