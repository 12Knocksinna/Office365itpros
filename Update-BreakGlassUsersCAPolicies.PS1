# Update-BreakGlassUsersCAPolicies.PS1
# Make sure that our break glass users are excluded from every CA policy that's active in the organization

Connect-MgGraph -NoWelcome -Scopes Policy.ReadWrite.ConditionalAccess
[array]$BreakGlassUsers = "91813a30-f048-48f1-a0f2-fd7c72020515", "b7289bc7-7e4e-44e2-ae1b-7e13e94e3749"
# Parameters needed to update a CA policy
$Parameters = @{
    Conditions = @{
        users = @{  
            excludeUsers = @(
                "91813a30-f048-48f1-a0f2-fd7c72020515"
                "b7289bc7-7e4e-44e2-ae1b-7e13e94e3749"
            )
        }
    }
}
# If you prefer using a group instead, change excludeUsers to excludeGroups 

[array]$Policies = Get-MgIdentityConditionalAccessPolicy | Sort-Object DisplayName
ForEach ($Policy in $Policies) {
    Write-Host ("Checking conditional access policy {0}" -f $Policy.displayName)
    [array]$ExcludedUsers = $Policy.conditions.users.excludeUsers
    ForEach ($User in $BreakGlassUsers) {
        If ($User -notin $ExcludedUsers) {
           Write-Host ("Can't find user {0} in CA policy {1}" -f (Get-MgUser -UserId $User).DisplayName, $Policy.DisplayName)
           If ($Policy.State -eq 'enabled') {
              Write-Host "Policy is enabled so updating it with break glass accounts" -ForegroundColor Red
              Update-MgIdentityConditionalAccessPolicy -BodyParameter $Parameters -ConditionalAccessPolicyId $Policy.Id
           }
        }
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.