# Update-BreakGlassUsersCAPolicies.PS1
# Make sure that our break glass users are excluded from every CA policy that's active in the organization
# https://github.com/12Knocksinna/Office365itpros/blob/master/Update-BreakGlassUsersCAPolicies.PS1
# V1.0 5-Dec-2023
# V1.1 12-Apr-2024 Added update for CA policies in report only mode

Connect-MgGraph -NoWelcome -Scopes Policy.ReadWrite.ConditionalAccess
# These are the policy states that will be updated with break glass accounts if required. 
[array]$PolicyStatesToProcess = "enabledForReportingButNotEnforced", "enabled","disabled"
[array]$BreakGlassUsers = "91813a30-f048-48f1-a0f2-fd7c72020515", "b7289bc7-7e4e-44e2-ae1b-7e13e94e3749"

[int]$PoliciesUpdated = 0
Write-Host "Finding conditional access policies... "
[array]$Policies = Get-MgIdentityConditionalAccessPolicy | Sort-Object DisplayName

ForEach ($Policy in $Policies) {
    Write-Host ("Checking settings for conditional access policy {0}" -f $Policy.displayName) -foregroundcolor Yellow    
    if (($Policy.Conditions.ClientApplications.IncludeServicePrincipals -ne $null) -or ($Policy.Conditions.ClientApplications.ExcludeServicePrincipals -ne $null)) {
        Write-Host ("CA Policy {0} is scoped to Workload Identities - it is not relevant to add break glass accounts here" -f $Policy.displayName) -ForegroundColor Green
    } else {
        [array]$ExcludedUsers = $Policy.conditions.users.excludeUsers
        [bool]$needChange = $false
        ForEach ($User in $BreakGlassUsers) {
            If ($User -notin $ExcludedUsers) {
            Write-Host ("Can't find user {0} in CA policy {1}" -f (Get-MgUser -UserId $User).DisplayName, $Policy.DisplayName)
            If ($Policy.State -in $PolicyStatesToProcess) {
                Write-Host ("Updating {0} with break glass accounts" -f $Policy.displayName) -ForegroundColor Red              
                $Policy.conditions.users.excludeUsers += $User
                $needChange = $true
            }
            }
        }
        if ($needChange) {        
            # Parameters needed to update a CA policy
            $Parameters = @{
                Conditions = @{
                    users = @{  
                        excludeUsers = @(
                            $Policy.conditions.users.excludeUsers
                        )
                    }
                }
            }
            $PoliciesUpdated++
            Update-MgIdentityConditionalAccessPolicy -BodyParameter $Parameters -ConditionalAccessPolicyId $Policy.Id
            $needChange = $false
        } else {
            Write-Host ("CA Policy {0} already has break glass accounts excluded" -f $Policy.displayName) -ForegroundColor Green
        }
    }
}

Write-Host ("{0} policies processed and {1} updated" -f $Policies.count, $PoliciesUpdated) 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
