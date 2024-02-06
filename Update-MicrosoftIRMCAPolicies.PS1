# Update-MicrosoftIRMCAPolicies.PS1
# Update conditional access policies to make sure that they exclude the Microsoft Rights Management Services app
# V1.0 6-Feb-2024
# https://github.com/12Knocksinna/Office365itpros/blob/master/Update-MicrosoftIRMCAPolicies.PS1

Connect-MgGraph -NoWelcome -Scopes Policy.ReadWrite.ConditionalAccess

# Define app id for the Microsoft Rights Management Services app
$MSFTIRMServices = "00000012-0000-0000-c000-000000000000"

# Parameters needed to update a CA policy
$Parameters = @{
    Conditions = @{
        applications = @{  
            excludeapplications = @(
                "00000012-0000-0000-c000-000000000000"
            )
        }
    }
}

[array]$Policies = Get-MgIdentityConditionalAccessPolicy -Filter "State eq 'enabled'" | Sort-Object DisplayName
ForEach ($Policy in $Policies) {
    Write-Host ("Checking conditional access policy {0}" -f $Policy.displayName)
    If ($Policy.conditions.applications.IncludeAuthenticationContextClassReferences) {
        Write-Host ("Policy {0} uses an authentication context. Can't apply an app exclusion" -f $Policy.displayName) -ForegroundColor Yellow
    } Else {
        [array]$ExcludedApps = $Policy.conditions.applications.excludeapplications
        If ($MSFTIRMServices -notin $ExcludedApps) {
            Write-Host ("Exclusion for Microsoft Rights Management Services app not present in CA policy {0}" -f $Policy.DisplayName)
            [array]$AuthenticationStrength = $Policy.grantcontrols | Select-Object -ExpandProperty AuthenticationStrength
            If (($Policy.grantcontrols.builtincontrols -eq 'mfa') -or ($AuthenticationStrength.AllowedCombinations)) {
                Write-Host "Checking policy to see if exclusion for Microsoft Rights Management Services app is possible" -ForegroundColor Red
                If ($Policy.grantcontrols.builtincontrols -eq 'passwordchange') {
                    Write-Host "Forced password change control means app exclusion is not possible" -ForegroundColor Yellow
                } Else {
                    Write-Host "Updating policy with exclusion" -ForegroundColor DarkRed
                    Update-MgIdentityConditionalAccessPolicy -BodyParameter $Parameters -ConditionalAccessPolicyId $Policy.Id
                }
            } Else {
                Write-Host "Policy doesn't use MFA - ignoring" -ForegroundColor Yellow
            }
        } Else {
            Write-Host "Exclusion for Microsoft Rights Management Services app present" -ForegroundColor DarkGray
        }
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.