# Find-ConnectionsM365AdminCenter.PS1
# Find the set of user accounts that have signed into the Microsoft 365 admin center in the last 30 days and
# report on their MFA status.

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-ConnectionsM365AdminCenter.PS1

Connect-MgGraph -Scope AuditLogs.Read.All

$M365AdminCenterId = (Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Office 365 Portal'").AppId
Write-Host "Checking for sign-ins to the Microsoft 365 Admin center..."
[array]$M365PortalSignIns = Get-MgBetaAuditLogSignIn -Filter "AppId eq '$M365AdminCenterId' and status/ErrorCode eq 0" -All -PageSize 500
If (!($M365PortalSignIns)) {
    Write-Host "No sign-ins found to the Microsoft 365 Admin center"
    Break
}
Write-Host ("Found {0} sign-ins to the Microsoft 365 Admin center for the last 30 days" -f $M365PortalSignIns.Count)
Write-Host "Checking MFA status for users who sign into the Microsoft 365 Admin center..."
[array]$UniqueUsers = $M365PortalSignIns | Sort-Object UserPrincipalName -Unique
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($User in $UniqueUsers) {
    $MFA = "Not enabled"
    If ($User.authenticationRequirement -eq 'multifactorauthentication') {
        $MFA = "Enabled"
    }
    $ReportLine = [PSCustomObject] @{ 
        User                = $User.UserDisplayName
        'MFA Status'        = $MFA
        'Last sign-in'      = $User.createdDateTime
    }
    $Report.Add($ReportLine)
}

Write-Host ""
Write-Host "Accounts that sign into the Microsoft 365 Admin Center"
Write-Host "------------------------------------------------------"
$Report 
$Report | Out-GridView -Title "Accounts that sign into the Microsoft 365 Admin Center"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.