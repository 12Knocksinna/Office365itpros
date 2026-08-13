# Get-UnlicensedEntraP1Users.PS1
# A script to show how to determine the set of member user accounts which are covered by Conditional Access policies but do not have an
# Enrta P1 or P2 license assigned. The script uses the Microsoft Graph PowerShell SDK to fetch the required data.

# V1.0 12-Aug-2026

# GitHub Link: https://github.com/Office365itpros/Office365itpros

Connect-MgGraph -Scopes "Policy.Read.All", "Directory.Read.All", "User.Read.All", "GroupMember.Read.All"

# The first step is to check number of users covered by CA policies
# Start by finding the set of enabled Conditional Access policies...
Write-Host "Fetching enabled Conditional Access policies..." -ForegroundColor Green
[array]$Policies = Get-MgIdentityConditionalAccessPolicy -All -Filter "state eq 'enabled'"
# Find licensed users in the tenant (excluding guests and disabled accounts used for shared mailboxes etc.)
Write-Host "Fetching licensed member users in the tenant..." -ForegroundColor Green
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member' and accountEnabled ne false" `
    -ConsistencyLevel Eventual -CountVariable Count -Select Id,DisplayName,UserPrincipalName -PageSize 500

Write-Host ("Total of {0} users found in tenant" -f $Users.count)

[array]$TotalUsersInScope = $null
ForEach ($Policy in $Policies) {
    [array]$IncludedUsers = $Policy.Conditions.Users.IncludeUsers
    [array]$ExcludedUsers = $Policy.Conditions.Users.ExcludeUsers
    [array]$ExcludedGroups = $Policy.Conditions.Users.ExcludeGroups

    Write-Host ("Processing policy {0}" -f $Policy.DisplayName) -ForegroundColor Green

    If ($IncludedUsers -contains 'All') {
        [array]$ExcludedUserIds = $null
        If ($ExcludedUsers) {
            $ExcludedUserIds += $ExcludedUsers | Where-Object { $_ -ne 'GuestsOrExternalUsers' }
        }
        If ($ExcludedGroups) {
            ForEach ($GroupId in $ExcludedGroups) {
                $Members = Get-MgGroupMember -GroupId $GroupId -All
                $ExcludedUserIds += $Members.Id
            }
        }
        $ExcludedUserIds = $ExcludedUserIds | Select-Object -Unique
        $UsersInScope = $Users.Count - $ExcludedUserIds.Count
        $IncludedUserIds = $Users.Id | Where-Object { $_ -notin $ExcludedUserIds }
        [array]$TotalUsersInScope += $IncludedUserIds   

        Write-Host ("The {0} policy targets all users" -f $Policy.DisplayName) -ForegroundColor White
        If ($ExcludedUserIds) {
            Write-Host ("Member users excluded from policy: $ExcludedUserIds.Count") -ForegroundColor White
        }   
        Write-Host ("Member users in scope: {0}" -f $UsersInScope) -ForegroundColor Yellow
    } Else { # Policy focuses on specific users or groups
        [array]$IncludedUserIds = $null
        If ($Policy.Conditions.Users.IncludeGroups) {
            ForEach ($GroupId in $Policy.Conditions.Users.IncludeGroups) {
                $Members = Get-MgGroupMember -GroupId $GroupId -All
                $IncludedUserIds += $Members.Id
                $GroupName = (Get-MgGroup -GroupId $GroupId).DisplayName
                Write-Host ("Policy includes group: $GroupName ($($Members.Count) members)") -ForegroundColor White
            }
        }
        $IncludedUserIds = $IncludedUserIds | Select-Object -Unique
        Write-Host ("{1} member users in scope for policy {0}" -f $Policy.DisplayName, $IncludedUserIds.Count) -ForegroundColor Yellow
        [array]$TotalUsersInScope += $IncludedUserIds
    }
}
$TotalUsersInScope = $TotalUsersInScope | Sort-Object -Unique
Write-Host ("A total of {0} unique member user accounts come within the scope of enabled CA policies" -f $TotalUsersInScope.Count) -ForegroundColor Green

# Define the service plan identifiers for Entra P1 and P2 licenses
$SkuP1 = "41781fb2-bc02-4b7c-bd55-b576c07bb09d"
$SkuP2 = "eec0eb4f-6444-4f95-aba0-50c24d67f998"
    
# Find any product subscriptions in the tenant that include Entra P1 or P2 licenses
[array]$EntraSkus = Get-MgSubscribedSku -All | Where-Object {
    $_.ServicePlans.ServicePlanId -contains $SkuP1 -or
    $_.ServicePlans.ServicePlanId -contains $SkuP2 }  | Select-Object SkuId, SkuPartNumber, PrepaidUnits

If ($EntraSkus) {   
    [int]$TotalP1Licenses = ($EntraSkus.PrepaidUnits.Enabled | Measure-Object -Sum | Select-Object -ExpandProperty Sum)
    Write-Host ("Skus found in tenant with {0} Entra P1 licenses" -f $TotalP1Licenses) -ForegroundColor Green
}

# Do we have enough Entra P1 licenses to cover all users in scope for CA policies?
[int]$LicensingDeficit = $TotalUsersInScope.Count - $TotalP1Licenses
If ($LicensingDeficit -gt 0) {
    Write-Host ("Warning: {0} users within scope for CA policies do not have an Entra P1 license" -f $LicensingDeficit)  -ForegroundColor Red
} Else {
    Write-Host ("The tenant has sufficient Entra P1 licenses to cover {0} users" -f $TotalUsersInScope.Count) -ForegroundColor Green
}

# Licensed users for P1 or P2
[array]$LicensedP1Users = Get-MgUser -ConsistencyLevel Eventual -CountVariable Licenses `
    -Filter "(assignedPlans/any(s:s/servicePlanId eq $SkuP2) or assignedPlans/any(s:s/servicePlanId eq $SkuP1)) and userType eq 'Member' and accountEnabled ne false"

Write-Host ""
Write-Host "These users are not licensed for conditional access policies and are within the scope of current policies"

# Get the set of users who aren't licensed for Entra P1 or P2 but are in scope for CA policies
[array]$InScopeUsersWithOutLicenses = $TotalUsersInScope | Where-Object { $_ -notin $LicensedP1Users.Id }
# Create an array holding the user details for those users (most admins aren't great at reading GUIDs)
[array]$UsersWithoutEntraP1Licenses = $Users | Where-Object { $_.Id -in $InScopeUsersWithOutLicenses } | Select-Object DisplayName, UserPrincipalName, Id
Write-Host ("{0} member accounts" -f $UsersWithoutEntraP1Licenses.Count)
$UsersWithoutEntraP1Licenses

Write-Host ""
# Generate report file
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $OutputXLSXFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Unlicensed Entra P1 Accounts.xlsx"
    Remove-Item -Path $OutputXLSXFile -ErrorAction SilentlyContinue
    $UsersWithoutEntraP1Licenses | Export-Excel -Path $OutputXLSXFile -WorksheetName "Unlicensed Entra P1 Accounts" -Title ("Unlicensed Entra P1 Accounts {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "UnlicensedEntraP1Accounts" 
} Else {
    $OutputCSVFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Unlicensed Entra P1 Accounts.csv"
    Remove-Item -Path $OutputCSVFile -ErrorAction SilentlyContinue
    $UsersWithoutEntraP1Licenses | Export-Csv -Path $OutputCSVFile -NoTypeInformation -Encoding Utf8
}
  
If ($ExcelGenerated) {
  Write-Host ("An Excel workbook containing unlicensed Entra P1 accounts is available in {0}" -f $OutputXLSXFile)
} Else {
  Write-Host ("A CSV file containing unlicensed Entra P1 accounts is available in {0}" -f $OutputCSVFile)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.