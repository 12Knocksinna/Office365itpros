# Get-UnlicensedEntraP1Users.PS1
# A script to show how to determine the set of member user accounts which are covered by Conditional Access policies but do not have an
# Entra P1 or P2 license assigned. The script uses the Microsoft Graph PowerShell SDK to fetch the required data.

# V1.0 12-Aug-2026

# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Get-UnlicensedEntraP1Users.PS1

# Permissions required to run this script:
# - Policy.Read.All:        Read conditional access policy settings
# - User.ReadBasic.All:     Read  user information like identity, display name, and email address
# - GroupMember.Read.All:   Read group information

Connect-MgGraph -Scopes "Policy.Read.All", "User.ReadBasic.All", "GroupMember.Read.All" -NoWelcome

# The first step is to check number of users covered by CA policies
# Start by finding the set of enabled Conditional Access policies...
Write-Host "Fetching enabled Conditional Access policies..." -ForegroundColor Green
[array]$Policies = Get-MgIdentityConditionalAccessPolicy -All -Filter "state eq 'enabled'"
# Find licensed users in the tenant (excluding guests and disabled accounts used for shared mailboxes etc.)
Write-Host "Fetching licensed member users in the tenant..." -ForegroundColor Green
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member' and accountEnabled ne false" `
    -ConsistencyLevel Eventual -CountVariable Count -Select Id,DisplayName,UserPrincipalName -PageSize 500

Write-Host ("Total of {0} licensed member user accounts found in tenant" -f $Users.count)
$Report = [System.Collections.Generic.List[object]]::new()
[array]$TotalUsersInScope = $null

ForEach ($Policy in $Policies) {
    [array]$PolicyIncludedUsers = $Policy.Conditions.Users.IncludeUsers
    [array]$PolicyExcludedUsers = $Policy.Conditions.Users.ExcludeUsers
    $Groups = $null

    Write-Host ("Processing policy {0}" -f $Policy.DisplayName) -ForegroundColor Green

    If ($PolicyIncludedUsers -contains 'All') {
        [array]$ExcludedGroupUserIds = $null
        If ($PolicyExcludedUsers) {
            $ExcludedGroupUserIds += $PolicyExcludedUsers | Where-Object { $_ -ne 'GuestsOrExternalUsers' }
        }                       
        If ($Policy.Conditions.Users.ExcludeGroups) {
            ForEach ($GroupId in $Policy.Conditions.Users.ExcludeGroups) {
                [array]$Members = Get-MgGroupTransitiveMemberAsUser -GroupId $GroupId -All | Select-Object -ExpandProperty Id
                $ExcludedGroupUserIds += $Members
            }
        }
        $ExcludedGroupUserIds = $ExcludedGroupUserIds | Select-Object -Unique
        $UsersInScope = $Users.Count - $ExcludedGroupUserIds.Count
        $IncludedUserIds = $Users.Id | Where-Object { $_ -notin $ExcludedGroupUserIds }
        $TotalUsersInScope += $IncludedUserIds   

        Write-Host ("The {0} policy targets all users" -f $Policy.DisplayName) -ForegroundColor White
        If ($ExcludedGroupUserIds) {
            Write-Host ("Member users excluded from policy: $ExcludedGroupUserIds.Count") -ForegroundColor White
        }
        Write-Host ("Member users in scope: {0}" -f $UsersInScope) -ForegroundColor Yellow
    } Else { # Policy focuses on specific groups
        If ($Policy.Conditions.Users.IncludeGroups) {
            [array]$IncludedGroupUsersIds = $null
            ForEach ($GroupId in $Policy.Conditions.Users.IncludeGroups) {
                [array]$Members = Get-MgGroupTransitiveMemberAsUser -GroupId $GroupId -All | Select-Object -ExpandProperty Id
                $IncludedGroupUsersIds += $Members
                $GroupName = (Get-MgGroup -GroupId $GroupId).DisplayName
                $Groups += $GroupName 
                Write-Host ("Policy includes group: $GroupName ($($Members.Count) members)") -ForegroundColor White
            }
        }
        $IncludedGroupUsersIds = $IncludedGroupUsersIds 
        # Add individual users
        $IncludedUserIds = $PolicyIncludedUsers | Where-Object { $_ -notin $IncludedGroupUsersIds }           
        $IncludedUserIds += $IncludedGroupUsersIds
        $IncludedUserIds = $IncludedUserIds | Sort-Object -Unique
        Write-Host ("{1} member users in scope for policy {0}" -f $Policy.DisplayName, $IncludedGroupUsersIds.Count) -ForegroundColor Yellow
        $TotalUsersInScope += $PolicyIncludedUsers
    }
    # Remove any duplicates
    $TotalUsersInScope = $TotalUsersInScope | Sort-Object -Unique

    $ReportLine = [PSCustomObject]@{
        PolicyName      = $Policy.DisplayName
        UsersInScope    = $IncludedUserIds.Count
        UserIds         = $IncludedUserIds
        Groups          = $Groups -join ", "
    }
    $Report.Add($ReportLine)
    
}

$TotalUsersInScope = $TotalUsersInScope | Sort-Object -Unique
Write-Host ("A total of {0} unique member user accounts come within the scope of enabled CA policies" -f $TotalUsersInScope.Count) -ForegroundColor Green

# Define the service plan identifiers for Entra P1 and P2 licenses
$EntraPlan1 = "41781fb2-bc02-4b7c-bd55-b576c07bb09d"
$EntraPlan2 = "eec0eb4f-6444-4f95-aba0-50c24d67f998"

# Find any product subscriptions in the tenant that include Entra P1 or P2 licenses
[array]$EntraSkus = Get-MgSubscribedSku -All | Where-Object {
    $_.ServicePlans.ServicePlanId -contains $EntraPlan1 -or
    $_.ServicePlans.ServicePlanId -contains $EntraPlan2 }  | Select-Object SkuId, Sku``PartNumber, PrepaidUnits

Write-Host ""
Write-Host "Product SKUs with entitlement to Entra P1 or P2 licenses in the tenant:"
$EntraSKus

# Fetch licensed users with a license that includes an Entra P1 or P2 service plan
[array]$LicensedP1Users = Get-MgUser -ConsistencyLevel Eventual -CountVariable Licenses -PageSize 500 -Property Id, displayName, assignedPlans `
    -Filter "(assignedPlans/any(s:s/servicePlanId eq $EntraPlan2) or assignedPlans/any(s:s/servicePlanId eq $EntraPlan1)) and userType eq 'Member' and accountEnabled ne false"
    
# Check that the service plan is enabled and not deleted because it was proviously assighned and then removed
$ListCALicensedUsers = [System.Collections.Generic.List[object]]::new()
ForEach ($P1User in $LicensedP1Users) {
   $CheckPlan = $P1User.assignedPlans | Where-Object {$_.ServicePlanId -eq $EntraPlan1} | Select-Object -ExpandProperty CapabilityStatus
   If ($CheckPlan -eq "Enabled") {
        $ReportLine = [PSCustomObject]@{
            UserId          = $P1User.Id
            DisplayName     = $P1User.DisplayName
            UserPrincipalName = $P1User.UserPrincipalName
        }
        $ListCALicensedUsers.Add($ReportLine)
   }    
}

If ($EntraSkus) {   
    [int]$TotalP1Licenses = ($EntraSkus.PrepaidUnits.Enabled | Measure-Object -Sum | Select-Object -ExpandProperty Sum)
    Write-Host ("Products with entitlement to {0} Entra P1 licenses are available. The licenses are assigned to {1} user accounts." -f $TotalP1Licenses, $ListCALicensedUsers.Count) -ForegroundColor Green
}

Write-Host ""
Write-Host "These users are not licensed for conditional access policies and are within the scope of current policies"

# Get the set of users who aren't licensed for Entra P1 or P2 but are within the scope for at least one CA policy
[array]$InScopeUsersWithOutLicenses = $TotalUsersInScope | Where-Object { $_ -notin $ListCALicensedUsers.UserId }
# Create an array holding the user details for those users (most admins aren't great at reading GUIDs)
[array]$UsersWithoutEntraP1Licenses = $Users | Where-Object { $_.Id -in $InScopeUsersWithOutLicenses } | Select-Object DisplayName, UserPrincipalName, Id
Write-Host ""
Write-Host ("{0} member accounts should be assigned Entra P1 licenses" -f $UsersWithoutEntraP1Licenses.Count)
$UsersWithoutEntraP1Licenses | Sort-Object DisplayName | Format-Table -AutoSize

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