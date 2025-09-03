# Migrate-UserProfileCardSettings.PS1
# An example script to show how to migrate custom properties used in the Microsoft 365 user profile card and replace those
# properties with standard Entra ID properties.

# V1.0 3-Sept-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Migrate-UserProfileCardSettings.PS1

# Use  Get-MgAdminPeopleProfileCardProperty | Format-List to see the set of custom properties defined for the Microsoft 365 profile card
# For this example, we use
# CustomAttribute12 = Employee Id   - Replaced by Entra ID EmployeeId
# CustomAttribute11 = Employee Type - Replaced by EntraID EmployeeType
# CustomAttribute15 = Cost Center   - Replaced by CostCenter

# Your tenant might use different attributes...

# The signed-in user must have Global Administrator or People Administrator role. The PeopleSettings.ReadWrite.All permission
# does not support app-only mode.

Connect-MgGraph -NoWelcome

# Check that we have the right permissions - in Azure Automation, we assume that the automation account has the right permissions
If ($Interactive) {
    [string[]]$CurrentScopes = (Get-MgContext).Scopes
    [string[]]$RequiredScopes = @('PeopleSettings.ReadWrite.All', 'User.ReadWrite.All')

    $CheckScopes =[object[]][Linq.Enumerable]::Intersect($RequiredScopes,$CurrentScopes)
    If ($CheckScopes.Count -ne 2) { 
        Write-Host ("To run this script, you need to connect to Microsoft Graph with the following scopes: {0}" -f $RequiredScopes) -ForegroundColor Red
        Disconnect-Graph
        Break
    }
}

# Get licensed member accounts for the tenant, excluding utility accounts
[array]$Users = Get-MgUser -All -PageSize 500 -Filter "usertype eq 'member' and accountenabled eq true and employeeType ne 'Utility' and assignedLicenses/`$count ne 0" `
    -Sort displayName -ConsistencyLevel Eventual -CountVariable Count `
    -Property DisplayName, UserPrincipalName, id, employeeType, userType, employeeOrgData, accountEnabled, employeeId, OnPremisesExtensionAttributes, department
    
Write-Host ("Found {0} licensed member accounts to process" -f $Users.Count) -ForegroundColor Green
$Report = [System.Collections.Generic.List[Object]]::new() 
[int]$i = 0
ForEach ($User in $Users) {
    $i++
    Write-Host ("Processing user {0} of {1}: {2}" -f $i, $Count, $User.DisplayName) -ForegroundColor Yellow
    # Fetch the data from the on-premises extension attributes
    $CostCenter = $User.onpremisesExtensionAttributes.extensionAttribute15
    $EmployeeId = $User.onpremisesExtensionAttributes.extensionAttribute12
    $EmployeeType = $User.onpremisesExtensionAttributes.extensionAttribute11
    $Division = $User.department

    If ($null -eq $CostCenter) { $CostCenter = "Not defined" }
    If ($null -eq $EmployeeId) { $EmployeeId = "Not defined" }
    If ($null -eq $EmployeeType) { $EmployeeType = "Permanent" } 
    If ($null -eq $Division) { $Division = "HQ" }

    $Parameters = @{
        employeeOrgData = @{
            "costCenter" = $CostCenter
            "division" = $Division
        }
        EmployeeId = $EmployeeId
        EmployeeType = $EmployeeType
    }
    Try {
        Update-MgUser -UserId $User.Id -BodyParameter $Parameters -ErrorAction Stop
        $DataLine = [PSCustomObject][ordered]@{
                User                = $User.displayName
                UserPrincipalName   = $User.UserPrincipalName
                EmployeeId          = $EmployeeId
                EmployeeType        = $EmployeeType
                CostCenter          = $CostCenter
                Division            = $Division
            }
            $Report.Add($DataLine)
    } Catch {
        Write-Host ("Failed to update user {0}: {1}" -f $User.DisplayName, $_.Exception.Message) -ForegroundColor Red
    }
}

# Now remove the old custom attributes

Write-Host "Removing old custom attributes from profile card" -ForegroundColor Green
Remove-MgAdminPeopleProfileCardProperty -ProfileCardPropertyId CustomAttribute11 
Remove-MgAdminPeopleProfileCardProperty -ProfileCardPropertyId CustomAttribute12
Remove-MgAdminPeopleProfileCardProperty -ProfileCardPropertyId CustomAttribute15

$Report | Export-Csv -Path ".\UserProfileCardMigrationReport.csv" -NoTypeInformation -Encoding UTF8
$Report | Out-GridView -Title "User Profile Card Migration Report"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 