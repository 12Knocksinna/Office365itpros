# Remove-LicenseDisabledAccounts.PS1
# Example of how to remove licenses from disabled Entra ID accounts
# https://github.com/12Knocksinna/Office365itpros/blob/master/Remove-LicensesDisabledAccounts.PS1

# Connect to the Graph
Connect-MgGraph -NoWelcome -Scopes Directory.ReadWrite.All
# Define service plans that we shouldn't disable
[array]$Exclusions = "ENTERPRISEPACK", "SPE_E3", "SPE_E5", "SPE_E5_NOPSTNCONF", "SPB", "ENTERPRISEPREMIUM_NOPSTNCONF", "ENTERPRISE_PREMIUM"
Clear-Host

# Look for disabled accounts
Write-Host "Checking for disabled Entra ID user accounts..."
[array]$Accounts = Get-MgUser -All -Filter "accountEnabled eq false and userType eq 'member'" -Property `
       AccountEnabled, Id, DisplayName, userPrincipalName, assignedLicenses, licenseAssignmentStates
If (!($Accounts)) { 
    Write-Host "No disabled accounts found - exiting"
    Break
} Else {
    Write-Host ("Processing {0} disabled Entra ID user accounts" -f $Accounts.count)
}

[int]$LicenseRemovalCount = 0
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Account in $Accounts) {
    # Get the currently assigned licenses
    Write-Host ("Checking licenses for disabled account {0}" -f $Account.userPrincipalName) -ForegroundColor Yellow
    [array]$CurrentLicensesStage1 = Get-MgUserLicenseDetail -UserId $Account.Id | Select-Object SkuId, SkuPartNumber 
    [array]$LicensesToRemove = $Null
    [array]$RemovedLicenses = $Null
    [array]$CurrentLicensesStage2 = $Null
    $GroupName = $Null
    # See if any licenses are assigned by group-based licensing
    [array]$GroupAssignments = $Account.licenseAssignmentStates `
        | Where-Object {$Null -ne $_.AssignedByGroup -and $_.State -eq "Active"} | Select-Object SkuId, AssignedByGroup
    $GroupsHash = @{}
    ForEach ($G in $GroupAssignments) { $GroupsHash.Add([string]$G.SkuId,[string]$G.AssignedByGroup) }  
    ForEach ($License in $CurrentLicensesStage1)  {
    # If a group-based license is found, flag it. Otherwise add to the licenses to check in the next stage
        If ($License.SkuId -in $GroupAssignments.SkuId ) {
           $GroupName = (Get-MgGroup -GroupId $GroupsHash[$License.SkuId]).DisplayName
           Write-Host ("License {0} assigned by group-based licensing (group {1}): can't remove it" -f $License.SkuPartNumber, $GroupName ) -ForegroundColor Red 
        } Else {
            $CurrentLicensesStage2 += $License
        }       
    }    
    ForEach ($License in $CurrentLIcensesStage2) {
        If ($License.SkuPartNumber -in $Exclusions) {
            Write-Host ("Excluded license {0} found and will not be removed" -f $License.SkuPartNumber) -ForegroundColor Red
        } Else {
            $LicensesToRemove += $License.SkuId
            $RemovedLicenses += $License.SkuPartNumber
        }
     }
    # Remove the set of licenses that are safe to deduct, but only if there are some licenses to remove!
    If ($LicensesToRemove) {
        $LicenseNamesRemoved = $RemovedLicenses -join ", "
        Write-Host ("Removing {0} licenses from account {1}" -f $LicenseNamesRemoved, $Account.DisplayName)
        $Status = Set-MgUserLicense -UserId $Account.Id -AddLicenses @() -RemoveLicenses $LicensesToRemove
        If ($Status) {
            Write-Host ("{0} licenses removed: {1} from {2}" -f $LicensesToRemove.count, $LicenseNamesRemoved, $Account.DisplayName)
            $LicenseRemovalCount = $LicenseRemovalCount + $LicensesToRemove.Count
            $ReportLine = [PSCustomObject]@{
                User               = $Account.UserPrincipalName
                Name               = $Account.DisplayName
                Id                 = $Account.ID
                'Licenses Removed' = $LicenseNamesRemoved
                Timestamp          = (Get-Date -format 'dd-MMM-yyyy hh:mm')
            }
            $Report.Add($ReportLine)
        } Else {
            Write-Host ("Problem removing licenses from account {0} - please investigate!" -f $Account.DisplayName)
        }
    } Else {
        Write-Host ("No removable licenses found for account {0}" -f $Account.DisplayName) 
    }
}
Write-Host ""
Write-Host ("Script finished. A total of {0} licenses were removed from {1} accounts" -f $LicenseRemovalCount, $Accounts.count)
Write-Host ""
$Report | Format-Table Name, 'Licenses Removed', TimeStamp -AutoSize

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
