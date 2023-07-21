# RemoveKaizalaServicePlans.PS1
# A script to disable Kaizala service plans from Office 365 E1, E3, and E5 licenses
# https://github.com/12Knocksinna/Office365itpros/blob/master/RemoveKaizalaServicePlans.PS1

# Define service plan and license identifiers
$KaizalaPlan3 = "aebd3021-9f8f-4bf8-bbe3-0ed2f4f047a1" #E3
$KaizalaPlan2 = "54fc630f-5a40-48ee-8965-af0503c1386e" # E1
$KaizalaStandalone = "0898bdbb-73b0-471a-81e5-20f1fe4dd66e" # E5
$Office365E5NoConf = "26d45bd9-adf1-46cd-a9e1-51e9a5524128"
$Office365E3 = "6fd2c87f-b296-42f0-b197-1e91e994b900"
$Office365E5 = "c7df2760-2c81-4ef7-b578-5b5392b571df"
$Office365E1 = "18181a46-0d4e-45cd-891e-60aabd171b4e"

# Find and process licensed Azure AD users
$LicenseUsers = [System.Collections.Generic.List[Object]]::new() ; $i =0
Write-Host "Searching for licensed Azure AD accounts"
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All | Sort DisplayName
If (!($Users)) { Write-Host "Oddly, we can't find any licensed Azure AD users... exiting!" ; break }

ForEach ($User in $Users) {
 $i++
 Write-Host ("Checking licenses assigned to account {0} {1}/{2}" -f $User.Displayname, $i, $Users.count)
 $License = $Null; $LicenseSkuId = $Null
 # Find out what SKU the account has
 If ($Office365E3 -in $User.AssignedLicenses.SkuId) { $LicenseSkuId = $Office365E3 }
 If ($Office365E5NoConf -in $User.AssignedLicenses.SkuId)  { $LicenseSkuId = $Office365E5NoConf }
 If ($Office365E5 -in $User.AssignedLicenses.SkuId) { $LicenseSkuId = $Office365E5 }
 If ($Office365E1 -in $User.AssignedLicenses.SkuId) { $LicenseSkuId = $Office365E1 }

 # Examine service plans assigned to the account to see if they include Kaizala
 [array]$AllLicenses = Get-MgUserLicenseDetail -UserId $User.Id | Select-Object -ExpandProperty ServicePlans | Sort-Object ServicePlanId -Unique
 [array]$Licenses = $AllLicenses | Where-Object {$_.ProvisioningStatus -eq 'Success'}
 [array]$DisabledLicenses = $AllLicenses | Where-Object {$_.ProvisioningStatus -eq 'Disabled'}
 # Set the appropriate value for the Kaizala service plan appropriate for the license
 If ($KaizalaPlan2 -in $Licenses.ServicePlanId) { $License = "Office 365 E1"; $ServicePlanId = $KaizalaPlan2 }
 If ($KaizalaPlan3 -in $Licenses.ServicePlanId) { $License = "Office 365 E3"; $ServicePlanId = $KaizalaPlan3 }
 If ($KaizalaStandalone -in $Licenses.ServicePlanId) { $License = "Office 365 E5"; $ServicePlanId = $KaizalaStandalone }
   
 If ($License) { # Update the assigned license on the account

 # Add any previously disabled license if they exist to the Kaizala service plan and compose the license options
 [array]$DisabledSPs = $DisabledLicenses.ServicePlanId
 $DisabledSPs += $ServicePlanId

 $LicenseOptions = @{SkuId = $LicenseSkuId ; DisabledPlans = $DisabledSPs} 

 Write-Host ("Removing the Kaizala service plan from the {0} license for account {1}" -f $License, $User.DisplayName) -foregroundcolor Red
 $Status = Set-MgUserLicense -UserId $User.Id -AddLicenses $LicenseOptions -RemoveLicenses @()
 $ReportLine  = [PSCustomObject] @{
       User       = $User.DisplayName 
       UPN        = $User.UserPrincipalName
       Department = $User.Department
       Country    = $User.Country
       License    = $License } 
    $LicenseUsers.Add($ReportLine) }
}

CLS
Write-Host ("{0} Kaizala licenses were found and removed from these accounts:" -f $LicenseUsers.count)
$LicenseUsers.User


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
