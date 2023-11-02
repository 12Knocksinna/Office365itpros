# RemoveVivaEngageCoreServicePlan.PS1
# A script to disable the Viva Engage Core service plan from Office 365 E1, E3, and E5 licenses if
# the Yammer Enterprise service plan is already disabled
# https://github.com/12Knocksinna/Office365itpros/blob/master/RemoveVivaEngageCoreServicePlan.PS1
# Updated for SDK V2 21-Jul-2023

Connect-MgGraph -NoWelcome

# Define service plan and license identifiers
$YammerServicePlan = "7547a3fe-08ee-4ccb-b430-5077c5041653"
$VivaEngageCoreServicePlan = "a82fbf69-b4d7-49f4-83a6-915b2cf354f4"
$Office365E5NoConf = "26d45bd9-adf1-46cd-a9e1-51e9a5524128"
$Office365E3 = "6fd2c87f-b296-42f0-b197-1e91e994b900"
$Office365E5 = "c7df2760-2c81-4ef7-b578-5b5392b571df"
$Office365E1 = "18181a46-0d4e-45cd-891e-60aabd171b4e"

# Find and process licensed Entra ID users
Write-Host "Searching for licensed user accounts"
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All `
  -Property Id, displayName, userPrincipalName, assignedLicenses, department, country| Sort-Object DisplayName
If (!($Users)) { 
   Write-Host "Oddly, we can't find any licensed accounts ... exiting!" ; break 
}

$UserLicensesUpdated = [System.Collections.Generic.List[Object]]::new() ; $i =0
ForEach ($User in $Users) {
   [array]$DisabledSPs = $Null
   $i++
   Write-Host ("Checking licenses assigned to account {0} {1}/{2}" -f $User.Displayname, $i, $Users.count)
    $License = $Null; $LicenseSkuId = $Null
 # Find out what SKU the account has
   If ($Office365E3 -in $User.AssignedLicenses.SkuId) { 
      $LicenseSkuId = $Office365E3; $License = "Office 365 E3" 
   }
   If ($Office365E5NoConf -in $User.AssignedLicenses.SkuId)  { 
      $LicenseSkuId = $Office365E5NoConf; $License = "Office 365 E5 (No Conf)" 
   }
   If ($Office365E5 -in $User.AssignedLicenses.SkuId) { 
      $LicenseSkuId = $Office365E5; $License = "Office 365 E5"
   }
   If ($Office365E1 -in $User.AssignedLicenses.SkuId) { 
      $LicenseSkuId = $Office365E1; $License = "Office 365 E1" 
   }

   If ($Null -eq $LicenseSkuId) {
      Write-Host ("Can't find the right license SKU for account {0} - continuing" -f $User.UserPrincipalName)
      continue 
   }
 # Viva Engage Code only for Office 365 E3 or E1?
   If ($LicenseSkuId -eq $Office365E5 -or $LicenseSkuId -eq $Office365E5NoConf) {
      Write-Host ("Can't remove the Viva Engage Core service plan from the license assigned to {0} - continuing" -f $User.UserPrincipalName)
      continue 
   }

 # Examine service plans assigned to the account to see if they include Yammer
 [array]$AllLicenses = Get-MgUserLicenseDetail -UserId $User.Id | Select-Object -ExpandProperty ServicePlans | Sort-Object ServicePlanId -Unique
# [array]$Licenses = $AllLicenses | Where-Object {$_.ProvisioningStatus -eq 'Success'}
 [array]$DisabledLicenses = $AllLicenses | Where-Object {$_.ProvisioningStatus -eq 'Disabled'}
 
 # Check if Yammer Enterprise is disabled for the account. If it is, disable Viva Engage Core unless it's already disabled
   If ($YammerServicePlan -in $DisabledLicenses.ServicePlanId `
      -and $VivaEngageCoreServicePlan -notin $DisabledLicenses.ServicePlanId ) {

 # Add any previously disabled service plans if present and compose the license options
      If ($DisabledLicenses) { 
         $DisabledSPs = $DisabledLicenses.ServicePlanId 
      }
 # Add Viva Engage Core service plan
      $DisabledSPs += $VivaEngageCoreServicePlan  

      $LicenseOptions = @{SkuId = $LicenseSkuId ; DisabledPlans = $DisabledSPs} 

      Write-Host ("Disabling the Viva Engage Core service plan from the {0} license for account {1}" -f $License, $User.DisplayName) -foregroundcolor Red
      $Status = Set-MgUserLicense -UserId $User.Id -AddLicenses $LicenseOptions -RemoveLicenses @()
      If ($Status) {
         Write-Host "License adjustment succeeded"
      }
   
      $ReportLine  = [PSCustomObject] @{
         User                   = $User.DisplayName 
         UPN                    = $User.UserPrincipalName
         Department             = $User.Department
         Country                = $User.Country
         "Service Plan Removed" = $VivaEngageCoreServicePlan  
         License                = $License } 
      $UserLicensesUpdated.Add($ReportLine) 
   } #End if
} #End ForEach Users

Write-Host ""
Write-Host ("Viva Core Engage service plans were found and removed from these {0} accounts:" -f $LicenseUsers.count)
$UserLicensesUpdated.User
$UserLicensesUpdated | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
