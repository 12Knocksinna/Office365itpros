# ReportIndividualApplicationLicenses.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportIndividualApplicationLicenses.PS1
# An exmaple of how to report license assignments for applications like Forms or Teams assigned to users through license bundles like Office 365 E3.
# Requires a connection to the Azure AD PowerShell endpoint. Run Connect-AzureAD before running this code.
If (!(Get-Module -Name AzureAD).Name -eq "AzureAD") 
     {Write-Host "Please connect to Azure Active Directory PowerShell module before continuing"; break }

# Define hash table of application licenses we might want to change - add more as you wish. Application identifiers and 
# plan identifiers can be found at https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
$Plans = @{}
$Plans.Add(“199a5c09-e0ca-4e37-8f7c-b05d533e1ea2”, “Bookings”)
$Plans.Add(“efb87545-963c-4e0d-99df-69c6916d9eb0”, “Exchange Online”)
$Plans.Add(“5dbe027f-2339-4123-9542-606e4d348a72”, “SharePoint Online”)
$Plans.Add(“7547a3fe-08ee-4ccb-b430-5077c5041653”, “Yammer”)
$Plans.Add(“882e1d05-acd1-4ccb-8708-6ee03664b117”, “Intune”)
$Plans.Add(“57ff2da0-773e-42df-b2af-ffb7a2317929”, “Teams”)
$Plans.Add(“2789c901-c14e-48ab-a76a-be334d9d793a”, “Forms”)
$Plans.Add(“9e700747-8b1d-45e5-ab8d-ef187ceec156”, “Stream”)
$Plans.Add(“b737dad2-2f6c-4c65-90e3-ca563267e8b9”, “Planner”)

# Define the SKU identifiers for bundled plans we expect to search - again, you can add more if your tenant uses other plans
$Office365E5NoConf = "26d45bd9-adf1-46cd-a9e1-51e9a5524128"
$Office365E5 = "c7df2760-2c81-4ef7-b578-5b5392b571df"
$Office365E3 = "6fd2c87f-b296-42f0-b197-1e91e994b900"

Write-Host “Finding Azure AD Account Information”
$Users = Get-AzureADUser -All $True -Filter "Usertype eq 'Member'"
CLS
$Product = Read-Host "Enter the Office 365 application for a license check"
if (!($Plans.ContainsValue($Product))) { # Not found
   Write-Host “Can’t find” $Product “in our set of application SKUs”; break }
Foreach ($Key in $Plans.Keys) { # Lookup hash table to find product SKU
   If ($Plans[$Key] -eq $Product) { $PlanId = $Key }
}
$PlanUsers = [System.Collections.Generic.List[Object]]::new() 
ForEach ($User in $Users) {
  $License = $Null
  If ($PlanId -in $User.AssignedPlans.ServicePlanId) {
    $Status = ($User.AssignedPlans | ? {$_.ServicePlanId -eq $PlanId} | Select -ExpandProperty CapabilityStatus )
    If ($Office365E3 -in $User.AssignedLicenses.SkuId) {$LicenseName = "Office 365 E3"}
       ElseIf ($Office365E5 -in $User.AssignedLicenses.SkuId) {$LicenseName = "Office 365 E5"}
       ElseIf ($Office365E5NoConf -in $User.AssignedLicenses.SkuId) {$LicenseName = "Office 365 E5 without Audio Conferencing"}
       ElseIf ($Status -eq "Enabled") {$LicenseName = $Product }
       Else {$LicenseName = "Product not enabled"}

    $ReportLine  = [PSCustomObject] @{
          User       = $User.DisplayName 
          UPN        = $User.UserPrincipalName
          Department = $User.Department
          Country    = $User.Country
          SKU        = $PlanId
          Product    = $Product
          License    = $LicenseName
          Status     = $Status } 
    $PlanUsers.Add($ReportLine) }
}
Write-Host "Total Accounts scanned:" $PlanUsers.Count
$DisabledCount = $PlanUsers | ?{$_.Status -eq "Deleted"}
$EnabledCount = $PlanUsers | ? {$_.Status -eq "Enabled"}
Write-Host (“{0} is enabled for {1} accounts and disabled for {2} accounts” -f $Product, $EnabledCount.Count, $DisabledCount.Count)
$PlanUsers | Sort User | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
