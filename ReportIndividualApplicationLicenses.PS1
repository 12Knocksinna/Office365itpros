# ReportIndividualApplicationLicenses.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportIndividualApplicationLicenses.PS1
# An exmaple of how to report license assignments for applications like Forms or Teams assigned to users through license bundles like Office 365 E3.
# Requires a connection to the Azure AD PowerShell endpoint. Run Connect-AzureAD before running this code.

# Updated 9-May-2025 to use the Microsoft Graph PowerShell SDK
Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop

# Define hash table of application licenses we might want to change - add more as you wish. Application identifiers and 
# plan identifiers can be found at https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
$Plans = @{}
$Plans.Add("Bookings", “199a5c09-e0ca-4e37-8f7c-b05d533e1ea2”)
$Plans.Add(“Exchange Online", "efb87545-963c-4e0d-99df-69c6916d9eb0”)
$Plans.Add("SharePoint Online", "be027f-2339-4123-9542-606e4d348a72”)
$Plans.Add("Viva Engage", "7547a3fe-08ee-4ccb-b430-5077c5041653")
$Plans.Add("Intune", "882e1d05-acd1-4ccb-8708-6ee03664b117")
$Plans.Add("Teams",  "57ff2da0-773e-42df-b2af-ffb7a2317929”)
$Plans.Add("Forms",  "2789c901-c14e-48ab-a76a-be334d9d793a”)
$Plans.Add("Stream", "9e700747-8b1d-45e5-ab8d-ef187ceec156”)
$Plans.Add("Planner", "b737dad2-2f6c-4c65-90e3-ca563267e8b9”)

# Define the SKU identifiers for bundled plans we expect to search - again, you can add more if your tenant uses other plans
$Skus = @{}
$Skus.Add("cf50bae9-29e8-4775-b07c-56ee10e3776d", "Office 365 E5 No Teams")
$Skus.Add("3271cf8e-2be5-4a09-a549-70fd05baaa17", "Microsoft 365 E5 No Teams")
$Skus.Add("7e74bd05-2c47-404e-829a-ba95c66fe8e5", "Teams")
$Skus.Add("6fd2c87f-b296-42f0-b197-1e91e994b900", "Office 365 E3")
$Skus.Add("c7df2760-2c81-4ef7-b578-5b5392b571df", "Office 365 E5")
$Skus.Add("26d45bd9-adf1-46cd-a9e1-51e9a5524128", "Office 365 E5 No Audio Conferencing")

$PlanId = $null
$Product = Read-Host "Enter the Office 365 application to report"
$PlanId = $Plans[$Product]
If ($null -eq $PlanId) { # Not found
   Write-Host ("Unable to find product {0} in our set of application SKUs” -f $Product) -ForegroundColor Red
   Break 
} Else {
   Write-Host ("Product {0} found with SKU {1}” -f $Product, $PlanId) -ForegroundColor Green
   Write-Host "Looking for user accounts with that service plan"
   [guid]$PlanSearch = $PlanId
   [array]$Users = Get-MgUser -filter "assignedPlans/any(s:s/serviceplanid eq $PlanSearch)" -ConsistencyLevel eventual -CountVariable Test `
      	-Property Id, displayName, userprincipalName, assignedLicenses, assignedPlans, department, country -All -PageSize 500
}

$PlanUsers = [System.Collections.Generic.List[Object]]::new() 
ForEach ($User in $Users) {

   [array]$UserPlans = Get-MgUserLicenseDetail -UserId $User.Id | Select-Object -ExpandProperty ServicePlans
   $Status = ($UserPlans | Where-Object {$_.ServicePlanId -eq $PlanId} | Select-Object -ExpandProperty ProvisioningStatus )
   [array]$LicenseNames = $Null
   ForEach ($UserSku in $User.AssignedLicenses.SkuId) {
      $SkuProductName = $Skus[$UserSku]
      $LicenseNames += $SkuProductName
   }
   If ($null -ne $LicenseNames) {
      $ReportLine  = [PSCustomObject] @{
         User       = $User.DisplayName 
         UPN        = $User.UserPrincipalName
         Department = $User.Department
         Country    = $User.Country
         SKU        = $PlanId
         Product    = $Product
         License    = ($LicenseNames -join ", ")
         Status     = $Status
      }
      $PlanUsers.Add($ReportLine) 
   }
}

Write-Host "Total Accounts scanned:" $PlanUsers.Count
$EnabledCount = $PlanUsers | Where-Object {$_.Status -eq "Success"}
Write-Host (“{0} is enabled for {1} accounts” -f $Product, $EnabledCount.Count)
$PlanUsers | Sort-Object User | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
