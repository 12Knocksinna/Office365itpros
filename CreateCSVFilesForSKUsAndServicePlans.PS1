# CreateCSVFilesForSKUsAndServicePlans.PS1
# A script to grab the SKU (products) used in a Microsoft 365 tenant and extract SKU and service plan information
# into CSV files so that they can be edited (to add display name information) and then used to generate a licensing
# report for the tenant.
# See https://practical365.com/create-licensing-report-microsoft365-tenant/ for the article relating to this code
# Uses cmdlets from the Microsoft Graph SDK for PowerShell
# Connect to the Graph, specifing the tenant and profile to use - Add your tenant identifier here
Connect-MgGraph -TenantId xxxxxxx-14fc-43a2-9a7a-d2e27f4f3478
Select-MgProfile beta
# Get prpducts used in tenant
[Array]$Skus = Get-MgSubscribedSku

# Generate CSV of all product SKUs used in tenant
$Skus | Select SkuId, SkuPartNumber  | Export-Csv -NoTypeInformation c:\temp\SkuDataComplete.csv.Csv
# Generate list of all service plans used in SKUs in tenant
$SPData = [System.Collections.Generic.List[Object]]::new()
ForEach ($S in $Skus) {
   ForEach ($SP in $S.ServicePlans) {
     $SPLine = [PSCustomObject][Ordered]@{  
         ServicePlanId = $SP.ServicePlanId
         ServicePlanName = $SP.ServicePlanName
         ServicePlanDisplayName = $SP.ServicePlanName }
     $SPData.Add($SPLine)
 }
}
$SPData | Sort ServicePlanId -Unique | Export-csv c:\Temp\ServicePlanDataComplete.csv -NoTypeInformation
# After they are generated, you need to edit the CSV files to add the human-friendly values to use for products (SKUs) and service plans
#
# This is what the SKU information should looks like:
#
# SkuId	                              SkuPartNumber	   DisplayName
#078d2b04-f1bd-4111-bbd4-b4b1b354cef4	AAD_PREMIUM	      Azure AD Premium P1
#84a661c4-e949-4bd2-a560-ed7766fcaf2b	AAD_PREMIUM_P2	   Azure AD Premium P2
#c52ea49f-fe5d-4e95-93ba-1de91d380f89	RIGHTSMANAGEMENT	Azure Information Protection P1
#90d8b3f8-712e-4f7b-aa1e-62e7ae6cbe96	SMB_APPS	         Business Apps (free)
# 
# And this is what the service plan information looks like:
# 
#ServicePlanId	                        ServicePlanName	            ServicePlanDisplayName
#041fe683-03e4-45b6-b1af-c0cdc516daee	POWER_VIRTUAL_AGENTS_O365_P2	Power Virtual Agents for Office 365 P2
#0683001c-0492-4d59-9515-d9a6426b5813	POWER_VIRTUAL_AGENTS_O365_P1	Power Virtual Agents for Office 365 P1
#07699545-9485-468e-95b6-2fca3738be01	FLOW_O365_P3	               Flow for Office 365 P3
#0898bdbb-73b0-471a-81e5-20f1fe4dd66e	KAIZALA_STANDALONE	         Kaizala Standalone
#0f9b09cb-62d1-4ff4-9129-43f4996f83f4	FLOW_O365_P1	               Flow for Office 365 P1
