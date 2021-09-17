# CreateCSVFilesForSKUsAndServicePlans.PS1
# A script to grab the SKU (products) used in a Microsoft 365 tenant and extract SKU and service plan information
# into CSV files so that they can be edited (to add display name information) and then used to generate a licensing
# report for the tenant.

# Uses cmdlets from the Microsoft Graph SDK for PowerShell
# Connect to the Graph, specifing the tenant and profile to use - Add your tenant identifier here
Connect-MgGraph -TenantId xxxxxxx-14fc-43a2-9a7a-d2e27f4f3478
Select-MgProfile beta
# Get prpducts used in tenant
[Array]$Skus = Get-MgSubscribedSku

# Generate CSV of all product SKUs used in tenant
$Skus | Select SkuId, SkuPartNumber  | Export-Csv -NoTypeInformation c:\temp\ListOfSkus.Csv
# Generate list of all service plans used in SKUs in tenant
$SPData = [System.Collections.Generic.List[Object]]::new()
ForEach ($S in $Skus) {
   ForEach ($SP in $S.ServicePlans) {
     $SPLine = [PSCustomObject][Ordered]@{  
         ServicePlanId = $SP.ServicePlanId
         ServicePlanName = $SP.ServicePlanName
         DisplayName = $SP.ServicePlanName }
     $SPData.Add($SPLine)
 }
}
$SPData | Sort ServicePlanId -Unique | Export-csv c:\Temp\ServicePlanData.csv -NoTypeInformation
# You need to edit these CSV files to add the human-friendly values to use for products (SKUs) and service plans
