# CreateCSVFilesForSKUsAndServicePlans.PS1
# A script to grab the SKU (products) used in a Microsoft 365 tenant and extract SKU and service plan information
# into CSV files so that they can be edited (to add display name information) and then used to generate a licensing
# report for the tenant.
# See https://practical365.com/create-licensing-report-microsoft365-tenant/ for the article relating to this code
# Uses cmdlets from the Microsoft Graph SDK for PowerShell
# Connect to the Graph, specifing the tenant and profile to use - Add your tenant identifier here
Connect-MgGraph -Scope Directory.Read.All -NoWelcome

#Import the Product names and service plan identifiers for licensing CSV file downloaded from https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
# Remember to move the CSV file downloaded from Microsoft to c:\temp\
[array]$Identifiers = Import-Csv -Path "C:\temp\Product names and service plan identifiers for licensing.csv"
#select all SKUs with friendly display name
[array]$SKU_friendly = $identifiers | Select-Object GUID, String_Id, Product_Display_Name -Unique
#select the service plans with friendly display name 
[array]$SP_friendly = $identifiers | Select-Object Service_Plan_Id, Service_Plan_Name, Service_Plans_Included_Friendly_Names -Unique

# Get prpducts used in tenant
[Array]$Skus = Get-MgSubscribedSku

# Generate CSV of all product SKUs used in tenant
$Skus | Select-Object SkuId, SkuPartNumber, @{Name = "DisplayName"; Expression = { ($SKU_friendly | Where-object -Property GUID -eq $_.SkuId).Product_Display_Name } } | Export-Csv -NoTypeInformation c:\temp\SkuDataComplete.csv
# Generate list of all service plans used in SKUs in tenant
$SPData = [System.Collections.Generic.List[Object]]::new()
ForEach ($S in $Skus) {
    ForEach ($SP in $S.ServicePlans) {
        $SPLine = [PSCustomObject][Ordered]@{  
            ServicePlanId          = $SP.ServicePlanId
            ServicePlanName        = $SP.ServicePlanName
            #use 'Service_Plans_Included_Friendly_Names' from $SKU_friendly for 'ServicePlanDisplayName'
            ServicePlanDisplayName = ($SP_friendly | Where-Object { $_.Service_Plan_Id -eq $SP.ServicePlanId }).Service_Plans_Included_Friendly_Names | Select-Object -First 1 
        }
        $SPData.Add($SPLine)
    }
}
$SPData | Sort-Object ServicePlanId -Unique | Export-csv c:\Temp\ServicePlanDataComplete.csv -NoTypeInformation
