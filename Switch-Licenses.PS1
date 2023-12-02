# Switch-Licenses.PS1
# Switch licenses for a set of Entra ID user accounts. This example shows the processing to switch user accounts
# from Office 365 E3 licenses to Microsoft 365 E5 licenses.
# https://github.com/12Knocksinna/Office365itpros/blob/master/Switch-Licenses.PS1
# V1.0 2-Dec-2023

Connect-MgGraph -Scopes User.ReadWrite.All -NoWelcome

# Define SKUs to process. We remove the original SkU and replace it with the new SKU
# Office 365 E3
$OriginalSku = '6fd2c87f-b296-42f0-b197-1e91e994b900'
# Microsoft 365 E5
$NewSku = '06ebc4ee-1bb5-47dd-8120-11324bc54e06'

# Create arrays of product information from the CSV file published at
# https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
Write-Host "Retrieving product SKU information..."
[array]$TenantSKUs = Get-MgSubscribedSKU | Select-Object SkuId, SkuPartNumber
[array]$ProductData = Import-CSV "C:\Temp\Product names and service plan identifiers for licensing.csv"
[array]$ProductInfo = $ProductData | Sort-Object GUID -Unique
# Create Hash table of the SKUs used in the tenant with the product display names from the Microsoft data file
$TenantSkuHash = @{}
ForEach ($P in $TenantSKUs) { 
    $ProductDisplayName = $ProductInfo | Where-Object {$_.GUID -eq $P.SkuId} | `
        Select-Object -ExpandProperty Product_Display_Name
    If ($Null -eq $ProductDisplayName) {
        $ProductDisplayname = $P.SkuPartNumber
    }
    $TenantSkuHash.Add([string]$P.SkuId, [string]$ProductDisplayName) 
}
# Extract service plan information and build a hash table
[array]$ServicePlanData = $ProductData | Select-Object Service_Plan_Id, Service_Plan_Name, Service_Plans_Included_Friendly_Names | `
    Sort-Object Service_Plan_Id -Unique
$ServicePlanHash = @{}
ForEach ($SP in $ServicePlanData) { 
    $ServicePlanHash.Add([string]$SP.Service_Plan_Id,[string]$SP.Service_Plans_Included_Friendly_Names)
}

# This section creates an array containing service plans from the original SKU with matching service plans
# from the new SKU. The idea is that when assigning the new license, if any service plans are disabled for the old SKU,
# appropriate service plans should be also disabled for the new SKU.
[array]$NewSP = $ProductData | Where-Object {$_.GUID -eq $NewSKU} | Sort-Object Service_Plan_Name -Descending
[array]$OldSP = $ProductData | Where-Object {$_.GUID -eq $OriginalSKU} 
$SwappedServicePlans = [System.Collections.Generic.List[Object]]::new()
ForEach ($ServicePlan in $OldSP) {
    $Check = $ServicePlan.Service_Plan_Name.SubString(0,$ServicePlan.Service_Plan_Name.Length -2) + "*"
    $Found = $NewSP | Where-Object {$_.Service_Plan_Name -like $Check} | Select-Object -First 1
    If ($Found) {
        $ReportLine = [PSCustomObject]@{ 
           OldSP        = $ServicePlan.Service_Plan_Id
           OldSPName    = $ServicePlan.Service_Plan_Name
           NewSP        = $Found.Service_Plan_Id
           NewSPName    = $Found.Service_Plan_Name 
        }
        $SwappedServicePlans.Add($ReportLine)
    }
}
# Make any adjustments - for example KAIZALA_O365_P3 is used in Office 365 E3 but is KAIZALA_STANDALONE in Microsoft 365 E5 
$ReportLine = [PSCustomObject]@{ 
    OldSP        = 'aebd3021-9f8f-4bf8-bbe3-0ed2f4f047a1'
    OldSPName    = 'KAIZALA_O365_P3'
    NewSP        = '0898bdbb-73b0-471a-81e5-20f1fe4dd66e'
    NewSPName    = 'KAIZALA_STANDALONE'
 }
 $SwappedServicePlans.Add($ReportLine)

# Fetch user accounts licensed with the original SKU
Write-Host ("Looking for user accounts licensed for {0} ({1}...)" -f ($TenantSkuHash[$OriginalSKU]), $OriginalSKU)
[array]$Users = Get-MgUser -filter "assignedLicenses/any(s:s/skuId eq $OriginalSKU)" -All `
    -Property Id, displayName, assignedLicenses | Sort-Object displayName
If (!($Users)) {
    Write-Host "No users found - exiting" ; break
}

# We have user accounts with the correct license, so we can replace them
Write-Host ("Preparing to replace licenses for {0} accounts" -f $Users.count)
$Report = [System.Collections.Generic.List[Object]]::new()
# Capture the information about disabled service plans for the currently assigned license 
ForEach ($User in $Users) {
    ForEach ($License in $User.AssignedLicenses) {
        If ($License.SkuId -eq $OriginalSku) {
            $DisabledServicePlans = $null
            [array]$DisabledServicePlans = $License.DisabledPlans 
            $ReportLine = [PSCustomObject]@{ 
                UserId          = $User.Id
                DisplayName     = $User.DisplayName
                SkuId           = $License.SkuId
                DisabledPlans   = $DisabledServicePlans }
        $Report.Add($ReportLine)        
        }
    }
}

# Loop to go through user accounts and remove the old license and assign the new - with appropriate service plans
ForEach ($User in $Report) {
    [array]$DisabledServicePlansToUse = $Null
    Write-Host ("Processing licenses for account: {0}" -f $User.DisplayName)
    If ($Null -ne $User.DisabledPlans) {
    # Process disabled service plans to make sure that we have a good set to bring to the new license
        # Write-Host ("Disabled service plans {0}" -f ($User.DisabledPlans -join ", "))  
        ForEach ($SP in $User.DisabledPlans) {
            $FoundSP = $SwappedServicePlans | Where-Object {$_.OldSP -eq $SP}
            If ($FoundSP) {
                $DisabledServicePlansToUse += $FoundSP.NewSP
            } Else {
                $DisabledServicePlansToUse += $SP
            }
        }
        # Write-Host ("Calculated service plans to disable {0}" -f ($DisabledServicePlansToUse -join ", "))
    }
    $Status = Set-MgUserLicense -UserId $User.UserId `
        -AddLicenses @{SkuId = $NewSKU; DisabledPlans = $DisabledServicePlansToUse} `
        -RemoveLicenses @() -ErrorAction SilentlyContinue
    If (!($Status)) {
        Write-Host "Error assigning license - please check availability" -ForegroundColor Red
    } Else {
        Write-Host ("{0} license assigned to account {1}" -f ($TenantSkuHash[$NewSKU]), $User.DisplayName )
        # Now to remove the old license
        $Status = Set-MgUserLicense -UserId $User.UserId -AddLicenses @() -RemoveLicenses $OriginalSku -ErrorAction SilentlyContinue
        If ($Status) {
            Write-Host ("{0} license removed from account {1}" -f ($TenantSkuHash[$OriginalSKU]), $User.DisplayName )
        }
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
