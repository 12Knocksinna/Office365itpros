# Create report of licenses assigned to users in a tenant

$Report = @() 
$Users = Get-MsolUser -All | where {$_.isLicensed -eq $true}
Write-Host "Processing Users"
ForEach ($User in $Users) {
   $SKUs = @(Get-MsolUser -UserPrincipalName $User.UserPrincipalName | Select -ExpandProperty Licenses)
   ForEach ($Sku in $Skus) {   
   $ReportLine = [PSCustomObject][Ordered]@{ 
        User       = $User.UserPrincipalName
        SKU        = $Sku.AccountSkuId.Split(":")[1]
        Name       = $User.DisplayName
        Title      = $User.Title
        City       = $User.City
        Country    = $User.UsageLocation
        Department = $User.Department
        CreatedOn  = Get-Date($User.WhenCreated) -Format g} 
   $Report += $ReportLine }
}
Cls
Write-Host "License information"
$Groupdata = $Report | Group-Object -Property SKU
$Groupdata | Sort Count -Descending | Select Name, Count
# Set sort properties so that we get ascending sorts for one property after another
$Sort1 = @{Expression='SKU'; Ascending=$true }
$Sort2 = @{Expression='Name'; Ascending=$true }

$Report | Select SKU, Name, User | Sort-Object $Sort1, $Sort2 | Export-CSV c:\Temp\UserLicenses.CSV -NoTypeInformation
