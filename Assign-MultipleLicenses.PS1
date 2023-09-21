# Assign-MultipleLicenses.PS1
# An example script to demonstrate the principles of checking already-assigned licenses or 
# availablity of licenses before attempting to assign licenses to Entra ID user accounts
# https://github.com/12Knocksinna/Office365itpros/blob/master/Assign-MultipleLicenses.PS1

# Retrieve the products used in the tenant and construct a list of licenses and the units used
[array]$Skus = Get-MgSubscribedSku

# Build a hash table of license SKU identifiers and available units
$AvailableLicenses = @{}
ForEach ($S in $Skus) {
  $AvailableUnits = ($S.PrepaidUnits.Enabled - $S.ConsumedUnits)
  $AvailableLicenses.Add([string]$S.SkuId, $AvailableUnits)
}

# Define licenses that we'd like to assign 
[array]$DesiredSKUs = '6fd2c87f-b296-42f0-b197-1e91e994b900', 'f30db892-07e9-47e9-837c-80727f46fd3d', '1f2f344a-700d-42c9-9427-5cea1d5d7ba6', '6ee9b90c-0a7a-46c4-bc96-6698aa3bf8d2'

[array]$TargetUsers = "Andy.Ruth@Office365itpros.com", "Lotte.Vetler@Office365itpros.com"

$LicenseAssigments = [System.Collections.Generic.List[Object]]::new()    
ForEach ($User in $TargetUsers) {
   [array]$AdjustedSkus = $Null
   [array]$SkustoAssign = $Null
   Write-Host ("Checking licenses for assignment to {0}" -f $User) -ForegroundColor Yellow
   # Get current licenses
   [array]$CurrentLicenses = Get-MgUserLicenseDetail -UserId $User | Select-Object -ExpandProperty SkuId
   # Check each of the licenses we want to assign to make sure that it's not already assigned
   ForEach ($Sku1 in $DesiredSkus) {
    If ($Sku1 -in $CurrentLicenses) {
      Write-Host ("SKU {0} is already assigned to {1} so its assignment will be ignored" -f $SKU1, $User) -ForegroundColor DarkRed   
      $DataLine = [PSCustomObject][Ordered]@{
        Action    = "License already assigned"
        User      = $User
        License   = $SKU1
        Timestamp = (Get-Date)  }
      $LicenseAssigments.Add($DataLine)
    } Else {
      $AdjustedSKUs += $Sku1
    }
   } # End Foreach to determine if the user hasn't got a license already

   # Now check if a license is available for assignment for each SKU that isn't already assigned
   ForEach ($Sku2 in $AdjustedSkus) {
     If ($AvailableLicenses[$Sku2] -gt 0) {
        $SkusToAssign += $Sku2
     } Else {
        $DataLine = [PSCustomObject][Ordered]@{
            Action    = "License unavailable"
            User      = $User
            License   = $SKU2
            Timestamp = (Get-Date)  }
        $LicenseAssigments.Add($DataLine) 
       Write-Host ("No licenses are available to assign SKU {0} to user {1}" -f $SKU2, $User) -ForegroundColor Red
     }
   } # End ForEach to determine if licenses are available

   # Now do the magic with Set-MgUserLicense
   ForEach ($Sku3 in $SkusToAssign) {
    Write-Host ("Assiging SKU {0} to user {1}" -f $SKU3, $User)
    Try {
      $Status = Set-MgUserLicense -UserId $User -AddLicenses @{SkuId = $SKU3} -RemoveLicenses @()
    # Remove the assigned license from the available count
      $AvailableLicenses[$SKU3]  = ($AvailableLicenses[$SKU3] - 1)
      $DataLine = [PSCustomObject][Ordered]@{
        Action    = "License assigned"
        User      = $User
        License   = $SKU3
        Timestamp = (Get-Date)  }
      $LicenseAssigments.Add($DataLine)
    } Catch {
      Write-Host ("Whoops - Error assiging SKU {0} to user {1}" -f $SKU3, $User) }
    } # End of assigning licenses

} # End Foreach users

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
