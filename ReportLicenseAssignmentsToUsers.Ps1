# Report license assignments to users 
# Uses Microsoft Online Services Module
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportLicenseAssignmentsToUsers.Ps1
If (-Not (Get-Module -Name MsOnline)) {
   Write-Host "Please run Connect-MSOnline before attempting to run this script"; break }

$Report = [System.Collections.Generic.List[Object]]::new()
$Users = Get-MsolUser -All | where {$_.isLicensed -eq $true}
Write-Host "Processing Users"
ForEach ($User in $Users) {
   $SKUs = @(Get-MsolUser -UserPrincipalName $User.UserPrincipalName | Select -ExpandProperty Licenses)
   ForEach ($Sku in $Skus) {  
   $Sku = $Sku.AccountSkuId.Split(":")[1]
   Switch ($Sku) {
    "EMSPREMIUM" { $License = "Enterprise Mobility & Security E5" }
    "ENTERPRISEPACK" { $License = "Office 365 E3" }
    "ENTERPRISEPREMIUM_NOPSTNCONF" { $License = "Office 365 E5 No PSTN" }
    "FLOW_FREE" { $License = "Power Automate" }
    "POWER_BI_STANDARD" { $License = "Power BI" }
    "RIGHTSMANAGEMENT_ADHOC" { $License = "Rights Management" }
    "SMB_APPS" { $License = "Business Apps" }
    "STREAM"  { $License = "Microsoft Stream"}
    "WIN_DEF_ATP" { $License = "Windows Defender ATP" }
    default   { $License = "Unknown license" }
   } #End Switch 
   $ReportLine = [PSCustomObject][Ordered]@{ 
        User       = $User.UserPrincipalName
        SKU        = $Sku
        License    = $License
        Name       = $User.DisplayName
        Title      = $User.Title
        City       = $User.City
        Country    = $User.UsageLocation
        Department = $User.Department
        CreatedOn  = Get-Date($User.WhenCreated) -Format g} 
   $Report.Add($ReportLine) }
}
Cls
Write-Host "License information"
Write-Host "-------------------"
$Groupdata = $Report | Group-Object -Property License | Sort Count -Descending | Select Name, Count
$GroupData
# Set sort properties so that we get ascending sorts for one property after another
$Sort1 = @{Expression='SKU'; Ascending=$true }
$Sort2 = @{Expression='Name'; Ascending=$true }

$Report | Select SKU, Name, User | Sort-Object $Sort1, $Sort2 | Export-CSV c:\Temp\UserLicenses.CSV -NoTypeInformation

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
