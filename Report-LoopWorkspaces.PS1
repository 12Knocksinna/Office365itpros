# Report-LoopWorkspaces.PS1
# Script to report the usage of Loop workspaces in a tenant
#
# Connect to SharePoint Online
Import-Module Microsoft.Online.Sharepoint.PowerShell
Connect-SPOService -url https://redmondassociates-admin.sharepoint.com -Credential (Get-Credential)
# Connect to the Graph
Connect-MgGraph -NoWelcome -Scopes Directory.Read.All

[array]$LoopWorkspaces = Get-SPOContainer -OwningApplicationID a187e399-0c36-4b98-8f04-1edc167a0996
If (!($LoopWorkspaces)) {
    Write-Host "Can't get Loop workspaces - exiting"; break
}

# Define the licenses that allow users to create Loop workspaces
$LoopValidLicenses = @{}
$LoopValidLicenses.Add("f245ecc8-75af-4f8e-b61f-27d8114de5f3", "Microsoft 365 Business Standard")
$LoopValidLicenses.Add("cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46", "Microsoft 365 Business Premium")
$LoopValidLicenses.Add("05e9a617-0261-4cee-bb44-138d3ef5d965", "Microosft 365 E3")
$LoopValidLicenses.Add("0c21030a-7e60-4ec7-9a0f-0042e0e0211a", "Microsoft 365 E3 Hub Min 500")
$LoopValidLicenses.Add("06ebc4ee-1bb5-47dd-8120-11324bc54e06", "Microsoft 365 E5")

$Report = [System.Collections.Generic.List[Object]]::new()
$TotalBytes = 0; $LicenseOK = 0
ForEach ($LoopSpace in $LoopWorkspaces) {

    $LoopSpaceDetails =  Get-SPOContainer -OwningApplicationID a187e399-0c36-4b98-8f04-1edc167a0996 `
      -Identity $LoopSpace.ContainerId
    [array]$Owners = $LoopSpaceDetails.Owners
    ForEach ($Owner in $Owners) {
        $LicenseFound = $Null; $LoopLicenseStatus = "Unlicensed"
        $User = Get-MgUser -UserId $Owner -Property Id, displayName, department, UserPrincipalName, assignedLicenses
        [array]$SKUs = $User.AssignedLicenses.SkuId
        ForEach ($Sku in $Skus) {
            $LicenseFound = $LoopValidLicenses[$Sku]
            If ($LicenseFound) {
               $LicenseOK++
               $LoopLicenseStatus = "OK"
               $LicenseName = $LicenseFound
            }
        }
    }
    [array]$Members = $Null
    [array]$Managers = $LoopSpaceDetails.Managers
    ForEach ($Manager in $Managers) {
        $Member = Get-MgUser -UserId $Manager
        $Members += $Member.DisplayName
    }

    $StorageUsed = "{0:N2}" -f ($LoopSpaceDetails.StorageUsedInBytes/1MB)
    $TotalBytes = $TotalBytes + $LoopSpaceDetails.StorageUsedInBytes

    $ReportLine = [PSCustomObject]@{
        ContainerId    = $LoopSpace.ContainerId
        App            = $LoopSpaceDetails.OwningApplicationName
        Name           = $LoopSpace.ContainerName
        Description    = $LoopSpace.Description
        Owner          = $User.DisplayName
        UPN            = $User.UserPrincipalName
        License        = $LoopLicenseStatus
        Product        = $LicenseName
        Members        = ($Members -Join ", ")
        Created        = $LoopSpaceDetails.CreatedOn
        SiteURL        = $LoopSpaceDetails.ContainerSiteUrl
        "Storage (MB)" = $StorageUsed
    }
    $Report.Add($ReportLine)
}

$Report | Out-GridView
$Users = $Report | Select-Object Owner, License | Sort-Object Owner -Unique

$TotalLoopStorage = "{0:N2}" -f ($TotalBytes/1GB)
Write-Output ""
Write-Output ("Microsoft Loop is consuming {0} GB of the tenant's SharePoint Online storage quota" -f $TotalLoopStorage)
Write-Output ("Number of Loop workspaces found: {0}" -f $LoopWorkspaces.count)
Write-Output ""
Write-Output "Details of Users and License status (to create new Loop workspaces)"
$Users | Sort-Object License

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.