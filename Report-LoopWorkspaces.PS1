# Report-LoopWorkSpaces.PS1
# Example script to show how to report Loop workspaces, including pagination when there are more than 200 workspaces in the tenant
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-LoopWorkspaces.PS1
# V1.0 23-Nov-2023
# V1.1 29-March-2024
# V1.2 24-Jul-2024

# Connect to the Graph
Connect-MgGraph -NoWelcome -Scopes Directory.Read.All

# Connect to SharePoint Online
[array]$Domains = (Get-MgOrganization).verifiedDomains
$DefaultDomain = $Domains | Where-Object {$_.IsDefault -eq $true}
$SPOAdminRoot = ("https://{0}-admin.sharepoint.com" -f $DefaultDomain.Name.split('.')[0])
Write-Host "Connecting to SharePoint Online..."
Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell
Connect-SPOService -Url $SPOAdminRoot
If (Get-SPOTenant) {
    Write-Host ("Connected to SharePoint Online at {0}" -f $SPOAdminRoot)
} Else {
    Write-Host "Failed to connect to SharePoint Online"
    Break
}

Write-Host "Looking for Loop workspaces..."
[array]$LoopWorkspaces = Get-SPOContainer -OwningApplicationID a187e399-0c36-4b98-8f04-1edc167a0996 -Paged
If (!($LoopWorkspaces)) {
    Write-Host "Can't get Loop workspaces - exiting"; break
}

# Figure out if there are more than 200 workspaces.
[string]$Token = $null
If ($LoopWorkspaces[200]) {
    # Extract the token for the next page of workspace information
    $Token = $LoopWorkSpaces[200].split(":")[1].Trim()
    # Remove the last item in the array because it's the one that contains the token
    $LoopWorkspaces = $LoopWorkspaces[0..199]
}
Write-Host "Token for next page found - retrieving more workspaces..."
While ($Token) {
    # Loop while we can get a token for the next page of workspaces
    [array]$NextSetofWorkSpaces = Get-SPOContainer -OwningApplicationID a187e399-0c36-4b98-8f04-1edc167a0996 `
      -PagingToken $Token -Paged
    If ($NextSetofWorkSpaces[200]) {
        $Token = $NextSetofWorkSpaces[200].split(":")[1].Trim()
        $NextSetofWorkspaces = $NextSetofWorkspaces[0..199]
    } Else {
        $Token = $Null
        If (($NextSetofWorkSpaces[$NextSetofWorkspaces.count -1]) -eq "End of containers view.") {  
            # Remove the last item in the array because it contains the message "End of containers view."
            $NextSetofWorkspaces = $NextSetofWorkspaces[0..($NextSetofWorkspaces.count -2)]
        }             
    }
    $LoopWorkspaces += $NextSetofWorkspaces
} 
Write-Host ("{0} Loop workspaces found" -f $LoopWorkspaces.count)

$CSVOutputFile = "C:\temp\LoopWorkSpaces.CSV"
$LoopServicePlan = 'c4b8c31a-fb44-4c65-9837-a21f55fcabda'

# Define the licenses that allow users to create Loop workspaces. We use this information to figure out if the
# licenses assigned to the owner account allows them to create more workspaces.
$LoopValidLicenses = @{}
$LoopValidLicenses.Add("f245ecc8-75af-4f8e-b61f-27d8114de5f3", "Microsoft 365 Business Standard")
$LoopValidLicenses.Add("cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46", "Microsoft 365 Business Premium")
$LoopValidLicenses.Add("05e9a617-0261-4cee-bb44-138d3ef5d965", "Microosft 365 E3")
$LoopValidLicenses.Add("0c21030a-7e60-4ec7-9a0f-0042e0e0211a", "Microsoft 365 E3 Hub Min 500")
$LoopValidLicenses.Add("06ebc4ee-1bb5-47dd-8120-11324bc54e06", "Microsoft 365 E5")

$LoopWorkspaces = $LoopWorkspaces | Sort-Object ContainerName

$Report = [System.Collections.Generic.List[Object]]::new()
$TotalBytes = 0; $LicenseOK = 0; $i = 0
ForEach ($LoopSpace in $LoopWorkspaces) {
    $i++
    Write-Output ("Analyzing workspace {0} {1}/{2}" -f $LoopSpace.ContainerName, $i, $LoopWorkspaces.count)
    # Get detail of the workspace
    $LoopSpaceDetails =  Get-SPOContainer -OwningApplicationID a187e399-0c36-4b98-8f04-1edc167a0996 `
      -Identity $LoopSpace.ContainerId
    # Get detail about the owner
    [array]$Owners = $LoopSpaceDetails.Owners
    ForEach ($Owner in $Owners) {
        $LicenseFound = $Null; $LoopLicenseStatus = "Unlicensed";  $LicenseName = $Null
        # Find if the Loop service plan is successfully provisioned or is awaiting provisioning for the account
        [array]$UserLicenseData = Get-MgUserLicenseDetail -UserId $Owner
        $LoopLicense = $UserLicenseData | Select-Object -ExpandProperty ServicePlans | `
             Where-Object {$_.ServicePlanId -eq $LoopServicePlan} | Select-Object -ExpandProperty ProvisioningStatus
        If ($LoopLicense -in 'Success', 'PendingProvisioning') {
            $LicenseOK++
            $LoopLicenseStatus = "OK"
        }
        # Find what SKU the Loop service plan belongs to
        $User = Get-MgUser -UserId $Owner -Property Id, displayName, department, UserPrincipalName
        [array]$SKUs = $UserLicenseData.SkuId
        ForEach ($Sku in $Skus) {
            $LicenseFound = $LoopValidLicenses[$Sku]
            If ($LicenseFound) {
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
        ContainerId         = $LoopSpace.ContainerId
       # App                 = $LoopSpaceDetails.OwningApplicationName
        'Workspace Name'    = $LoopSpace.ContainerName
        Description         = $LoopSpace.Description
        Owner               = $User.DisplayName
        UPN                 = $User.UserPrincipalName
        Department          = $User.Department
        License             = $LoopLicenseStatus
        Product             = $LicenseName
        Members             = ($Members -Join ", ")
        Created             = $LoopSpaceDetails.CreatedOn
        SiteURL             = $LoopSpaceDetails.ContainerSiteUrl
        "Storage (MB)"      = $StorageUsed
    }
    $Report.Add($ReportLine)
}

$Report | Out-GridView -Title "Details of Loop Workspaces" -OutputMode Multiple
$Users = $Report | Select-Object Owner, License | Sort-Object Owner -Unique
$Report | Export-CSV -NoTypeInformation $CSVOutputFile

$TotalLoopStorage = "{0:N2}" -f ($TotalBytes/1GB)
Write-Output ""
Write-Output ("Microsoft Loop is consuming {0} GB of the tenant's SharePoint Online storage quota" -f $TotalLoopStorage)
Write-Output ("Number of Loop workspaces found: {0}" -f $LoopWorkspaces.count)
Write-Output ""
Write-Output "Details of Users and License status (to create new Loop workspaces)"
$Users | Sort-Object License
Write-Output ""
Write-Output ("Details can be found in the CSV file {0}" -f $CSVOutputFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
