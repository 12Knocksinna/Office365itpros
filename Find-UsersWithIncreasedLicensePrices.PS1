# Find-UsersWithIncreasedLicensePrices.PS1

# V1.0 8-June-2026
# GitHub link: 

# A script to report price increases for licenses assigned to user accounts. 
# This is done by comparing the licenses assigned to user accounts with a list of licenses that have had price increases in the last 90 days. 
# The list of licenses with price increases is obtained from the Microsoft Graph API and is stored in the $IncreasedLicenses variable. 
# The script then checks each license assigned to user accounts against the list of licenses with price increases and outputs a message if a match is found.
# Create a table of the Microsoft 365 SKUs affected by the July 1, 2026 Price increases

# Create an array to hold details of the license price increase.
# Information from https://www.microsoft.com/en-us/licensing/news/2026-m365-packaging-pricing-updates
$Microsoft365PriceIncreases = @(

    # Office 365 Enterprise
    [PSCustomObject]@{
        SkuId = "6fd2c87f-b296-42f0-b197-1e91e994b900"
        Name  = "Office 365 E3"
        IncreasePercent = 13
    } # ENTERPRISEPACK 

    [PSCustomObject]@{
        SkuId = "c7df2760-2c81-4ef7-b578-5b5392b571df"
        Name  = "Office 365 E5"
        IncreasePercent = 8
    } # ENTERPRISEPREMIUM 

    [PSCustomObject]@{
        SkuId = "cf50bae9-29e8-4775-b07c-56ee10e3776d"
        Name  = "Office 365 E5 without Teams"
        IncreasePercent = 8
    } # ENTERPRISEPREMIUM without Teams

    # Microsoft 365 Enterprise 
    [PSCustomObject]@{
        SkuId = "05e9a617-0261-4cee-bb44-138d3ef5d965"
        Name  = "Microsoft 365 E3"
        IncreasePercent = 8
    } # SPE_E3 

    [PSCustomObject]@{
        SkuId = "3271cf8e-2be5-4a09-a549-70fd05baaa17"
        Name  = "Microsoft 365 E5 without Teams"
        IncreasePercent = 5
    } # Microsoft 365 E5 without Teams
    [PSCustomObject]@{
        SkuId = "06ebc4ee-1bb5-47dd-8120-11324bc54e06"
        Name  = "Microsoft 365 E5"
        IncreasePercent = 5
    } # SPE_E5 

    # Frontline
    [PSCustomObject]@{
        SkuId = "44575883-256e-4a79-9da4-ebe9acabe2b2"
        Name  = "Microsoft 365 F1"
        IncreasePercent = 33
    }

    [PSCustomObject]@{
        SkuId = "66b55226-6b4f-492c-910c-a3b7a3c9d993"
        Name  = "Microsoft 365 F3"
        IncreasePercent = 25
    }

    # Business
    [PSCustomObject]@{
        SkuId = "f245ecc8-75af-4f8e-b61f-27d8114de5f3"
        Name  = "Microsoft 365 Business Basic"
        IncreasePercent = 16
    }

    [PSCustomObject]@{
        SkuId = "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46"
        Name  = "Microsoft 365 Business Standard"
        IncreasePercent = 12
    }

    # Standalone / components
    [PSCustomObject]@{
        SkuId = "efccb6f7-5641-4e0e-bd10-b4976e1bf68e"
        Name  = "Enterprise Mobility + Security E3"
        IncreasePercent = 13
    } # EMS 

    [PSCustomObject]@{
        SkuId = "078d2b04-f1bd-4111-bbd4-b4b1b354cef4"
        Name  = "Entra ID Plan 1 (Azure AD Premium P1)"
        IncreasePercent = 16
    } # AAD_PREMIUM 【

    [PSCustomObject]@{
        SkuId = "cdd28e44-67e3-425e-be4c-737fab2899d3"
        Name  = "Microsoft 365 Apps for Business"
        IncreasePercent = 21
    } # O365_BUSINESS 
)

# Get tenant subscriptions and figure out which ones are affected by the price increases.
$TenantSkus = Get-MgSubscribedSku -All | Select-Object SkuId, SkuPartNumber,
   @{Name="ConsumedUnits";Expression={$_.ConsumedUnits}},
   @{Name="EnabledUnits";Expression={$_.PrepaidUnits.Enabled}}

# Join with price increases on SkuId
$IncreasedLicenses = ForEach ($Sku in $TenantSkus) {

    $Match = $Microsoft365PriceIncreases | Where-Object { $_.SkuId -eq $Sku.SkuId }

    If ($Match) {
        [PSCustomObject]@{
            Name             = $Match.Name
            SkuId            = $Sku.SkuId
            SkuPartNumber    = $Sku.SkuPartNumber
            ConsumedUnits    = $Sku.ConsumedUnits
            EnabledUnits     = $Sku.EnabledUnits
            IncreasePercent  = $Match.IncreasePercent
        }
    }
}

$IncreasedLicenses | Format-Table -AutoSize

# Go through each user and check if they are assigned any of the licenses that have had price increases. If they are, output a message to the console.

# Find licensed users...
Try {
    [array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
        -Property Id, displayName, userPrincipalName, assignedLicenses, SignInSessionsValidFromDateTime, SignInactivity, licenseAssignmentStates, AccountEnabled `
        -ConsistencyLevel eventual -CountVariable Records -All -PageSize 500 | Sort-Object displayName
    Write-Host ("Fetched {0} licensed user accounts." -f $Records) -ForegroundColor Green
} Catch {
    Write-Host ("Error fetching user data: {0}" -f $_.Exception.Message)
    Break
}

ForEach ($User in $Users) {
    ForEach ($License in $User.AssignedLicenses.SkuId) {
        $LicenseMatch = $IncreasedLicenses | Where-Object { $_.SkuId -eq $License }
        If ($LicenseMatch ) {
            Write-Host ("User {0} is assigned the {1} license ({2}% increase)." -f $User.DisplayName, $LicenseMatch.SkuPartNumber, $LicenseMatch.IncreasePercent) -ForegroundColor Yellow
            Break
        }
    }  
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 