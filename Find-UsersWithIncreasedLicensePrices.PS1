# Find-UsersWithIncreasedLicensePrices.PS1

# A script to report price increases for licenses assigned to user accounts. 
# This is done by comparing the licenses assigned to user accounts with a list of licenses that have had price increases in the last 90 days. 
# The list of licenses with price increases is obtained from the Microsoft Graph API and is stored in the $IncreasedLicenses variable. 
# The script then checks each license assigned to user accounts against the list of licenses with price increases and outputs a message if a match is found.

# V1.0 8-June-2026
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-UsersWithIncreasedLicensePrices.PS1

# Permissions required are User.Read.All to fetch user accounts with details of their assigned licenses and
# LicenseAssignment.Read.All to read details of the tenant subscriptions
Connect-MgGraph -Scopes User.Read.All, LicenseAssignment.Read.All -NoWelcome

# Default currency for script
$DefaultCurrency = "US $"

# Create an array to hold details of the license price increase.
# Information from https://www.microsoft.com/en-us/licensing/news/2026-m365-packaging-pricing-updates

$Microsoft365PriceIncreases = @(

    # Office 365 Enterprise
    [PSCustomObject]@{
        SkuId = "6fd2c87f-b296-42f0-b197-1e91e994b900"
        Name  = "Office 365 E3"
        IncreasePercent = 13
        OldPrice = 23
        NewPrice = 26
    } # ENTERPRISEPACK 

    [PSCustomObject]@{
        SkuId = "c7df2760-2c81-4ef7-b578-5b5392b571df"
        Name  = "Office 365 E5"
        IncreasePercent = 8
        OldPrice = 38
        NewPrice = 41
    } # ENTERPRISEPREMIUM 

    [PSCustomObject]@{
        SkuId = "cf50bae9-29e8-4775-b07c-56ee10e3776d"
        Name  = "Office 365 E5 without Teams"
        IncreasePercent = 8
        OldPrice = '29.45'
        NewPrice = '32.45'
    } # ENTERPRISEPREMIUM without Teams

    # Microsoft 365 Enterprise 
    [PSCustomObject]@{
        SkuId = "05e9a617-0261-4cee-bb44-138d3ef5d965"
        Name  = "Microsoft 365 E3"
        IncreasePercent = 8
        OldPrice = 36
        NewPrice = 39
    } # SPE_E3 

    [PSCustomObject]@{
        SkuId = "3271cf8e-2be5-4a09-a549-70fd05baaa17"
        Name  = "Microsoft 365 E5 without Teams"
        IncreasePercent = 5
        OldPrice = '48.45'
        NewPrice  = '51.45'
    } # Microsoft 365 E5 without Teams
    [PSCustomObject]@{
        SkuId = "06ebc4ee-1bb5-47dd-8120-11324bc54e06"
        Name  = "Microsoft 365 E5"
        IncreasePercent = 5
        OldPrice = 57
        NewPrice = 60
    } # SPE_E5 

    # Frontline
    [PSCustomObject]@{
        SkuId = "44575883-256e-4a79-9da4-ebe9acabe2b2"
        Name  = "Microsoft 365 F1"
        IncreasePercent = 33
        OldPrice = '2.25'
        NewPrice = 3
    }

    [PSCustomObject]@{
        SkuId = "66b55226-6b4f-492c-910c-a3b7a3c9d993"
        Name  = "Microsoft 365 F3"
        IncreasePercent = 25
        OldPrice = 8
        NewPrice = 10
    }

    # Business
    [PSCustomObject]@{
        SkuId = "f245ecc8-75af-4f8e-b61f-27d8114de5f3"
        Name  = "Microsoft 365 Business Basic"
        IncreasePercent = 16
        OldPrice = 6
        NewPrice = 7
    }

    [PSCustomObject]@{
        SkuId = "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46"
        Name  = "Microsoft 365 Business Standard"
        IncreasePercent = 12
        OldPrice = '12.50'
        NewPrice = '14'
    }

    # Standalone / components
    [PSCustomObject]@{
        SkuId = "efccb6f7-5641-4e0e-bd10-b4976e1bf68e"
        Name  = "Enterprise Mobility + Security E3"
        IncreasePercent = 13
        OldPrice = '16.40'
        NewPrice = 18
    } # EMS 

    [PSCustomObject]@{
        SkuId = "078d2b04-f1bd-4111-bbd4-b4b1b354cef4"
        Name  = "Entra Plan 1 (Azure AD Premium P1)"
        IncreasePercent = 16
        OldPrice = 6
        NewPrice = 7
    } # AAD_PREMIUM 【

    [PSCustomObject]@{
        SkuId = "84a661c4-e949-4bd2-a560-ed7766fcaf2b"
        Name  = "Entra Plan 2 (Azure AD Premium P2)"
        IncreasePercent = 16
        OldPrice = 9
        NewPrice = 11
    } # AAD_PREMIUM_P2 【

    [PSCustomObject]@{
        SkuId = "cdd28e44-67e3-425e-be4c-737fab2899d3"
        Name  = "Microsoft 365 Apps for Business"
        IncreasePercent = 21
        OldPrice = '8.25'
        NewPrice = '10'
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
            OldPrice         = $Match.OldPrice
            NewPrice         = $Match.NewPrice
        }
    }
}

Write-Host ""
Write-Host "Licenses with price increases:" -ForegroundColor Cyan
$IncreasedLicenses | Select-Object Name, SKUPartNumber, ConsumedUnits, OldPrice, NewPrice | Format-Table -AutoSize

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

$Report = [System.Collections.Generic.List[Object]]::new()  
# Go through each user and check if they are assigned any of the licenses that have had price increases. If they are, output a message to the console.
ForEach ($User in $Users) {
    ForEach ($License in $User.AssignedLicenses.SkuId) {
        $LicenseMatch = $IncreasedLicenses | Where-Object { $_.SkuId -eq $License }
        If ($LicenseMatch ) {
            $ReportLine = [PSCustomObject][Ordered]@{
                Name            = $User.DisplayName
                UserPrincipalName = $User.UserPrincipalName
                SkuId           = $LicenseMatch.SkuId
                SkuPartNumber   = $LicenseMatch.SkuPartNumber
                IncreasePercent = $LicenseMatch.IncreasePercent
                OldPrice        = $LicenseMatch.OldPrice
                NewPrice        = $LicenseMatch.NewPrice
            }
            $Report.Add($ReportLine)
        }
    }  
}

Write-Host ""
Write-Host "Summary of licensed users with price increases:" -ForegroundColor Cyan
$Report | Select-Object Name, SkuPartNumber, IncreasePercent, OldPrice, NewPrice | Format-Table -AutoSize

Write-Host ""
Write-Host ("Total licensed users with price increases: {0}" -f $Report.Count) -ForegroundColor Yellow
Write-Host ("Total license spend before price increases: {0} {1:N2}" -f $defaultCurrency, ([decimal](($Report | Measure-Object -Property OldPrice -Sum).Sum))) -ForegroundColor Yellow
Write-Host ("Total license spend after price increases: {0} {1:N2}" -f $defaultCurrency, ([decimal](($Report | Measure-Object -Property NewPrice -Sum).Sum)))  -ForegroundColor Yellow
Write-Host ("Overall monthly price increase for affected licenses: {0} {1:N2}" -f $defaultCurrency, ([decimal](($Report | Measure-Object -Property NewPrice -Sum).Sum) - [decimal](($Report | Measure-Object -Property OldPrice -Sum).Sum)))  -ForegroundColor Yellow
Write-Host ("Overall percentage price increase for affected licenses: {0:N2}%" -f (([decimal](($Report | Measure-Object -Property NewPrice -Sum).Sum) - [decimal](($Report | Measure-Object -Property OldPrice -Sum).Sum)) / [decimal](($Report | Measure-Object -Property OldPrice -Sum).Sum) * 100))  -ForegroundColor Yellow

# Set up to use the ImportExcel module to generate Excel worksheets if it is available. If not, the report will be generated in CSV format instead.
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
}

# Generate the output file in either Excel worksheet or CSV format, depending on if the ImportExcel module is available. If interactive, create
# the file in the user's Downloads folder, otherwise create it in the current directory.
If ($ExcelGenerated) {
    $OutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Price Increases.xlsx"
    Remove-Item -Path $OutputFile -ErrorAction SilentlyContinue
    $Report | Export-Excel -Path $OutputFile -WorksheetName "Price Increases" -Title ("License Price Increases {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "PriceIncreases" -AutoSize -AutoFilter -FreezeTopRow
} Else {
    $OutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Price Increases.csv"
    Remove-Item -Path $OutputFile -ErrorAction SilentlyContinue
    $Report | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding Utf8
}

Write-Host ("Report generated: {0}" -f $OutputFile) -ForegroundColor Green
# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# The Automating Microsoft 365 with PowerShell eBook is available from https://o365itpros.gumroad.com/l/M365PS in EPUB and PDF format 
# or in paperback at https://amzn.to/3L4IBmU

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 