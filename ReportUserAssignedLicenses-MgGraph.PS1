# ReportUserAssignedLicenses-MgGraph.PS1
# Create a report of licenses assigned to Entra ID user accounts using the Microsoft Graph PowerShell SDK cmdlets
# Github link: https://github.com/12Knocksinna/Office365itpros/blob/master/ReportUserAssignedLicenses-MgGraph.PS1
# See https://practical365.com/create-licensing-report-microsoft365-tenant/ for an article describing how to run the report and 
# https://practical365.com/report-user-license-costs/ for information about how to include licensing cost information
# in the output

# V1.1  27-Sep-2022   Add sign in data for users and calculate how long it's been since they signed in and used a license.
# V1.2  23-Nov-2022   Added SKU usage summary to HTML report
# V1.3  29-Sep-2023   Added support for group-based licensing
# V1.4  13-Oct-2023   Fixed some bugs
# V1.5  26-Jan-2024   Added license pricing computation
# V1.5  8-Feb-2024    Added cost analysis for departments and countries
# V1.6  12-Feb-2024   Added info to report when license costs can't be attributed to countries or departments because of missing user account properties
# V1.7  7-Mar-2024    Added company name to the set of properties output by report
# V1.8  17-Mar-2024   Tidy up
# V1.9  1-June-2024   Add cost analysis for inactive and disabled accounts
# V1.91 26-June-2024  Fix bug with color formatting of inactive account column
# V1.92 04-Jul-2024   Add optional cost center license analysis based on cost center details stored in a
#                     custom attribute for Exchange mailboxes that's synchronized to Entra ID. The custom attribute
#                     to use is held in $CostCenterAttribute. If this variable is undefined or $null, the report
#                     script ignores cost centers
# V1.93 05-Sep-2024   Add analysis of costs by company name
# V1.94 08-Sep-2024   Remove licenses assigned for expired subscriptions from the report

Function Get-LicenseCosts {
  # Function to calculate the annual costs of the licenses assigned to a user account  
  [cmdletbinding()]
  Param( [array]$Licenses )
  [int]$Costs = 0
  ForEach ($License in $Licenses) {
    Try {
      [string]$LicenseCost = $PricingHashTable[$License]
      # Convert monthly cost to cents (because some licenses cost sums like 16.40)
      [float]$LicenseCostCents = [float]$LicenseCost * 100
      If ($LicenseCostCents -gt 0) {
        # Compute annual cost for the license
        [float]$AnnualCost = $LicenseCostCents * 12
        # Add to the cumulative license costs
        $Costs = $Costs + ($AnnualCost)
        # Write-Host ("License {0} Cost {1} running total {2}" -f $License, $LicenseCost, $Costs)
      }
    }
    Catch {
      Write-Host ("Error finding license {0} in pricing table - please check" -f $License)
    }
  }
  # Return 
  Return ($Costs / 100)
} 

[datetime]$RunDate = Get-Date
[string]$ReportRunDate = Get-Date ($RunDate) -format 'dd-MMM-yyyy HH:mm'
$Version = "1.95"

# Default currency - can be overwritten by a value read into the $ImportSkus array
[string]$Currency = "USD"

# Connect to the Graph, specifing the tenant and profile to use - Add your tenant identifier here
Connect-MgGraph -Scope "Directory.AccessAsUser.All, Directory.Read.All, AuditLog.Read.All" -NoWelcome

<#
Alternative: Use Application ID and Secured Password for authentication (you could also pass a certificate thumbprint)
$ApplicationId = "<applicationId>"
$SecuredPassword = "<securedPassword>"
$tenantID = "<tenantId>"

$SecuredPasswordPassword = ConvertTo-SecureString -String $SecuredPassword -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationId, $SecuredPasswordPassword
Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $ClientSecretCredential
#>

# This step depends on the availability of some CSV files generated to hold information about the product licenses used in the tenant and 
# the service plans in those licenses. See https://github.com/12Knocksinna/Office365itpros/blob/master/CreateCSVFilesForSKUsAndServicePlans.PS1 
# for code to generate the CSVs. After the files are created, you need to edit them to add the display names for the SKUs and plans.
# Build Hash of Skus for lookup so that we report user-friendly display names - you need to create these CSV files from SKU and service plan
# data in your tenant.

$SkuDataPath = "C:\temp\SkuDataComplete.csv"
$ServicePlanPath = "C:\temp\ServicePlanDataComplete.csv"
$UnlicensedAccounts = 0

# Define this variable if you want to do cost center reporting based on a cost center stored in one of the
# 15 Exchange Online custom attributes synchronized to Entra ID. Use the Entra ID attribute (like extensionAttribute6) 
# name not the Exchange Online attribute name (CustomAttribute6) Set the variable to $null or don't define it at all 
# to ignore cost centers
#$CostCenterAttribute = "extensionAttribute6"

If ((Test-Path $skuDataPath) -eq $False) {
    Write-Host ("Can't find the product data file ({0}). Exiting..." -f $skuDataPath) ; break 
}
If ((Test-Path $servicePlanPath) -eq $False) {
    Write-Host ("Can't find the serivice plan data file ({0}). Exiting..." -f $servicePlanPath) ; break 
}
   
$ImportSkus = Import-CSV $skuDataPath
$ImportServicePlans = Import-CSV $servicePlanPath
$SkuHashTable = @{}
ForEach ($Line in $ImportSkus) { $SkuHashTable.Add([string]$Line.SkuId, [string]$Line.DisplayName) }
$ServicePlanHashTable = @{}
ForEach ($Line2 in $ImportServicePlans) { $ServicePlanHashTable.Add([string]$Line2.ServicePlanId, [string]$Line2.ServicePlanDisplayName) }

# If pricing information is in the $ImportSkus array, we can add the information to the report. We prepare to do this
# by setting the $PricingInfoAvailable to $true and populating the $PricingHashTable
$PricingInfoAvailable = $false

If ($ImportSkus[0].Price) {
  $PricingInfoAvailable = $true
  $Global:PricingHashTable = @{}
  ForEach ($Line in $ImportSkus) { 
    $PricingHashTable.Add([string]$Line.SkuId, [string]$Line.Price) 
  }
  If ($ImportSkus[0].Currency) {
    [string]$Currency = ($ImportSkus[0].Currency)
  }
}

# Find tenant accounts - but filtered so that we only fetch those with licenses
Write-Host "Finding licensed user accounts..."
[Array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'"  `
  -ConsistencyLevel eventual -CountVariable Records -All -PageSize 999 `
  -Property id, displayName, userPrincipalName, country, department, assignedlicenses, OnPremisesExtensionAttributes, `
  licenseAssignmentStates, createdDateTime, jobTitle, signInActivity, companyName, accountenabled |  `
  Sort-Object DisplayName

If (!($Users)) { 
  Write-Host "No licensed user accounts found - exiting"; break 
}
Else { 
  Write-Host ("{0} Licensed user accounts found - now processing their license data..." -f $Users.Count) 
}

# These are the properties used to create analyses for.
[array]$Departments = $Users.Department | Sort-Object -Unique
[array]$Countries = $Users.Country | Sort-Object -Unique
[array]$CostCenters = $Users.OnPremisesExtensionAttributes.($CostCenterAttribute) | Sort-Object -Unique
[array]$Companies = $Users.CompanyName | Sort-Object -Unique

# Control whether to use the detailed license report information to generate a line-by-line
# report of license assignments to users. This report is useful to detect duplicate licenses and
# to help allocate license costs to operating units within an organization. Set the value to false
# if you don't want to generate the detailed report.
$DetailedCompanyAnalyis = $true

$OrgName = (Get-MgOrganization).DisplayName

# Current subscriptions in the tenant. We use this table to remove expired licenses from the calculation
[array]$CurrentSubscriptions = Get-MgSubscribedSku
$CurrentSubscriptionsHash = @{}
ForEach ($S in $CurrentSubscriptions) {
  $CurrentSubscriptionsHash.Add($S.SkuId, $S.SkuPartNumber) 
}

$DuplicateSKUsAccounts = 0; $DuplicateSKULicenses = 0; $LicenseErrorCount = 0
$Report = [System.Collections.Generic.List[Object]]::new()
$DetailedLicenseReport = [System.Collections.Generic.List[Object]]::new()
$i = 0
[float]$TotalUserLicenseCosts = 0
[float]$TotalBoughtLicenseCosts = 0

ForEach ($User in $Users) {
  $UnusedAccountWarning = "OK"; $i++; $UserCosts = 0
  $ErrorMsg = ""; $LastLicenseChange = ""
  Write-Host ("Processing account {0} {1}/{2}" -f $User.UserPrincipalName, $i, $Users.Count)
  If ([string]::IsNullOrWhiteSpace($User.licenseAssignmentStates) -eq $False) {
    # Only process account if it has some licenses
    [array]$LicenseInfo = $Null; [array]$DisabledPlans = $Null; 
    #  Find out if any of the user's licenses are assigned via group-based licensing
    [array]$GroupAssignments = $User.licenseAssignmentStates | `
      Where-Object { $null -ne $_.AssignedByGroup -and $_.State -eq "Active" }
    #  Find out if any of the user's licenses are assigned via group-based licensing have an error
    [array]$GroupErrorAssignments = $User.licenseAssignmentStates | `
      Where-Object { $Null -ne $_.AssignedByGroup -and $_.State -eq "Error" }
    [array]$GroupLicensing = $Null
    # Find out when the last license change was made
    If ([string]::IsNullOrWhiteSpace($User.licenseAssignmentStates.lastupdateddatetime) -eq $False) {
      $LastLicenseChange = Get-Date(($user.LicenseAssignmentStates.lastupdateddatetime | Measure-Object -Maximum).Maximum) -format g
    }
    # Figure out the details of group-based licensing assignments if any exist
    ForEach ($G in $GroupAssignments) {
      $GroupName = (Get-MgGroup -GroupId $G.AssignedByGroup).DisplayName
      $GroupProductName = $SkuHashTable[$G.SkuId]
      $GroupLicensing += ("{0} assigned from {1}" -f $GroupProductName, $GroupName)
    }
    ForEach ($G in $GroupErrorAssignments) {
      $GroupName = (Get-MgGroup -GroupId $G.AssignedByGroup).DisplayName
      $GroupProductName = $SkuHashTable[$G.SkuId]
      $ErrorMsg = $G.Error
      $LicenseErrorCount++
      $GroupLicensing += ("{0} assigned from {1} BUT ERROR {2}!" -f $GroupProductName, $GroupName, $ErrorMsg)
    }
    $GroupLicensingAssignments = $GroupLicensing -Join ", "

    #  Find out if any of the user's licenses are assigned via direct licensing
    [array]$DirectAssignments = $User.licenseAssignmentStates | `
      Where-Object { $null -eq $_.AssignedByGroup -and $_.State -eq "Active" }

    # Figure out details of direct assigned licenses
    [array]$UserLicenses = $User.AssignedLicenses
    ForEach ($License in $DirectAssignments) {
      If ($SkuHashTable.ContainsKey($License.SkuId) -eq $True) {
        # We found a match in the SKU hash table
        $LicenseInfo += $SkuHashTable.Item($License.SkuId) 
      } Else {
        # Nothing found in the SKU hash table, so output the SkuID
        $LicenseInfo += $License.SkuId
      }
    }

    # Report any disabled service plans in licenses
    $License = $UserLicenses | Where-Object { -not [string]::IsNullOrWhiteSpace($_.DisabledPlans) }
    # Check if disabled service plans in a license
    ForEach ($DisabledPlan in $License.DisabledPlans) {
      # Try and find what service plan is disabled
      If ($ServicePlanHashTable.ContainsKey($DisabledPlan) -eq $True) {
        # We found a match in the Service Plans hash table
        $DisabledPlans += $ServicePlanHashTable.Item($DisabledPlan) 
      }
      Else {
        # Nothing doing, so output the Service Plan ID
        $DisabledPlans += $DisabledPlan 
      }
    } # End ForEach disabled plans

    # Detect if any duplicate licenses are assigned (direct and group-based)
    # Build a list of assigned SKUs
    $SkuUserReport = [System.Collections.Generic.List[Object]]::new()
    ForEach ($S in $DirectAssignments) {
      If ($CurrentSubscriptionsHash[$S.SkuId]) {
        $ReportLine = [PSCustomObject][Ordered]@{ 
          User        = $User.Id
          Name        = $User.DisplayName 
          Sku         = $S.SkuId
          Method      = "Direct"  
          Country     = $User.Country
          Department  = $User.Department
          Company     = $User.CompanyName
        }
        $SkuUserReport.Add($ReportLine)
      }
    }
    ForEach ($S in $GroupAssignments) {
      If ($CurrentSubscriptionsHash[$S.SkuId]) {
        $ReportLine = [PSCustomObject][Ordered]@{ 
          User        = $User.Id
          Name        = $User.DisplayName
          Sku         = $S.SkuId
          Method      = "Group" 
          Country     = $User.Country
          Department  = $User.Department
          Company     = $User.CompanyName
        }
        $SkuUserReport.Add($ReportLine)
      }
    }

    # Check if any duplicates exist
    [array]$DuplicateSkus = $SkuUserReport | Group-Object Sku | `
      Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Name

    # If duplicates exist, resolve their SKU IDs into Product names and generate a warning for the report
    [string]$DuplicateWarningReport = "N/A"
    If ($DuplicateSkus) {
      [array]$DuplicateSkuNames = $Null
      $DuplicateSKUsAccounts++
      $DuplicateSKULicenses = $DuplicateSKULicenses + $DuplicateSKUs.Count
      ForEach ($DS in $DuplicateSkus) {
        $SkuName = $SkuHashTable[$DS]
        $DuplicateSkuNames += $SkuName
      }
      $DuplicateWarningReport = ("Warning: Duplicate licenses detected for: {0}" -f ($DuplicateSkuNames -join ", "))
    }
  } Else { 
      $UnlicensedAccounts++
  }
  # Figure out the last time the account signed in. This is important for detecting unused accounts
  $LastSignIn = $User.SignInActivity.LastSignInDateTime
  $LastNonInteractiveSignIn = $User.SignInActivity.LastNonInteractiveSignInDateTime

  If (-not $LastSignIn -and -not $LastNonInteractiveSignIn) {
      $DaysSinceLastSignIn = "Unknown"
      $UnusedAccountWarning = ("Unknown last sign-in for account")
      $LastAccess = "Unknown"
  } Else {
    # Get the newest date, if both dates contain values
    If ($LastSignIn -and $LastNonInteractiveSignIn) {
      If ($LastSignIn -gt $LastNonInteractiveSignIn) {
          $CompareDate = $LastSignIn
      } Else {
          $CompareDate = $LastNonInteractiveSignIn
      }
    } Elseif ($LastSignIn) {
        # Only $LastSignIn has a value
        $CompareDate = $LastSignIn
    } Else {
        # Only $LastNonInteractiveSignIn has a value
        $CompareDate = $LastNonInteractiveSignIn
    }

    $DaysSinceLastSignIn = ($RunDate - $CompareDate).Days
    $LastAccess = Get-Date($CompareDate) -format g
    If ($DaysSinceLastSignIn -gt 60) { 
        $UnusedAccountWarning = ("Account unused for {0} days - check!" -f $DaysSinceLastSignIn) 
    }
  }

  $AccountCreatedDate = $null
  If ($User.CreatedDateTime) {
      $AccountCreatedDate = Get-Date($User.CreatedDateTime) -format 'dd-MMM-yyyy HH:mm' 
  }

  # If cost center reporting is enabled, extract the cost center for the user
  [string]$CostCenter = $Null
  If ($CostCenterAttribute) {
    $CostCenter = $User.OnPremisesExtensionAttributes.($CostCenterAttribute)
  }

  # Report information
  [string]$DisabledPlans = $DisabledPlans -join ", " 
  [string]$LicenseInfo = $LicenseInfo -join (", ")

  If ($User.AccountEnabled -eq $False) {
      $AccountStatus = "Disabled" 
  } Else {
      $AccountStatus = "Enabled"
    }

  If ($PricingInfoAvailable) { 
    # Output report line with pricing info
    [float]$UserCosts = Get-LicenseCosts -Licenses $UserLicenses.SkuId
    $TotalUserLicenseCosts = $TotalUserLicenseCosts + $UserCosts
    $ReportLine = [PSCustomObject][Ordered]@{  
      User                       = $User.DisplayName
      UPN                        = $User.UserPrincipalName
      Country                    = $User.Country
      Department                 = $User.Department
      Title                      = $User.JobTitle
      Company                    = $User.companyName
      "Direct assigned licenses" = $LicenseInfo
      "Disabled Plans"           = $DisabledPlans.Trim() 
      "Group based licenses"     = $GroupLicensingAssignments
      "Annual License Costs"     = ("{0} {1}" -f $Currency, ($UserCosts.toString('F2')))
      "Last license change"      = $LastLicenseChange
      "Account created"          = $AccountCreatedDate
      "Last Signin"              = $LastAccess
      "Days since last signin"   = $DaysSinceLastSignIn
      "Duplicates detected"      = $DuplicateWarningReport
      Status                     = $UnusedAccountWarning
      "Account status"           = $AccountStatus
      UserCosts                  = $UserCosts  
      'Cost Center'              = $CostCenter
    }
  } Else { 
    # No pricing information
    $ReportLine = [PSCustomObject][Ordered]@{  
      User                       = $User.DisplayName
      UPN                        = $User.UserPrincipalName
      Country                    = $User.Country
      Department                 = $User.Department
      Title                      = $User.JobTitle
      Company                    = $User.companyName
      "Direct assigned licenses" = $LicenseInfo
      "Disabled Plans"           = $DisabledPlans.Trim() 
      "Group based licenses"     = $GroupLicensingAssignments
      "Last license change"      = $LastLicenseChange
      "Account created"          = $AccountCreatedDate
      "Last Signin"              = $LastAccess
      "Days since last signin"   = $DaysSinceLastSignIn
      "Duplicates detected"      = $DuplicateWarningReport
      Status                     = $UnusedAccountWarning
      "Account status"           = $AccountStatus
    }
  }  
  $Report.Add($ReportLine)

  # Populate the detailed license assignment report
  $SkuUserReport = $SkuUserReport | Sort-Object Sku -Unique
  ForEach ($Item in $SkuUserReport) {
    $SkuReportLine = [PSCustomObject][Ordered]@{  
      User        = $Item.User
      Name        = $Item.name
      Sku         = $Item.Sku
      SkuName     = ($SkuHashTable[$Item.Sku])
      Method      = $Item.Method
      Country     = $Item.Country
      Department  = $Item.Department
      Company     = $Item.Company    
    }
  $DetailedLicenseReport.Add($SkuReportLine)
  }
} # End ForEach Users

$UnderusedAccounts = $Report | Where-Object { $_.Status -ne "OK" }
$PercentUnderusedAccounts = ($UnderUsedAccounts.Count / $Report.Count).toString("P")

# This code grabs the SKU summary for the tenant and uses the data to create a SKU summary usage segment for the HTML report
$SkuReport = [System.Collections.Generic.List[Object]]::new()
[array]$SkuSummary = Get-MgSubscribedSku | Select-Object SkuId, ConsumedUnits, PrepaidUnits
$SkuSummary = $SkuSummary | Where-Object { $_.ConsumedUnits -ne 0 }
ForEach ($S in $SkuSummary) {
  $SKUCost = Get-LicenseCosts -Licenses $S.SkuId
  $SkuDisplayName = $SkuHashtable[$S.SkuId]
  If ($S.PrepaidUnits.Enabled -le $S.ConsumedUnits ) {
    $BoughtUnits = $S.ConsumedUnits 
  } Else {
    $BoughtUnits = $S.PrepaidUnits.Enabled
  }
  If ($PricingInfoAvailable) {
    $SKUTotalCost = ($SKUCost * $BoughtUnits)
    $SkuReportLine = [PSCustomObject][Ordered]@{  
      "SKU Id"                = $S.SkuId
      "SKU Name"              = $SkuDisplayName 
      "Units Used"            = $S.ConsumedUnits 
      "Units Purchased"       = $BoughtUnits
      "Annual license costs"  = $SKUTotalCost
      "Annual licensing cost" = ("{0} {1}" -f $Currency, ('{0:N2}' -f $SKUTotalCost))
    }
  } Else {
    $SkuReportLine = [PSCustomObject][Ordered]@{  
      "SKU Id"          = $S.SkuId
      "SKU Name"        = $SkuDisplayName 
      "Units Used"      = $S.ConsumedUnits 
      "Units Purchased" = $BoughtUnits
    }
  }
  $SkuReport.Add($SkuReportLine) 
  $TotalBoughtLicenseCosts = $TotalBoughtLicenseCosts + $SKUTotalCost
}

If ($PricingInfoAvailable) {
  $AverageCostPerUser = ($TotalUserLicenseCosts / $Users.Count)
  $AverageCostPerUserOutput = ("{0} {1}" -f $Currency, ('{0:N2}' -f $AverageCostPerUser))
  $TotalUserLicenseCostsOutput = ("{0} {1}" -f $Currency, ('{0:N2}' -f $TotalUserLicenseCosts))
  $TotalBoughtLicenseCostsOutput = ("{0} {1}" -f $Currency, ('{0:N2}' -f $TotalBoughtLicenseCosts))
  $PercentBoughtLicensesUsed = ($TotalUserLicenseCosts / $TotalBoughtLicenseCosts).toString('P')
  $SkuReport = $SkuReport | Sort-Object "Annual license costs" -Descending
} Else {
  $SkuReport = $SkuReport | Sort-Object "SKU Name" -Descending
}

  # Generate the department analysis
  $DepartmentReport = [System.Collections.Generic.List[Object]]::new()
  ForEach ($Department in $Departments) {
    [array]$DepartmentRecords = $Report | Where-Object {$_.Department -eq $Department}
    $DepartmentReportLine = [PSCustomObject][Ordered]@{
      Department    = $Department
      Accounts      = $DepartmentRecords.count
      Costs       = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($DepartmentRecords | Measure-Object UserCosts -Sum).Sum))
      AverageCost = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($DepartmentRecords | Measure-Object UserCosts -Average).Average))
    } 
    $DepartmentReport.Add($DepartmentReportLine)
  }
  $DepartmentHTML = $DepartmentReport | ConvertTo-HTML -Fragment
  # Anyone without a department?
  [array]$NoDepartment = $Report | Where-Object { $null -eq $_.Department }
  If ($NoDepartment) {
    $NoDepartmentCosts = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($NoDepartment | Measure-Object UserCosts -Sum).Sum))
  } Else {
    $NoDepartmentCosts = "Zero"
  }

  # Generate the country analysis
  $CountryReport = [System.Collections.Generic.List[Object]]::new()
  ForEach ($Country in $Countries) {
    [array]$CountryRecords = $Report | Where-Object {$_.Country -eq $Country}
    $CountryReportLine = [PSCustomObject][Ordered]@{
      Country     = $Country
      Accounts    = $CountryRecords.count
      Costs       = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CountryRecords | Measure-Object UserCosts -Sum).Sum))
      AverageCost = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CountryRecords | Measure-Object UserCosts -Average).Average))
    } 
    $CountryReport.Add($CountryReportLine)
  }
  $CountryHTML = $CountryReport | ConvertTo-HTML -Fragment
  # Anyone without a country?
  [array]$NoCountry = $Report | Where-Object { $null -eq $_.Country }
  If ($NoCountry) {
    $NoCountryCosts = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($NoCountry | Measure-Object UserCosts -Sum).Sum))
  } Else {
    $NoCountryCosts = "Zero"
  }

  # Generate cost center analysis
  If ($PricingInfoAvailable -and $null -ne $CostCenterAttribute) { 
    $CostCenterReport = [System.Collections.Generic.List[Object]]::new()
    ForEach ($CostCenter in $CostCenters) {
      [array]$CostCenterRecords = $Report | Where-Object {$_.'Cost Center' -eq $CostCenter}
      $CostCenterReportLine = [PSCustomObject][Ordered]@{
        'Cost Center' = $CostCenter
        Accounts      = $CostCenterRecords.count
        Costs         = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CostCenterRecords | Measure-Object UserCosts -Sum).Sum))
        AverageCost   = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CostCenterRecords | Measure-Object UserCosts -Average).Average))
      } 
      $CostCenterReport.Add($CostCenterReportLine)
    }
    $CostCenterHTML = $CostCenterReport | ConvertTo-HTML -Fragment
    # Anyone without a cost center?
    [array]$NoCostCenter = $Report | Where-Object { $null -eq $_.'Cost Center' }
    If ($NoCostCenter) {
        $NoCostCenterCosts = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($NoCostCenter | Measure-Object UserCosts -Sum).Sum))
    } Else {
        $NoCostCenterCosts = "Zero"
    }
  }

  # Generate the company analysis
  $CompanyReport = [System.Collections.Generic.List[Object]]::new()
  ForEach ($Company in $Companies) {
    [array]$CompanyRecords = $Report | Where-Object {$_.Company -eq $Company}
    $CompanyReportLine = [PSCustomObject][Ordered]@{
      Company     = $Company
      Accounts    = $CompanyRecords.count
      Costs       = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CompanyRecords | Measure-Object UserCosts -Sum).Sum))
      AverageCost = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CompanyRecords | Measure-Object UserCosts -Average).Average))
    }
    $CompanyReport.Add($CompanyReportLine)
  }
  $CompanyHTML = $CompanyReport | ConvertTo-HTML -Fragment
  # Anyone without an assigned company?
  [array]$NoCompany = $Report | Where-Object { $null -eq $_.Company }
  If ($NoCompany) {
    $NoCompanyCosts = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($NoCompany | Measure-Object UserCosts -Sum).Sum))
  } Else {
    $NoCompanyCosts = "Zero"
  }

  $CompanyAnalysisHTML = $null
  # Detailed company analysis - example of breaking down costs by SKU for each company
  ForEach ($Company in $Companies) {
    [array]$CompanyAssignments = $DetailedLicenseReport | Where-Object {$_.Company -eq $Company}
    $CompanyAnalysisHTML = $CompanyAnalysisHTML + ("<h2>Company Analysis: Product Licenses for {0}</h2><p>" -f $Company)
    [array]$Skus = $CompanyAssignments.Sku | Sort-Object -Unique
      
    ForEach ($Sku in $Skus) {
      [float]$AnnualCost = 0; [float]$AnnualCostLicense = 0; $AnnualCostLicenseFormatted = $null
      $SkuHTMLFooter = $null; [float]$AnnualCost = $null; $AnnualCostLicense = $null
      $SkuHeader = ("<h3>{0}</h3>" -f $SkuHashTable[$Sku])
      $AssignedSkus = $CompanyAssignments | Where-Object {$_.Sku -eq $Sku} | Select-Object Sku, Name, SkuName, Country, Department, Company 
      $LicenseCostSKU = $PricingHashTable[$Sku]
      [float]$LicenseCostCents = [float]$LicenseCostSKU * 100
      If ($LicenseCostCents -gt 0) {
        # Compute annual cost for the license
        [float]$AnnualCost = $LicenseCostCents * 12
        # Compute cost for this SKU assigned to this company
        $AnnualCostLicense = ($AnnualCost * $AssignedSkus.count)/100
        $AnnualCostLicenseFormatted = ("{0} {1}" -f $Currency, ('{0:N2}' -f $AnnualCostLicense))
      } Else {
        $AnnualCostLicenseFormatted = ("{0} {1}" -f $Currency, ('{0:N2}' -f 0))
      }
      # Report the set of people assigned this SKU      
      $AssignedSkusHTML = $AssignedSkus | ConvertTo-HTML -fragment
      $CompanySKUDetailHTML = $SkuHeader + "<p>" + $AssignedSkusHTML + "<p>"
      $SkuHTMLFooter = ("<p>Annual cost for {0} license(s): {1}</p>" -f $AssignedSKUs.count, $AnnualCostLicenseFormatted)
      $CompanyAnalysisHTML = $CompanyAnalysisHTML + "</p>" + $CompanySKUDetailHTML + $SkuHTMLFooter
    }
    $CompanyAnalysisHTML =  "<p>" + $CompanyAnalysisHTML + "</p>" + $CompanyHTMLFooter
} 

# Inactive user accounts - these are accounts that have never signed in or whose last sign-in was more than 60 days ago  $InactiveUserAccounts = $Report | Where-Object {($_."Days since last signin" -ge 60) -or ($_.'Days since last signin' -eq "Unknown")}
$InactiveUserAccountsCost = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($InactiveUserAccounts | Measure-Object UserCosts -Sum).Sum)) # Disabled user accounts
$DisabledUserAccounts = $Report | Where-Object { $_."Account status" -eq "disabled" }
$DisabledUserAccountsCost = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($DisabledUserAccounts | Measure-Object UserCosts -Sum).Sum)) 

# Cost spans for license comparison
$LowCost = $AverageCostPerUser * 0.8
$MediumCost = $AverageCostPerUser

# Create the HTML report
$HtmlHead = "<html>
	  <style>
	  BODY{font-family: Arial; font-size: 8pt;}
	  H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	  H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	  H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	  TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	  TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	  TD{border: 1px solid #969595; padding: 5px; }
    TD.disabledaccount{background: #FFC0CB;}
	  TD.higherlicensecost{background: #E3242B;}
    TD.averagelicensecost{background: #FFFF00;}
    TD.lowerlicensecost{background: #4cff47;}
    TD.inactiveaccount{background: #FF474C;}
    TD.duplicatelicenses{background: #F8FF00}
	  </style>
	  <body>
           <div align=center>
           <p><h1>Microsoft 365 Licensing Report</h1></p>
           <p><h2><b>For the " + $Orgname + " tenant</b></h2></p>
           <p><h3>Generated: " + $ReportRunDate + "</h3></p></div>"

If ($PricingInfoAvailable) {
  $HtmlBody1 = $Report | Select-Object User, UPN, Country, Department, Title, Company, "Direct assigned licenses", "Disabled Plans", "Group based licenses", "Annual License Costs", "Last license change", "Account created", "Last Signin", "Days since last signin", "Duplicates detected", Status, "Account status" | ConvertTo-Html -Fragment
  # Create an attribute class to use, name it, and append to the XML table attributes
  [xml]$XML = $HTMLBody1
  $TableClass = $XML.CreateAttribute("class")
  $TableClass.Value = "State"
  $XML.table.Attributes.Append($TableClass) | Out-Null
  # Conditional formatting for the table rows.  
  ForEach ($TableRow in $XML.table.SelectNodes("tr")) {
    # each TR becomes a member of class "tablerow"
    $TableRow.SetAttribute("class","tablerow")
    # Make sure that we can deal with the days since last sign-in property, which could be blank for new accounts
    Try {
      [int]$DaysSinceLastSignIn = $TableRow.td[13]
    } Catch {
      # This Catch block will be executed if the value of the DaysSinceLastSignIn property is not an integer
      # like "Unknown"
      $DaysSinceLastSignIn = 99999
    }
    # Level of license cost
    Try {
      $UserUPN = $TableRow.td[1]
    }
    Catch { # ignore error in fetching user principal name for header lines
      Continue
    } 
    # Fetch total license cost without currency prefix
    $Cost = $Report.Where{$_.UPN -eq $UserUPN} | Select-Object -ExpandProperty UserCosts
    Switch ($Cost) {
      {$Cost -gt $MediumCost} { 
        $TableRow.SelectNodes("td")[9].SetAttribute("class","higherlicensecost") 
        #Write-Host ("High cost for {0}" -f $TableRow.td[1])
      }
      {($Cost -gt $LowCost) -and ($Cost -le $MediumCost)} { 
        $TableRow.SelectNodes("td")[9].SetAttribute("class","averagelicensecost") 
        #Write-Host ("Medium cost for {0}" -f $TableRow.td[1])
      }
      {$Cost -le $LowCost} { 
        $TableRow.SelectNodes("td")[9].SetAttribute("class","lowerlicensecost") 
        #Write-Host ("Low cost for {0}" -f $TableRow.td[1])
      }
    }
    # Highlight accounts that haven't signed in for more than 90 days 
    If (($TableRow.td) -and ($DaysSinceLastSignIn -ge 90))  {
        ## tag the TD as inactive
      $TableRow.SelectNodes("td")[13].SetAttribute("class","inactiveaccount")
    }
    # If duplicate licenses are detected
    If (($TableRow.td) -and ([string]$TableRow.td[14] -ne 'N/A'))  {
        # tag the TD with the color for duplicate licenses
        # Write-Host "Detected duplicate licenses for $($TableRow.td[1])"
      $TableRow.SelectNodes("td")[14].SetAttribute("class","duplicatelicenses")
    }
    # If row has the account status set to disabled
    If (($TableRow.td) -and ([string]$TableRow.td[16] -eq 'disabled'))  {
      ## tag the TD with the color for a disabled account
      $TableRow.SelectNodes("td")[16].SetAttribute("class","disabledaccount")
    }
  }
    # Wrap the output table with a div tag
    $HTMLBody1 = [string]::Format('<div class="tablediv">{0}</div>',$XML.OuterXml)
} Else {
  $HtmlBody1 = $Report | ConvertTo-Html -Fragment
}
$HtmlBody1 = $HtmlBody1 + "<p>Report created for: " + $OrgName + "</p><p>Created: " + $ReportRunDate + "<p>" 

$HtmlBody2 = $SkuReport | Select-Object "SKU Id", "SKU Name", "Units used", "Units purchased", "Annual licensing cost" | ConvertTo-Html -Fragment
$HtmlSkuSeparator = "<p><h2>Product License Distribution</h2></p>"

$HtmlTail = "<p></p>"

# Add first set of cost analysis if pricing information is available
If ($PricingInfoAvailable) {
  $HTMLTail = $HTMLTail + "<h2>Licensing Cost Analysis</h2>" +
  "<p>Total licensing cost for tenant:              " + $TotalBoughtLicenseCostsOutput + "</p>" +
  "<p>Total cost for assigned licenses:             " + $TotalUserLicenseCostsOutput + "</p>" +
  "<p>Percent bought licenses assigned to users:    " + $PercentBoughtLicensesUsed + "</p>" +
  "<p>Average licensing cost per user:              " + $AverageCostPerUserOutput + "</p>" +
  "<p><h2>License Costs by Country</h2></p>         " + $CountryHTML +
  "<p>License costs for users without a country:    " + $NoCountryCosts +
  "<p><h2>License Costs by Department</h2></p>      " + $DepartmentHTML +
  "<p>License costs for users without a department: " + $NoDepartmentCosts +
  "<p><h2>License Costs by Company</h2></p>         " + $CompanyHTML +
  "<p>License costs for users without a department: " + $NoCompanyCosts

  If ($DetailedCompanyAnalyis) {
    $HTMLTail = $HTMLTail + $CompanyAnalysisHTML
  }
}

# Add cost center information if we've been asked to generate it
If ($CostCenterAttribute) {
  $HTMLTail = $HtmlTail + "<h2>Cost Center Analysis</h2><p></p>" + $CostCenterHTML + "<p></p>" +
  "<p>License costs for users without a cost center:    " + $NoCostCenterCosts 
}

# Add the second part of the cost analysis if pricing information is available
If ($PricingInfoAvailable) {
  $HTMLTail = $HTMLTail +
  "<p><h2>Inactive User Accounts</h2></p>" +
  "<p>Number of inactive user accounts:             " + $InactiveUserAccounts.Count + "</p>" +
  "<p>Names of inactive accounts:                   " + ($InactiveUserAccounts.User -join ", ") + "</p>" +
  "<p>Cost of inactive user accounts:               " + $InactiveUserAccountsCost + "</p>" +
  "<p><h2>Disabled User Accounts</h2></p>" +
  "<p>Number of disabled accounts:                  " + $DisabledUserAccounts.Count + "</p>" +
  "<p>Names of disabled accounts:                   " + ($DisabledUserAccounts.User -join ", ") + "</p>" +
  "<p>Cost of disabled user accounts:               " + $DisabledUserAccountsCost + "</p>"
}

$HTMLTail = $HTMLTail +
"<p>-----------------------------------------------------------------------------------------------------------------------------</p>" +  
"<p>Number of licensed user accounts found:    " + $Report.Count + "</p>" +
"<p>Number of underused user accounts found:   " + $UnderUsedAccounts.Count + "</p>" +
"<p>Percent underused user accounts:           " + $PercentUnderusedAccounts + "</p>" +
"<p>Accounts detected with duplicate licenses: " + $DuplicateSKUsAccounts + "</p>" +
"<p>Count of duplicate licenses:               " + $DuplicateSKULicenses + "</p>" +
"<p>Count of errors:                           " + $LicenseErrorCount + "</p>" +
"<p>-----------------------------------------------------------------------------------------------------------------------------</p>"
          
$HTMLTail = $HTMLTail + "<p>Microsoft 365 Licensing Report<b> " + $Version + "</b></p>"	
$HtmlReportFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Microsoft 365 Licensing Report.html"
$HtmlReport = $Htmlhead + $Htmlbody1 + $HtmlSkuSeparator + $HtmlBody2 + $Htmltail
$HtmlReport | Out-File $HtmlReportFile -Encoding UTF8

# Generate output report files

If (Get-Module ImportExcel -ListAvailable) {
  $ExcelGenerated = $True
  Import-Module ImportExcel -ErrorAction SilentlyContinue
  $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Microsoft 365 Licensing Report.xlsx"
  $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Microsoft 365 Licensing Report" -Title ("Microsoft 365 Licensing Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "Microsoft365LicensingReport" 
  If ($DetailedCompanyAnalyis) {
    $DetailedExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Detailed Microsoft 365 Licensing Report.xlsx"
    $DetailedLicenseReport | Export-Excel -Path $DetailedExcelOutputFile -WorksheetName "Detailed Microsoft 365 Licensing" `
    	  -Title ("Detailed Microsoft 365 Licensing Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "DetailedMicrosoft365LicensingReport" 
  }
  } Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Microsoft 365 Licensing Report.CSV"
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}

Write-Host ""
Write-Host "Microsoft 365 Licensing Report complete"
Write-Host "---------------------------------------"
Write-Host ""
Write-Host ("An HTML report is available in {0}" -f $HtmlReportFile)
If ($ExcelGenerated) {
  Write-Host ("An Excel report is available in {0}" -f $ExcelOutputFile)
} Else {
  Write-Host ("A CSV report is available in {0}" -f $CSVOutputFile)
}

Disconnect-MgGraph

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
