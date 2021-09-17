# ReportUserAssignedLicenses-MgGraph.PS1
# Create a report of licenses assigned to Azure AD user accounts using the Microsoft Graph PowerShell SDK cmdlets
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportUserAssignedLicenses-MgGraph.PS1

$CreationDate = Get-Date -format g
$Version = "1.0"
$CSVOutputFile = "c:\temp\Microsoft365LicensesReport.CSV"
$ReportFile = "c:\temp\Microsoft365LicensesReport.html"
$UnlicensedAccounts = 0

# Connect to the Graph, specifing the tenant and profile to use - Add your tenant identifier here
Connect-MgGraph -TenantId xxxxxxx-14fc-43a2-9a7a-d2e27f4f3478
Select-MgProfile beta
$Details = Get-MgContext
$Scopes = $Details | Select -ExpandProperty Scopes
$Scopes = $Scopes -Join ", "
$ProfileName = (Get-MgProfile).Name
$OrgName = (Get-MgOrganization).DisplayName
CLS
Write-Host "Microsoft Graph Connection Information"
Write-Host "--------------------------------------"
Write-Host " "
Write-Host ("Connected to Tenant {0} ({1}) as account {2}" -f $Details.TenantId, $OrgName, $Details.Account)
Write-Host "+-------------------------------------------------------------------------------------------------------------------+"
Write-Host ("Profile set as {0}. The following permission scope is defined: {1}" -f $ProfileName, $Scopes)
Write-Host ""

# This step depends on the availability of some CSV files generated to hold information about the product licenses used in the tenant and 
# the service plans in those licenses. See https://github.com/12Knocksinna/Office365itpros/blob/master/CreateCSVFilesForSKUsAndServicePlans.PS1 
# for code to generate the CSVs. After the files are created, you need to edit them to add the display names for the SKUs and plans.
# Build Hash of Skus for lookup so that we report user-friendly display names - you need to create these CSV files from SKU and service plan
# data in your tenant.

If ((Test-Path c:\temp\SkuDataComplete.csv) -eq $False)  {
    Write-Host "Can't find the product data file (c:\temp\SkuDataComplete.csv). Exiting..." ; break }
If ((Test-Path c:\temp\ServicePlanDataComplete.csv) -eq $False) {
   Write-Host "Can't find the serivice plan data file (c:\temp\ServicePlanDataComplete.csv). Exiting..." ; break }
   
$ImportSkus = Import-CSV c:\temp\SkuDataComplete.csv
$ImportServicePlans = Import-CSV c:\temp\ServicePlanDataComplete.csv
$SkuHashTable = @{}
ForEach ($Line in $ImportSkus) { $SkuHashTable.Add([string]$Line.SkuId, [string]$Line.DisplayName) }
$ServicePlanHashTable = @{}
ForEach ($Line2 in $ImportServicePlans) { $ServicePlanHashTable.Add([string]$Line2.ServicePlanId, [string]$Line2.ServicePlanDisplayName) }

$Report = [System.Collections.Generic.List[Object]]::new()

# Find tenant accounts
[Array]$Users = Get-MgUser -Filter "UserType eq 'Member'" -All | Sort DisplayName

ForEach ($User in $Users) {
 If ([string]::IsNullOrWhiteSpace($User.AssignedLicenses) -eq $False) { # Only process account if it has some licenses
  Write-Host "Processing" $User.DisplayName
  [array]$LicenseInfo = $Null; [array]$DisabledPlans = $Null
  ForEach ($License in $User.AssignedLicenses) {
     If ($SkuHashTable.ContainsKey($License.SkuId) -eq $True) { # We found a match in the SKU hash table
        $LicenseInfo += $SkuHashTable.Item($License.SkuId) }
     Else { # Nothing doing, so output the SkuID
        $LicenseInfo += $License }
  # Report any disabled service plans in licenses
   If ([string]::IsNullOrWhiteSpace($License.DisabledPlans) -eq $False ) { # Check if disabled service plans in a license
     ForEach ($DisabledPlan in $License.DisabledPlans) { # Try and find what service plan is disabled
       If ($ServicePlanHashTable.ContainsKey($DisabledPlan) -eq $True) { # We found a match in the Service Plans hash table
          $DisabledPlans += $ServicePlanHashTable.Item($DisabledPlan) }
       Else { # Nothing doing, so output the Service Plan ID
          $DisabledPlans += $DisabledPlan }
      } # End ForEach disabled plans
    } # End if check for disabled plans  
   } # End Foreach Licenses
  # Report information
  [string]$DisabledPlans = $DisabledPlans -join ", " 
  [string]$LicenseInfo = $LicenseInfo -join (", ")
  $ReportLine = [PSCustomObject][Ordered]@{  
       User            = $User.DisplayName
       UPN             = $User.UserPrincipalName
       Country         = $User.Country
       Department      = $User.Department
       Title           = $User.JobTitle
       Licenses        = $LicenseInfo
      "Disabled Plans" = $DisabledPlans }
  $Report.Add($ReportLine)
  } #end If account is licensed
  Else { $UnlicensedAccounts++}
} # End ForEach Users

# Create the HTML report
$htmlhead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.pass{background: #B7EB83;}
	   td.warn{background: #FFF275;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Microsoft 365 License Report</h1></p>
           <p><h2><b>For the " + $Orgname + " tenant</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format g) + "</h3></p></div>"

$htmlbody1 = $Report | ConvertTo-Html -Fragment

$htmltail = "<p>Report created for: " + $OrgName + "</p>" +
             "<p>Created: " + $CreationDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of licensed user accounts found:    " + $Report.Count + "</p>" +
             "<p>Number of unlicensed user accounts found:  " + $UnlicensedAccounts + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Microsoft 365 Licensing Report<b> " + $Version + "</b>"	

$htmlreport = $htmlhead + $htmlbody1 + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8

$Report | Export-CSV -NoTypeInformation $CSVOutputFile
Write-Host ""
Write-Host "All done. Output files are" $CSVOutputFile "and" $ReportFile


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
