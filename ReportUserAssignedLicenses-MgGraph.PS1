# ReportUserAssignedLicenses-MgGraph.PS1
# Create a report of licenses assigned to Azure AD user accounts using the Microsoft Graph PowerShell SDK cmdlets
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportUserAssignedLicenses-MgGraph.PS1
# See https://practical365.com/create-licensing-report-microsoft365-tenant/ for an article describing how to run the report

# V1.1 27-Sep-2022 Add sign in data for users and calculate how long it's been since they signed in and used a license.
# V1.2 23-Nov-2022 Added SKU usage summary to HTML report
# V1.3 29-Sep-2023 Added support for group-based licensing

[datetime]$RunDate = Get-Date
$Version = "1.3"
$CSVOutputFile = "c:\temp\Microsoft365LicensesReport.CSV"
$ReportFile = "c:\temp\Microsoft365LicensesReport.html"
$UnlicensedAccounts = 0

# Connect to the Graph, specifing the tenant and profile to use - Add your tenant identifier here
Connect-MgGraph -Scope "Directory.AccessAsUser.All, Directory.Read.All, AuditLog.Read.All"

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

# Find tenant accounts - but filtered so that we only fetch those with licenses
Write-Host "Finding licensed user accounts..."
[Array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
    -ConsistencyLevel eventual -CountVariable Records -All `
    -Property id, displayName, userPrincipalName, country, department, assignedlicenses, `
       LicenseAssignmentStates, createdDateTime, jobTitle, signInActivity | `
    Sort-Object DisplayName

If (!($Users)) { Write-Host "No licensed user accounts found - exiting"; break }
  Else { Write-Host ("{0} Licensed user accounts found - now processing their license data..." -f $Users.Count) }

$Report = [System.Collections.Generic.List[Object]]::new()
$i = 0
ForEach ($User in $Users) {
 $UnusedAccountWarning = "OK"; $i++
 Write-Host ("Processing account {0} {1}/{2}" -f $User.UserPrincipalName, $i, $Users.Count)
 If ([string]::IsNullOrWhiteSpace($User.AssignedLicenses) -eq $False) { # Only process account if it has some licenses
  [array]$LicenseInfo = $Null; [array]$DisabledPlans = $Null
  ForEach ($License in $User.AssignedLicenses) {
     If ($SkuHashTable.ContainsKey($License.SkuId) -eq $True) { # We found a match in the SKU hash table
        $LicenseInfo += $SkuHashTable.Item($License.SkuId) 
   } Else { # Nothing doing, so output the SkuID
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

  # Calculate how long it's been since someone signed in
  If ([string]::IsNullOrWhiteSpace($User.SignInActivity.LastSignInDateTime) -eq $False) {
    [datetime]$LastSignInDate = $User.SignInActivity.LastSignInDateTime
    $DaysSinceLastSignIn = ($RunDate - $LastSignInDate).Days
    $LastAccess = Get-Date($User.SignInActivity.LastSignInDateTime) -format g
    If ($DaysSinceLastSignIn -gt 60) { $UnusedAccountWarning = ("Account unused for {0} days - check!" -f $DaysSinceLastSignIn) }
  }
  Else {
    $DaysSinceLastSignIn = "Unknown"
    $UnusedAccountWarning = ("Unknown last sign-in for account")
    $LastAccess = "Unknown"
  }

  $AccountCreatedDate = $Null
  If ($User.CreatedDateTime) {
     $AccountCreatedDate = Get-Date($User.CreatedDateTime) -format g 
  }
  
#  Find out if any of the user's licenses are assigned via group-based licensing
[array]$GroupAssignments = $User.licenseAssignmentStates | Where-Object {$Null -ne $_.AssignedByGroup}
[array]$GroupLicensing = $Null
# Figure out group-based licensing assignments if any exist
ForEach ($G in $GroupAssignments) {
       $GroupName = (Get-MgGroup -GroupId $G.AssignedByGroup).DisplayName
       $GroupProductName = $SkuHashTable[$G.SkuId]
       $GroupLicensing += ("{0} assigned from {1}" -f $GroupProductName, $GroupName)
}
$GroupLicensingAssignments = $GroupLicensing -Join ", "

  # Report information
  [string]$DisabledPlans = $DisabledPlans -join ", " 
  [string]$LicenseInfo = $LicenseInfo -join (", ")
  $ReportLine = [PSCustomObject][Ordered]@{  
       User                     = $User.DisplayName
       UPN                      = $User.UserPrincipalName
       Country                  = $User.Country
       Department               = $User.Department
       Title                    = $User.JobTitle
       Licenses                 = $LicenseInfo
       "Disabled Plans"         = $DisabledPlans 
       "Group based licensing"  = $GroupLicensingAssignments
       "Account created"        = $AccountCreatedDate
       "Last Signin"            = $LastAccess
       "Days since last signin" = $DaysSinceLastSignIn
       Status                   = $UnusedAccountWarning
  }
  $Report.Add($ReportLine)
  } #end If account is licensed
  Else { $UnlicensedAccounts++}
} # End ForEach Users

$UnderusedAccounts = $Report | Where-Object {$_.Status -ne "OK"}
$PercentUnderusedAccounts = ($UnderUsedAccounts.Count/$Report.Count).toString("P")

# This code grabs the SKU summary for the tenant and uses the data to create a SKU summary usage segment for the HTML report
$SkuReport = [System.Collections.Generic.List[Object]]::new()
[array]$SkuSummary = Get-MgSubscribedSku | Select-Object SkuId, ConsumedUnits
$SkuSummary = $SkuSummary | Where-Object {$_.ConsumedUnits -ne 0 }
ForEach ($S in $SkuSummary) {
   $SkuDisplayName = $SkuHashtable[$S.SkuId]
   $SkuReportLine = [PSCustomObject][Ordered]@{  
     "SKU Id"      = $S.SkuId
     "SKU Name"    = $SkuDisplayName 
     "Units Used"  = $S.ConsumedUnits }
  $SkuReport.Add($SkuReportLine) 
}
$SkuReport = $SkuReport | Sort-Object "Units Used" -Descending

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
           <p><h3>Generated: " + (Get-Date($RunDate) -format 'dd-MMM-yyyy hh:mm') + "</h3></p></div>"

$Htmlbody1 = $Report | ConvertTo-Html -Fragment
$Htmlbody2 = $SkuReport | ConvertTo-Html -Fragment
$HtmlSkuSeparator = "<p><h2>Product License Distribution</h2></p>"
$Htmltail = "<p>Report created for: " + $OrgName + "</p>" +
             "<p>Created: " + (Get-Date($RunDate) -format 'dd-MMM-yyy hh:mm') + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of licensed user accounts found:    " + $Report.Count + "</p>" +
             "<p>Number of underused user accounts found:   " + $UnderUsedAccounts.Count + "</p>" +
             "<p>Percent underused user accounts:           " + $PercentUnderusedAccounts + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Microsoft 365 Licensing Report<b> " + $Version + "</b>"	

$htmlreport = $Htmlhead + $Htmlbody1 + $HtmlSkuSeparator + $HtmlBody2 + $Htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8

$Report | Export-CSV -NoTypeInformation $CSVOutputFile
Write-Host ""
Write-Host "All done. Output files are" $CSVOutputFile "and" $ReportFile

Disconnect-MgGraph

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
