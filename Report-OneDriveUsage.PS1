# Report-OneDriveUsage.PS1
# Example of generating a report about OneDrive for Business storage consumption using Microsoft 365 usage reports
# V1.0 21-Feb-2024
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-OneDriveUsage.PS1

Connect-MgGraph -NoWelcome -Scopes User.Read.All, Reports.Read.All, ReportSettings.ReadWrite.All

$ObscureFlag = $false
$CSVOutputFile = "c:\temp\OneDriveSiteConsumption.CSV"
$TempExportFile = "c:\temp\TempExportFile.CSV"
If (Get-Item $TempExportFile -ErrorAction SilentlyContinue) {
    Remove-Item $TempExportFile
}

# Check if the tenant has obscured real names for reports - see https://office365itpros.com/2022/09/09/graph-usage-report-tips/
If ((Get-MgBetaAdminReportSetting).DisplayConcealedNames -eq $true) {
   $Parameters = @{ displayConcealedNames = $False }
   Write-Host "Unhiding obscured report data for the script to run..."
   Update-MgBetaAdminReportSetting -BodyParameter $Parameters
   $ObscureFlag = $true
}

# Get user account information and load it into a hash table so that we can use it along with the OneDrive info
Write-Host "Finding user account information..."
[array]$Users = Get-MgUser -All -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
    -ConsistencyLevel Eventual -CountVariable UserCount -Sort 'displayName' `
    -Property Id, displayName, userPrincipalName, city, country, department, jobTitle, officeLocation

$UserHash = @{}
ForEach ($User in $Users) {
    $UserHash.Add($User.userPrincipalName, $User)
}    

# Get a list of OneDrive for Business sites in the tenant sorted by the biggest consumer of quota
Write-Host "Finding OneDrive sites..."
# This could also be done with Get-MgReportOneDriveUsageAccountDetail -Period D7 -Outfile $TempExportFile
$Uri = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D7')"
Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath $TempExportFile
[array]$ODFBSites = Import-CSV $TempExportFile | Sort-Object 'User display name'
If (!($ODFBSites)) { 
    Write-Host "No OneDrive sites found (surprisingly...)" ; 
    break 
}
# Calculate total storage used by OneDrive for Business accounts
$TotalODFBGBUsed = [Math]::Round(($ODFBSites.'Storage Used (Byte)' | Measure-Object -Sum).Sum /1GB,2)
# Create list to store report data
$Report = [System.Collections.Generic.List[Object]]::new()
# Store information for each OneDrive site
ForEach ($Site in $ODFBSites) {
    [array]$UserData = $UserHash[$Site.'Owner Principal name']
    $ReportLine   = [PSCustomObject]@{
        Owner       = $Site.'Owner display name'
        UPN	        = $Site.'Owner Principal name'
        City        = $UserData.city
        Country     = $UserData.Country
        Department  = $UserData.Department
        'Job Title' = $UserData.Jobtitle
        QuotaGB     = [Math]::Round($Site.'Storage Allocated (Byte)'/1GB,2) 
        UsedGB      = [Math]::Round($Site.'Storage Used (Byte)'/1GB,4)
        PercentUsed = [Math]::Round(($Site.'Storage Used (Byte)'/$Site.'Storage Allocated (Byte)' * 100),4) 
    }
    $Report.Add($ReportLine) 
}

$Report | Export-CSV -NoTypeInformation $CSVOutputFile
# You don't have to do this, but it's useful to view the data via Out-GridView
$Report | Sort-Object UsedGB -Descending | Out-GridView
Write-Host ("Current OneDrive for Business storage consumption is {0} GB. Report is in {1}" -f $TotalODFBGBUsed, $CSVOutputFile)

# Switch the tenant report obscure data setting back if necessary
If ($ObscureFlag -eq $true) {
    Write-Host "Resetting tenant data concealment for reports to True" -foregroundcolor red
    $Parameters = @{ displayConcealedNames = $True }
    Update-MgBetaAdminReportSetting -BodyParameter $Parameters
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
