

$CSVOutputFile = "c:\temp\OneDriveSiteConsumption.CSV"
# Get a list of OneDrive for Business sites in the tenant sorted by the biggest consumer of quota
Write-Host "Finding OneDrive sites..."
$uri = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D180')"
Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath x.csv
[array]$ODFBSites = import-csv x.csv
If (!($ODFBSites)) { Write-Host "No OneDrive sites found (surprisingly...)" ; break }
# Calculate total used
$TotalODFBGBUsed = [Math]::Round(($ODFBSites.'Storage Used (Byte)' | Measure-Object -Sum).Sum /1GB,2)
# Create list to store report data
$Report = [System.Collections.Generic.List[Object]]::new()
# Store information for each OneDrive site
ForEach ($Site in $ODFBSites) {
      $ReportLine   = [PSCustomObject]@{
        Owner       = $Site.'Owner display name'
        UPN	    = $Site.'Owner Principal name'
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