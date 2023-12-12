# Report-ComplianceCaseHolds.PS1
# Generate a report about holds for eDiscovery cases
# V1.0 12-Dec-2023
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ComplianceCaseHolds.PS1

$Status = Get-ConnectionInformation
If (!($Status)) {
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

# Connect to the compliance endpoint
Connect-IPPSession

$CSVFile = "C:\temp\ComplianceHoldReport.CSV"

# Find eDiscovery (standard) cases
Write-Host "Looking for eDiscovery cases"
[array]$Cases = Get-ComplianceCase -CaseType eDiscovery
[array]$PremiumCases = Get-ComplianceCase -CaseType AdvancedEdiscovery
[int]$NumberofCases = $Cases.Count + $PremiumCases.Count
If ($NumberofCases -eq 0) {
    Write-Host "No eDiscovery cases found - exiting" ; break
}
Write-Host ("Found {0} eDiscovery (standard) cases and {1} eDiscovery Premium cases - checking for holds" -f $Cases.Count, $PremiumCases.Count)
$Report = [System.Collections.Generic.List[Object]]::new()
[int]$i = 0
Write-Host "Processing standard cases"
ForEach ($Case in $Cases) {
    $i++
    Write-Host ("Processing case {0} ({1}/{2})" -f $Case.Name, $i, $Cases.Count)
    [array]$CaseHolds = Get-CaseHoldPolicy -Case $Case.Identity -DistributionDetail
    If ($CaseHolds) {
        ForEach ($CaseHold in $CaseHolds) {
            $CaseData = [PSCustomObject][Ordered]@{
                Name            = $CaseHold.Name
                Workload        = $CaseHold.Workload
                Enabled         = $CaseHold.Enabled
                Mode            = $CaseHold.Mode
                Exchange        = $CaseHold.ExchangeLocation
                SharePoint      = $CaseHold.SharePointLocation
                PublicFolders   = $CaseHold.PublicFolderLocation
                LastUpdate      = $Rule.LastStatusUpdateTime
                CaseType        = 'Standard'
            }
            $Report.Add($CaseData)
      }
   }
}

[int]$i = 0
Write-Host "Processing premium cases"
ForEach ($Case in $PremiumCases) {
    $i++
    Write-Host ("Processing case {0} ({1}/{2})" -f $Case.Name, $i, $PremiumCases.Count)
    [array]$CaseHolds = Get-CaseHoldPolicy -Case $Case.Identity -DistributionDetail
    If ($CaseHolds) {
        ForEach ($CaseHold in $CaseHolds) {
            $CaseData = [PSCustomObject][Ordered]@{
                Name            = $CaseHold.Name
                Workload        = $CaseHold.Workload
                Enabled         = $CaseHold.Enabled
                Mode            = $CaseHold.Mode
                Exchange        = $CaseHold.ExchangeLocation
                SharePoint      = $CaseHold.SharePointLocation
                PublicFolders   = $CaseHold.PublicFolderLocation
                LastUpdate      = $CaseHold.LastStatusUpdateTime
                CaseType        = 'Premium'
            }
        $Report.Add($CaseData)
      }
   }
}

$Report | Select-Object Name, CaseType, Enabled, Mode, Exchange, SharePoint, LastUpdate | Out-GridView
$Report | Export-CSV -NoTypeInformation $CSVFile
Write-Output ("The outfile is available in {0}" -f $CSVFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
