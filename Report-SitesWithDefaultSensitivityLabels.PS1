# Report-SitesWithDefaultSensitivityLabels.PS1
# An example of using PnP PowerShell to report on sites with default sensitivity labels

# V1.0 18-Mar-2025
# GitHub link: 

# Define these values for your tenant
$PnPClientApp = 'cb5f363f-fbc0-46cb-bcfd-0933584a8c57'
$Thumbprint = '59176CF0C682A3D43E4CC4B38129F5CA91156C94'
$SPOSite = 'https://redmondassociates.sharepoint.com'
$TenantName = 'redmondassociates.onmicrosoft.com'

# Get the sensitivity labels in the tenant
[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Loading Exchange Online Management module"
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}
# And load the compliance module too so that we can run Get-Label
Connect-IPPSSession -ShowBanner:$false

Write-Host "Getting detauls of sensitivity labels used with files..."
# Create a hash table to map label IDs to display names
$Labels = [array]$Labels = Get-Label | Where-Object {$_.ContentType -Like "*File*"}
$LabelsHash = @{}
ForEach ($L in $Labels) { $LabelsHash.add([string]$L.ImmutableId,[string]$L.DisplayName) }

# Define the libraries to exclude from the report
[array]$ExcludedLibraries = "Teams Wiki Data", "Form Templates", "Site Assets", "Site Pages", "Style Library", "Model"

# Connect to the tenant
Connect-PnPOnline -ClientId $PnPClientApp -Interactive -Url $SPOSite
# Get all group sites that aren't archived
[array]$Sites = Get-PnPTenantSite -Template "GROUP#0" -Detailed -Filter "ArchiveStatus -eq 'NotArchived'"
Write-Host ("{0} group sites found" -f $Sites.Count)
# Check each site for a default sensitivity label
$Report = [System.Collections.Generic.List[Object]]::new()
Write-Host "Checking sites for document libraries with default sensitivity labels..."
ForEach ($Site in $Sites) {
    Connect-PnPOnline -ClientId $PnPClientApp -Url $Site.Url -Thumbprint $Thumbprint -Tenant $TenantName
    [array]$Lists = Get-PnPList 
    $Lists = $Lists | Where-Object {$_.BaseType -eq 'DocumentLibrary' -and $_.Hidden -eq $false}
    ForEach ($List in $Lists) {
        If ($List.Title -in $ExcludedLibraries ) { Continue }
        If (!([string]::IsNullOrWhiteSpace($List.DefaultSensitivityLabelForLibrary))) {
            [string]$LabelGuid = $List.DefaultSensitivityLabelForLibrary
            $ReportLine = [PSCustomObject][Ordered]@{
                SiteTitle               = $Site.Title
                SiteUrl                 = $Site.Url
                ListTitle               = $List.Title
                Label                   = $LabelsHash[$LabelGuid]
                LabelId                 = $List.DefaultSensitivityLabelForLibrary
            }
            $Report.Add($ReportLine)
            Write-Host ("The document library {0} in the {1} site has default sensitivity label {2} ({3})" -f $List.Title, $Site.Url, `
                $LabelsHash[$List.DefaultSensitivityLabelForLibrary], $List.DefaultSensitivityLabelForLibrary)
        }
    }
}

Disconnect-PnPOnline    
Write-Host ""
Write-Host "Details of sites with default sensitivity labels"
$Report | Format-Table SiteTitle, ListTitle, Label -AutoSize
Write-Host ""
Write-Host "Summary of sites with default sensitivity labels"
$Report | Group-Object Label -NoElement | Sort-Object Count -Descending | Format-Table Name, Count

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.