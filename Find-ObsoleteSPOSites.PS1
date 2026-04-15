# Find-ObsoleteSPOSites.PS1
# A script to find SharePoint Online sites that have not had any content modifications for a specified number of days. This can be used to identify sites that are potentially obsolete and can be removed or archived.
# The script uses the Get-SPOSite cmdlet to retrieve information about all sites in the tenant, and then filters the sites based on the last content modification date. The script also retrieves the owners of the site if it is a Microsoft 365 Group-enabled site, and includes this information in the report.
# Requires Exchange Online and SharePoint Online PowerShell modules
# Session must be connected to an admin account

# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-ObsoleteSPOSites.PS1

Function Get-Microsoft365GroupOwners([String]$SiteURL) {
# Function to return the owners of an Microsoft 365 Group identified by the group GUID
$Owners = $Null; $DeletedGroup = $False; $SiteOwners = $Null
# Get the site properties. We need a separate call here because Get-SPOSite doesn't return all properties when it fetches a set of sites
$Group = (Get-SPOSite -Identity $SiteURL) 
If ($Group.Template -eq  "TEAMCHANNEL#0" -or $GroupId.Template -eq "TEAMCHANNEL#1") { # If Teams private or shared channel, we use the Related Group Id
   $GroupId = $Group | Select-Object -ExpandProperty RelatedGroupId 
} Else { # And for all other group-enabled sites, we use the GroupId
   $GroupId = $Group | Select-Object -ExpandProperty GroupId 
}

If ($GroupId.Guid -eq "00000000-0000-0000-0000-000000000000") { # Null group id stored in site
      $SiteOwners = "Deleted group"; $DeletedGroup = $True }
If ($DeletedGroup -eq $False) {      
   Try { 
        $GroupId = $GroupId.Guid
        [array]$Owners = Get-MgGroupOwner -GroupId $GroupId -ErrorAction Stop
   } Catch { 
        $SiteOwners = "Possibly deleted Microsoft 365 Group"; $DeletedGroup = $True 
   }
}

If ($Null -eq $Owners) { # Got nothing back, maybe because of an error
        $SiteOwners = "Possibly deleted Microsoft 365 Group"
   } Else { # We have some owners, now format them
        $SiteOwners = $Owners.additionalProperties.displayName -Join ", "
   }

Return $SiteOwners 
}

# Check that we are connected to Exchange Online and SharePoint Online
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "Microsoft.Online.Sharepoint.PowerShell")) {
   Write-Host "Please connect to the SharePoint Online Management module and then restart the script"; break
}
# Connect to the Microsoft Graph in app-only mode to read group owner information, SPO report information, and 
# Change these values for your tenant and app registration. The app registration needs to have the Group.Read.All and Reports.Read.All application permissions in order for the script to work.
$AppId = "34937156-d803-4a0a-b811-6a043ca67413"
$TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3478"
$CertificateThumbprint = "8CD4AD070C7447BA767EAB5DA659A02E6411BB80"
Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint 

# Get all SPO sites
Clear-Host
Write-Host "Fetching information about SharePoint Online sites..."
[array]$Sites = Get-SPOSite -Limit All | Select-Object Title, URL, StorageQuota, StorageUsageCurrent, ArchiveStatus, Template | `
   Sort-Object StorageUsageCurrent -Descending
If ($Sites.Count -eq 0) { 
   Write-Host "No SharePoint Online sites found.... exiting..." ; break 
}
$Sites = $Sites | Where-Object { $_.ArchiveStatus -eq "NotArchived" } # Remove archived sites
# Remove sites used for Teams private/shared channels. Comment this out if you want to include the channel sites in the report.
$Sites = $Sites | Where-Object {$_.Template -ne "TEAMCHANNEL#0" -and $_.Template -ne "TEAMCHANNEL#1"} 
# Remove redirect sites. These are not really sites, but just placeholders that point to the real site, and they don't have any activity or storage usage themselves, so we can remove them from the report. Comment this out if you want to include the redirect sites in the report.
$Sites = $Sites | Where-Object {$_.Template -ne "REDIRECTSITE#0"}

Write-Host "Fetching SharePoint Online site activity information from the Microsoft Graph..."
$TempDownLoadFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\download.csv"
$URI = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D180')"
Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath $TempDownloadFile
[array]$SitesActivityData = Import-CSV $TempDownloadFile

$SiteActivityHash = @{}
ForEach ($Site in $SitesActivityData) {
    $SiteId = $Site.'Site Id'
    Try {
        $SiteInfo = Get-MgSite -SiteId $SiteId -ErrorAction Stop
        If ($Site.'Last Activity Date') {
            $SiteActivityHash.Add($SiteInfo.WebURL, $Site)
        } Else {
            $SiteActivityHash.Add($SiteInfo.WebURL, "No activity")
        }
    } Catch {
        Write-Host "Failed to get site information for SiteId: $SiteId"
        Continue
    }
}

Clear-Host

$ProgressDelta = 100/($Sites.count); $PercentComplete = 0; $SiteNumber = 0
$TotalSPOStorageUsed = [Math]::Round(($Sites.StorageUsageCurrent | Measure-Object -Sum).Sum /1024,2)
$Report = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Site in $Sites) {   
   $SiteOwners = $Null ; $Process = $True; $SiteType = $null
   $SiteNumber++
   $SiteStatus = $Site.Title + " ["+ $SiteNumber +"/" + $Sites.Count + "]"
   Write-Progress -Activity "Processing site" -Status $SiteStatus -PercentComplete $PercentComplete
   $PercentComplete += $ProgressDelta
   $NoCheckGroup = $False
   Switch ($Site.Template.toUpper()) {  #Figure out the type of site and if we should process it - this might not be an exhaustive set of site templates
      "GROUP#0"                   {$SiteType = "Group-enabled team site"}
      "TEAMCHANNEL#0"             {$SiteType = "Teams Private or Shared Channel"}         
      "TEAMCHANNEL#1"             {$SiteType = "Teams Shared or Private Channel"}
      "STS#0"                     {$SiteType = "Team Site"; $NoCheckGroup = $True; $SiteOwners = "System"}
      "REVIEWCTR#0"               {$SiteType = "Review Center"; $Process = $False}
      "APPCATALOG#0"              {$SiteType = "App Catalog"; $Process = $False}
      "STS#3"                     {$SiteType = "Team Site"; $NoCheckGroup = $True; $SiteOwners = "System"}
      "SPSMSITEHOST#0"            {$SiteType = "Unknown"; $Process = $False}
      "SRCHCEN#0"                 {$SiteType = "Search Center"; $Process = $False}
      "EHS#1"                     {$SiteType = "Team Site - SPO Configuration"; $NoCheckGroup = $True; $SiteOwners = "System"}
      "EDISC#0"                   {$SiteType = "eDiscovery Center"; $Process = $False}
      "SITEPAGEPUBLISHING#0"      {$SiteType = "Site page"; $NoCheckGroup = $True; $SiteOwners = "System"}
      "POINTPUBLISHINGHUB#0"      {$SiteType = "Communications Site"; $NoCheckGroup = $True; $SiteOwners = "System" }
      "POINTPUBLISHINGPERSONAL#0" {$SiteType = "OneDrive for Business"; $Process = $False}
      "POINTPUBLISHINGTOPIC#0"    {$SiteType = "Office 365 Video"; $NoCheckGroup = $True; $SiteOwners = "System"} 
      "REDIRECTSITE#0"            {$SiteType = "Redirect site"; $Process = $False; $SiteOwners }
      Default                     {$SiteType = "Other"}
      }
   If ($NoCheckGroup -eq $False) { # Get owner information if it's a Microsoft 365 Group
      $SiteOwners = Get-Microsoft365GroupOwners($Site.URL)
   }

   $UsedGB = [Math]::Round($Site.StorageUsageCurrent/1024,2) 
   If ($Site.StorageUsageCurrent -gt 0) {
      $PercentTenant = ([Math]::Round($Site.StorageUsageCurrent/1024,4)/$TotalSPOStorageUsed).tostring("P") 
   } Else {
      $PercentTenant = "0%"
   }         

   # Calculate how long it's been since the last activity date from the usage report data, if we have one. This helps us 
   # understand if the site is still active or not, and can be used to filter the report if necessary
   $LastActivityInfo = $SiteActivityHash[$Site.URL]
   $LastActivityDate = $LastActivityInfo.'Last Activity Date'
   If ($LastActivityDate -and $LastActivityDate -ne "No activity") {
      $LastActivityDate = Get-Date $LastActivityDate -Format 'dd-MMM-yyyy'
      $DaysSinceActivity = (New-TimeSpan -Start $LastActivityDate -End (Get-Date)).Days
   } Else {
      $LastActivityDate = "No activity"
      $DaysSinceActivity = "N/A"
   }    

# And write out the information about the site
   If ($Process -eq $True) {
      $ReportLine = [PSCustomObject]@{
        URL               = $Site.URL
        SiteName          = $Site.Title
        Owner             = $SiteOwners
        Template          = $SiteType
        QuotaGB           = [Math]::Round($Site.StorageQuota/1024,0) 
        UsedGB            = $UsedGB
        Files             = $LastActivityInfo.'File Count'
        PercentUsed       = ([Math]::Round(($Site.StorageUsageCurrent/$Site.StorageQuota),4).ToString("P")) 
        PercentTenant     = $PercentTenant
        'Last Activity'   = $LastActivityDate
        DaysSinceActivity = $DaysSinceActivity
      }
      $Report.Add($ReportLine)
   }
}

$Report = $Report | Sort-Object DaysSinceActivity -Descending

# Generate the report in either Excel worksheet or CSV format, depending on if the ImportExcel module is available
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\SPO Site List.xlsx"
    $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "SPO Sites" -Title ("SPO Sites {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "SPOSites" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\SPO Site List.CSV"
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}   

Write-Host ("{0} sites processed. Current SharePoint Online storage consumption is {1} GB." -f $Report.count, $TotalSPOStorageUsed)

If ($ExcelGenerated -eq $true) {
    Write-Host ("SPO sites report is available in Excel workbook {0}" -f $ExcelOutputFile)
} Else {
    Write-Host ("SPO sites report is available in CSV file {0}" -f $CSVOutputFile)
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
