# ReportSPOSiteStorageUsage.PS1
# Uses SharePoint Online and Exchange Online PowerShell modules
# Session must be connected to an admin account
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportSPOSiteStorageUsage.PS1

Function Get-Microsoft365GroupOwners([String]$SiteURL) {
# Function to return the owners of an Office 365 Group identified by the group GUID
$Owners = $Null; $DeletedGroup = $False; $i = 0; $SiteOwners = $Null
# Get the site properties. We need a separate call here because Get-SPOSite doesn't return all properties when it fetches a set of sites
$GroupId = (Get-SPOSite -Identity $SiteURL) 
If ($GroupId.Template -eq  "TEAMCHANNEL#0") { # If Teams private channel, we use the Related Group Id
   $GroupId = $GroupId | Select-Object -ExpandProperty RelatedGroupId 
} Else { # And for all other group-enabled sites, we use the GroupId
   $Groupid = $GroupId | Select-Object -ExpandProperty GroupId 
}

If ($GroupId.Guid -eq "00000000-0000-0000-0000-000000000000") { # Null group id stored in site
       $SiteOwners = "Deleted group"; $DeletedGroup = $True }
If ($DeletedGroup -eq $False) {      
   Try { 
      $Owners = (Get-UnifiedGroupLinks -Identity $GroupId.Guid -LinkType Owners -ErrorAction SilentlyContinue) 
   } Catch { 
      $SiteOwners = "Possibly deleted Microsoft 365 Group"; $DeletedGroup = $True 
   }
}

If ($Null -eq $Owners) { # Got nothing back, maybe because of an error
      $SiteOwners = "Possibly deleted Microsoft 365 Group"
   } Else { # We have some owners, now format them
      $Owners = $Owners | Select-Object -ExpandProperty DisplayName
      ForEach ($Owner in $Owners)  {
         If ($i -eq 0) { 
            $SiteOwners = $Owner; $i = 1 
         } Else { 
            $SiteOwners = $SiteOwners + "; " + $Owner
         }
      }
   }

Return $SiteOwners 
}

# Check that we are connected to Exchange Online and SharePoint Online
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {
   Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break
}
If (!($ModulesLoaded -match "Microsoft.Online.Sharepoint.PowerShell")) {
   Write-Host "Please connect to the SharePoint Online Management module and then restart the script"; break
}

# Get all SPO sites
Clear-Host
Write-Host "Fetching site information..."
[array]$Sites = Get-SPOSite -Limit All | Select-Object Title, URL, StorageQuota, StorageUsageCurrent, Template | `
   Sort-Object StorageUsageCurrent -Descending
If ($Sites.Count -eq 0) { Write-Host "No SharePoint Online sites found.... exiting..." ; break }
$TotalSPOStorageUsed = [Math]::Round(($Sites.StorageUsageCurrent | Measure-Object -Sum).Sum /1024,2)

Clear-Host
$ProgressDelta = 100/($Sites.count); $PercentComplete = 0; $SiteNumber = 0
$Report = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Site in $Sites) {
   $SiteOwners = $Null ; $Process = $True
   $SiteNumber++
   $SiteStatus = $Site.Title + " ["+ $SiteNumber +"/" + $Sites.Count + "]"
   Write-Progress -Activity "Processing site" -Status $SiteStatus -PercentComplete $PercentComplete
   $PercentComplete += $ProgressDelta
   $NoCheckGroup = $False
   Switch ($Site.Template) {  #Figure out the type of site and if we should process it - this might not be an exhaustive set of site templates
      "RedirectSite#0"            {$SiteType = "Redirect"; $Process = $False }
      "GROUP#0"                   {$SiteType = "Group-enabled team site"}
      "TEAMCHANNEL#0"             {$SiteType = "Teams Private Channel" }
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
      "POINTPUBLISHINGTOPIC#0"    {$SiteType = "Office 365 Video"; $NoCheckGroup = $True; $SiteOwners = "System"} }

   If ($NoCheckGroup -eq $False) { # Get owner information if it's a Microsoft 365 Group
      $SiteOwners = Get-Microsoft365GroupOwners($Site.URL)
   }

   $UsedGB = [Math]::Round($Site.StorageUsageCurrent/1024,2) 
   $PercentTenant = ([Math]::Round($Site.StorageUsageCurrent/1024,4)/$TotalSPOStorageUsed).tostring("P")           

# And write out the information about the site
   If ($Process -eq $True) {
      $ReportLine = [PSCustomObject]@{
         URL           = $Site.URL
         SiteName      = $Site.Title
         Owner         = $SiteOwners
         Template      = $SiteType
         QuotaGB       = [Math]::Round($Site.StorageQuota/1024,0) 
         UsedGB        = $UsedGB
         PercentUsed   = ([Math]::Round(($Site.StorageUsageCurrent/$Site.StorageQuota),4).ToString("P")) 
         PercentTenant = $PercentTenant}
      $Report.Add($ReportLine)
   }
}

# Now generate the report
$CSVOutputFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\SPOSiteConsumption.csv"
$Report | Export-CSV -NoTypeInformation $CSVOutputFile -Encoding utf8
Write-Host ("Current SharePoint Online storage consumption is {0} GB. Report is in {1}" -f $TotalSPOStorageUsed, $CSVOutputFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
