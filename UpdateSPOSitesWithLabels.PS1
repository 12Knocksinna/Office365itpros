# UpdateSPOSitesWithLabels.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateSPOSitesWithLabels.PS1
# A script to check SharePoint Online team sites created for Microsoft 365 Groups that don't seem to have received sensitivity labels.
# The script reads the sensitivity label information for groups from Exchange Online and uses the information to update the 
# SharePoint Online sites

# We need both the SharePoint Online and Exchange Online management modules to be loaded
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
If (!($ModulesLoaded -match "Microsoft.Online.Sharepoint.PowerShell")) {Write-Host "Please connect to the SharePoint Online Management module and then restart the script"; break}
CLS

[array]$Sites = Get-SPOSite -Limit All -Template Group#0
If (!($Sites)) { Write-Error "No sites for Microsoft 365 Groups found... exiting!" ; break}
   Else { Write-Host ("Processing {0} sites" -f $Sites.Count) }

$SitesNoLabels = [System.Collections.Generic.List[Object]]::new()
ForEach ($Site in $Sites) { #Check each site to see if it has a sensitivity label
        $SiteData = Get-SPOSite -Identity $Site.Url
        If ([string]::IsNullOrWhiteSpace(($SiteData.SensitivityLabel)) -eq $True) {
           Write-Host ("Site {0} has no label" -f $SiteData.Url) 
           $SiteInfo = [PSCustomObject][Ordered]@{  
              URL    = $SiteData.Url
              Title   = $SiteData.Title   }
           $SitesNoLabels.Add($SiteInfo) }
} #End ForEach Sites
Write-Host ("Finished checking. {0} of {1} sites for Microsoft 365 Groups have no sensitivity label" -f $SitesNoLabels.Count, $Sites.Count) 

Write-Host "Retrieving sensitivity label information for Microsoft 365 Groups"
[array]$Groups = Get-UnifiedGroup -ResultSize Unlimited 
$Groups = $Groups | ? {$_.SharePointSiteUrl -ne $Null}
$GroupsTable = @{}
$Groups.ForEach( {
       $GroupsTable.Add([String]$_.SharePointSiteUrl, $_.SensitivityLabel) } )

Write-Host "Starting to update SharePoint sites with sensitivity labels..."

[int]$Updates = 0; [int]$NoUpdates = 0
ForEach ($Site in $SitesNoLabels) {
    $Label = $Null
    $Label = $GroupsTable.Item($Site.Url)
    If ($Label) { # Update the site with the label we find
       Write-Host ("Updating site {0} with label {1}" -f $Site.Url, $Label.Guid) 
       Set-SPOSite -Identity $Site.Url -SensitivityLabel $Label.Guid 
       $Updates++ }
    Else {
       Write-Host ("Can't find sensitivity label for site {0} - group might be deleted" -f $Site.Url)
       $NoUpdates++ }
} #End ForEach Sites

Write-Host ("Finished Updating sites. {0} labels applied to sites. {1} sites could not be updated with labels" -f $Updates, $NoUpdates)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
