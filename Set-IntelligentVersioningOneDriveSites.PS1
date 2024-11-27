# Set-IntelligentVersioningOneDriveSites.PS1
# Set intelligent (automatic) versioning for OneDrive for Business sites for user accounts with SharePoint Online Plan 1 or Plan 2 licenses
# 
# V1.0 27-Nov-2024
# GitHub link:

Connect-MgGraph -NoWelcome -Scopes Directory.Read.All

# Connect to SharePoint Online
[array]$Domains = (Get-MgOrganization).verifiedDomains
$DefaultDomain = $Domains | Where-Object {$_.IsDefault -eq $true}
$SPOAdminRoot = ("https://{0}-admin.sharepoint.com" -f $DefaultDomain.Name.split('.')[0])
Write-Host "Connecting to SharePoint Online..."
Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell
Connect-SPOService -Url $SPOAdminRoot
If (Get-SPOTenant) {
    Write-Host ("Connected to SharePoint Online at {0}" -f $SPOAdminRoot)
} Else {
    Write-Host "Failed to connect to SharePoint Online"
    Break
}
$SPOPlan2 = '5dbe027f-2339-4123-9542-606e4d348a72'
$SPOPlan1 = 'c7699d2e-19aa-44de-8edf-1736da088ca1'
[array]$Users = Get-MgUser -Filter "(assignedPlans/any(s:s/serviceplanid eq $SPOPlan2)) or (assignedPlans/any(s:s/serviceplanid eq $SPOPlan1)) and userType eq 'Member'" `
    -ConsistencyLevel eventual -CountVariable UsersFound -PageSize 500 -Sort 'displayName'

If (!$Users) {
    Write-Host "No users found with SharePoint Online Plan 1 or Plan 2 licenses"
    Break
}
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($User in $Users.UserPrincipalName) {
    
    $OneDriveSiteURI = ("https://{0}-my.sharepoint.com/personal/{1}" -f ($DefaultDomain.Name.split('.')[0]), ($User -replace '\.|\@', '_'))

    Write-Host ("Checking OneDrive site {0} for user {1}" -f $OneDriveSiteURI, $User)
    $Site = Get-SPOSite -Identity $OneDriveSiteURI -ErrorAction SilentlyContinue
    If ($Site) {
        If ($Site. EnableAutoExpirationVersionTrim -ne $true) {
            Write-Host ("Updating {0} to enable intelligent versioning…" -f $Site.Url) -ForegroundColor Yellow
            Set-SPOSite -Identity $Site.Url -EnableAutoExpirationVersionTrim $true -Confirm:$false
            $Status = "Updated with intelligent versioning"
        } Else {
            Write-Host ("{0} already has intelligent versioning enabled" -f $Site.Url) -ForegroundColor Cyan
            $Status = "Intelligent versioning already enabled"
        }
        $ReportLine = [pscustomobject]@{
            User    = $User
            Site    = $Site.Url
            Status  = $Status
        }
        $Report.Add($ReportLine)
    } Else {
        Write-Host ("OneDrive site does not exist for {0}" -f $User)
    } 
}

Write-Host ("OneDrive accounts enabled with intelligent versioning: {0}" -f ($Report | Where-Object {$_.Status -eq "Updated with intelligent versioning"}).Count)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.