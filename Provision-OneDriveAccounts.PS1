# Provision-OneDriveAccounts.PS1
# A script to pre-provision OneDrive for Business accounts
# A fixed version of what's in https://learn.microsoft.com/en-us/sharepoint/pre-provision-accounts#pre-provision-onedrive-for-all-licensed-users-in-your-organization
# V1.0 5-June-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Provision-OneDriveAccounts.PS1

# Connect to the Microsoft Graph PowerShell SDK
Try {
    Connect-MgGraph -Scopes Directory.Read.All -NoWelcome
} Catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit
}

# This part is to generate a hash table of OneDrive sites that we can use to look up
Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -SkipEditionCheck
# Make sure that you use the right URL for your tenant here
Connect-SPOService -Url https://office365itpros-admin.sharepoint.com
Write-Host "Looking for existing OneDrive sites..."
[array]$OneDriveSites = Get-SPOSite -IncludePersonalSite $true -Limit All `
    -Filter "Url -like '-my.sharepoint.com/personal/'" | Select-Object Owner, URL
# Build the hash table
$OneDriveHashTable = @{}
ForEach ($Site in $OneDriveSites) { 
    [string]$SiteKey = $Site.URL + "/"
    Try {
        $OneDriveHashTable.Add($Site.Owner, $SiteKey) 
    } Catch {
        # Write-Host ("Error {0} when adding {1} to hash table" -f $_.Exception.Message, $Site.Owner)   
    }
}

Write-Host "Looking for users with OneDrive licenses..."
# OneDrive doesn't have a service plan, but we can use SharePoint Online (5dbe027f-2339-4123-9542-606e4d348a72)
[guid]$SPOPlanId = "5dbe027f-2339-4123-9542-606e4d348a72" 
[array]$Users = Get-MgUser -filter "assignedPlans/any(s:s/serviceplanid eq $SPOPlanId and capabilityStatus eq 'Enabled')" `
    -ConsistencyLevel eventual -CountVariable Test -All -PageSize 999 `
    -Property Id, displayName, userprincipalName, assignedLicenses, assignedPlans, department, country `
    -Sort displayName
Write-Host ("Found {0} users with SharePoint (OneDrive) licenses" -f $Users.Count)

    # Check each user to see if a OneDrive site exists. If not, force its provisioning

[int]$CountOfProvisionedSites = 0
ForEach ($User in $Users) {
    If ($OneDriveHashTable.ContainsKey($User.UserPrincipalName)) {
        Write-Host ("User {0} already has a OneDrive for Business site" -f $User.UserPrincipalName)
    }  Else {
        Write-Host ("User {0} does not have a OneDrive site" -f $User.UserPrincipalName)
        # Provision the OneDrive site
        Request-SPOPersonalSite -UserEmails $User.userprincipalName -NoWait
        $CountOfProvisionedSites++
        Write-Host ("Requested provisioning of OneDrive for Business site for {0}" -f $User.UserPrincipalName) -ForegroundColor Yellow
    }
}
Write-Host ("Provisioned {0} OneDrive sites" -f $CountOfProvisionedSites) -ForegroundColor Green

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.