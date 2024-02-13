# ReportPermissionsApps.PS1
# A script using Azure Automation and a managed identity to scan for apps assigned high-priority permissions and report them
# for administrator review

Connect-MgGraph -Identity -NoWelcome

# Get tenant information
$Tenant = (Get-MgOrganization)
$TenantId = $Tenant.Id
$TenantName = $Tenant.DisplayName
# Define the set of high-priority app roles (permissions) we're interested in
[array]$HighPriorityPermissions = "User.Read.All", "User.ReadWrite.All", "Mail.ReadWrite", `
  "Files.ReadWrite.All", "Calendars.ReadWrite", "Mail.Send", "User.Export.All", "Directory.Read.All", `
  "Exchange.ManageAsApp", "Directory.ReadWrite.All", "Sites.ReadWrite.All", "Application.ReadWrite.All", `
  "Group.ReadWrite.All", "ServicePrincipalEndPoint.ReadWrite.All", "GroupMember.ReadWrite.All", `
  "RoleManagement.ReadWrite.Directory", "AppRoleAssignment.ReadWrite.All"

# Define check period for new service principals
[datetime]$CheckforRecent = (Get-Date).AddDays(-360)
# Populate a set of hash tables with permissions used for different Office 365 management functions
$GraphApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
# Populate hash table with Graph permissions
$GraphRoles = @{}
ForEach ($Role in $GraphApp.AppRoles) { $GraphRoles.Add([string]$Role.Id, [string]$Role.Value) }
# Populate hash table with Exchange Online permissions
$ExoPermissions = @{}
$ExoApp = Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'"
ForEach ($Role in $ExoApp.AppRoles) { $ExoPermissions.Add([string]$Role.Id, [string]$Role.Value) }
$O365Permissions = @{}
$O365API = Get-MgServicePrincipal -Filter "DisplayName eq 'Office 365 Management APIs'"
ForEach ($Role in $O365API.AppRoles) { $O365Permissions.Add([string]$Role.Id, [string]$Role.Value) }
$AzureADPermissions = @{}
$AzureAD = Get-MgServicePrincipal -Filter "DisplayName eq 'Windows Azure Active Directory'"
ForEach ($Role in $AzureAD.AppRoles) { $AzureADPermissions.Add([string]$Role.Id, [string]$Role.Value) }
$TeamsPermissions = @{}
$TeamsApp = Get-MgServicePrincipal -Filter "DisplayName eq 'Skype and Teams Tenant Admin API'"
ForEach ($Role in $TeamsApp.AppRoles) { $TeamsPermissions.Add([string]$Role.Id, [string]$Role.Value) }
$RightsManagementPermissions = @{}
$RightsManagementApp = Get-MgServicePrincipal -Filter "DisplayName eq 'Microsoft Rights Management Services'"
ForEach ($Role in $RightsManagementApp.AppRoles) { $RightsManagementPermissions.Add([string]$Role.Id, [string]$Role.Value) }

# Create output tables
$Report = [System.Collections.Generic.List[Object]]::new() 
$ProblemApps = [System.Collections.Generic.List[Object]]::new()

# Find all service principals belonging to registered apps
[array]$SPs = Get-MgServicePrincipal -All -Filter "Tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')"
# And those for managed identities
[array]$ManagedIdentities = Get-MgServicePrincipal -All -Filter "ServicePrincipalType eq 'ManagedIdentity'"
[array]$Apps = $SPs + $ManagedIdentities

# Loop through the apps and managed identities we've found to resolve the permissions assigned to each app
Write-Output ("{0} service principals for Entra ID registered apps found" -f $Apps.Count)
$i = 0
ForEach ($App in $Apps) {
   $i++; $AppRecentFlag = $False
   # Write-Output ("Processing app {0} {1}/{2}" -f $App.DisplayName, $i, $Apps.Count)
   If ($App.AdditionalProperties.createdDateTime) {
        [datetime]$AppCreationDate = $App.AdditionalProperties.createdDateTime
    } Else {
     # For some reason, no app created date is available, so set it to a date in 1970
        [datetime]$AppCreationDate = '1970-01-01'
   }
   If ($AppCreationDate -gt $CheckforRecent) {$AppRecentFlag = $True}
   [array]$AppRoles = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $App.Id
   If ($AppRoles) {
     ForEach ($AppRole in $AppRoles) { 
       $Permission = $Null
       Switch ($AppRole.ResourceDisplayName) {
        "Microsoft Graph" { 
          [string]$Permission = $GraphRoles[$AppRole.AppRoleId] }
        "Office 365 Exchange Online" {
          [string]$Permission = $ExoPermissions[$AppRole.AppRoleId] }
        "Office 365 Management APIs" {
          [string]$Permission = $O365Permissions[$AppRole.AppRoleId] }
        "Windows Azure Active Directory" {
          [string]$Permission = $AzureADPermissions[$AppRole.AppRoleId] }
        "Skype and Teams Tenant Admin API" {
          [string]$Permission = $TeamsPermissions[$AppRole.AppRoleId] }
        "Microsoft Rights Management Services" {
          [string]$Permission = $RightsManagementPermissions[$AppRole.AppRoleId] }
       }
     
       If ($App.ServicePrincipalType -ne "ManagedIdentity") {
          If ($App.AppOwnerOrganizationId -eq $TenantId) { #Resolve tenant name
              $Name = $TenantName 
          } Else {
              $LookUpId = $App.AppOwnerOrganizationId.toString()
              $Uri = "https://graph.microsoft.com/beta/tenantRelationships/findTenantInformationByTenantId(tenantId='$LookUpId')"
              $ExternalTenantData = Invoke-MgGraphRequest -Uri $Uri -Method Get
              $Name = $ExternalTenantData.DisplayName 
          }
        $VerifiedPublisher = $Null
        If ($App.AdditionalProperties["verifiedpublisher"]) { 
            $VerifiedPublisher = $App.AdditionalProperties["verifiedpublisher"] 
        }
       } 
   
       $ReportLine  = [PSCustomObject] @{
           DisplayName        = $App.DisplayName
           ServicePrincipalId = $App.Id
           Publisher          = $App.PublisherName
           VerifiedPublisher  = $VerifiedPublisher
           Homepage           = $App.Homepage
           OwnerOrgId         = $App.AppOwnerOrganizationId
           OwnerOrgName       = $Name
           AppRoleId          = $AppRole.AppRoleId
           AppRoleCreation    = $AppRole.CreationTimeStamp
           Id                 = $AppRole.Id
           Resource           = $AppRole.ResourceDisplayName
           ResourceId         = $AppRole.ResourceId 
           Permission         = $Permission 
           SPType             = $App.ServicePrincipalType 
           CreatedDate        = Get-Date($AppCreationDate) -format g
           RecentApp          = $AppRecentFlag }
         $Report.Add($ReportLine) }
    } # End ForEach AppRole
} #End ForEach App

Write-Output ("{0} apps scanned and {1} permissions found" -f $Apps.Count, $Report.Count)
Write-Output ""
Write-Output "Permissions found"
Write-Output "-----------------"
Write-Output ""

$Report | Group-Object Permission | Sort-Object Name | Format-Table Name, Count
# Uncomment the next two lines if you want to generate the output file (interactive use only)
# $Report | Export-CSV -NoTypeInformation $OutputFile
# Write-Output ("Full permissions report is available in {0}" -f $OutputFile)

ForEach ($Record in $Report) {
  If ($Record.Permission -in $HighPriorityPermissions) {
     $ProblemApps.Add($Record) }
}
# Uncomment the next two lines if you want to generate the report about problem apps
# $ProblemApps | Export-CSV -NoTypeInformation $OutputFile2
# Write-Output ("List of apps for review is available in {0}" -f $OutputFile2)

[array]$AppSet = $ProblemApps | Sort-Object ServicePrincipalId -Unique

$AppOutput = [System.Collections.Generic.List[Object]]::new()
ForEach ($App in $AppSet) { 
   $Records = $ProblemApps | Where-Object {$_.ServicePrincipalId -eq $App.ServicePrincipalId}
   $AppPermissions = $Records.Permission -join ", "
   $ReportLine  = [PSCustomObject] @{
           DisplayName        = $App.DisplayName
           ServicePrincipalId = $App.ServicePrincipalId
           Publisher          = $App.Publisher
           Permissions        = $AppPermissions
           SPType             = $App.SPType 
           CreatedDate        = $App.CreatedDate
           RecentApp          = $App.RecentApp}
         $AppOutput.Add($ReportLine) 
}

# Add sign-in information for apps
[array]$MIAuditRecords = Get-MgAuditLogSignIn -Filter "(signInEventTypes/any(t:t eq 'managedIdentity'))" -Top 2000 -Sort "createdDateTime DESC"
[array]$AuditRecords = Get-MgAuditLogSignIn -Filter "(signInEventTypes/any(t:t eq 'servicePrincipal'))" -Top 2000 -Sort "createdDateTime DESC"
$AuditRecords = $AuditRecords + $MIAuditRecords

ForEach ($AppRecord in $AppOutput) {
  $SignInFound = $AuditRecords | Where-Object {$_.ServicePrincipalId -eq $AppRecord.ServicePrincipalId} | Select-Object -First 1
  If ($SignInFound) { Write-Output ("App {0} last signed in at {1}" -f $AppRecord.DisplayName, $SignInFound.CreatedDateTIme) 
  $AppRecord  | Add-Member -NotePropertyName LastSignIn -NotePropertyValue $SignInFound.CreatedDateTime
   }
}

#
If ($AppOutput) { # Generate a report and post it to Teams
  $AppOutput = $AppOutput | Sort-Object {$_.LastSignIn -as [datetime] } -Descending
  $Today = Get-Date -format dd-MMM-yyyy
  $Body = '
<style>
	.UserTable {
		border:1px solid #C0C0C0;
		border-collapse:collapse;
		padding:5px;
	}
	.UserTable th {
		border:1px solid #C0C0C0;
		padding:5px;
		background:#F0F0F0;
	}
	.UserTable td {
		border:1px solid #C0C0C0;
		padding:5px;
	}
</style>
<p><font size="2" face="Segoe UI">
<h3>Generated: ' + $Today + '</h3></font></p>
<table class="UserTable">
	<caption><h2><font face="Segoe UI">Azure Automation: Apps with High-Priority Permissions for Review</h2></font></caption>
	<thead>
	<tr>
	    <th>App Name</th>
	    <th>Service Principal</th>
          <th>Publisher</th>
          <th>Permissions</th>
          <th>App Type</th>
          <th>Created</th>
          <th>Last Signin</th>
	</tr>
	</thead>
	<tbody>'

ForEach ($A in $AppOutput) {
    $Body += "<tr><td><font face='Segoe UI'>$($A.DisplayName)</font></td><td><font face='Segoe UI'>$($A.ServicePrincipalId)</td></font><td><font face='Segoe UI'>$($A.Publisher)</td></font><td><font face='Segoe UI'>$($A.Permissions)</td></font><td><font face='Segoe UI'>$($A.SPType)</td><td><font face='Segoe UI'>$($A.CreatedDate)</td><td><font face='Segoe UI'>$($A.LastSignIn)</td></tr></font>"
    }
$Body += "</tbody></table><p>" 
$Body += '</body></html>'

Write-Output "Posting to Channel"

# Get username and password from Key Vault
$UserName = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -Name "ExoAccountName" -AsPlainText
$UserPassword = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -name "ExoAccountPassword" -AsPlainText
# Create credentials object from the username and password
[securestring]$SecurePassword = ConvertTo-SecureString $UserPassword -AsPlainText -Force
[pscredential]$UserCredentials = New-Object System.Management.Automation.PSCredential ($UserName, $SecurePassword)
# Get Site URL to use with PnP connection
$SiteURL = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -name "SPOSiteURL" -AsPlainText
# Target team and channel in that team to which we post a message containing the report
$TargetTeamId = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -name "TargetTeamId" -AsPlainText
$TargetTeamChannel = Get-AzKeyVaultSecret -VaultName "Office365ITPros" -name "TargetChannelID" -AsPlainText

# Connect to PnP using the account credentials we just retrieved
$PnpConnection = Connect-PnPOnline $SiteURL -Credentials $UserCredentials -ReturnConnection
# Post message to Teams
  Submit-PnPTeamsChannelMessage -Team $TargetTeamId -Channel $TargetTeamChannel -Message $Body -ContentType Html -Important -Connection $PnpConnection
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
