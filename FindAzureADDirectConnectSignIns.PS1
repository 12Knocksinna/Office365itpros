# FindAzureADDirectConnectSignIns.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindAzureADDirectConnectSignIns.PS1
# Script to show how to use the Azure AD sign-in audit log to find entries for user accounts accessing external tenants for Teams 
# Direct Connect (aka shared channels)

Connect-MgGraph -Scopes "AuditLog.Read.All, "Directory.Read.All"
Select-MgProfile -Name "beta"

$Tenant = Get-MgOrganization
$TenantId = $Tenant.Id
$TenantName = $Tenant.DisplayName
Write-Host "Finding Azure AD sign-ins for Azure B2B not from" $TenantName "..."
[array]$AzureADSignIns = Get-MgAuditLogSignIn -Filter "ResourceTenantId ne '$TenantID' and CrossTenantAccessType eq 'b2bDirectConnect'" -All

If (!($AzureADSignIns)) {
   Write-Host "No Azure AD sign-in records for B2B Direct Connct found from other Microsoft 365 tenants - exiting" ; break }
Else {
   Write-Host ("{0} Azure AD sign-in records from other Microsoft 365 tenants found - analyzing..." -f $AzureADSignIns.count ) }

$TenantNames = @{}
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Record in $AzureADSignIns){
   $ExternalTenantId = $Record.ResourceTenantId
   If (!($TenantNames[$ExternalTenantId])) {
   # Get the tenant name because we haven't stored it yet in the hash table
     $Uri = "https://graph.microsoft.com/beta/tenantRelationships/findTenantInformationByTenantId(tenantId='$ExternalTenantId')"
     $ExternalTenantData = Invoke-MgGraphRequest -Uri $Uri -Method Get
     $TenantNames.Add($ExternalTenantId,$ExternalTenantData.DisplayName)
     $ExternalTenantDisplayName = $ExternalTenantData.DisplayName 
   }
   Else { # We have seen the tenant name before, so just read the info.
     $ExternalTenantDisplayName = $TenantNames[$ExternalTenantId]
   }

   $ExternalData = [PSCustomObject][Ordered]@{  
       Timestamp    = $Record.CreatedDateTime
       User         = $Record.UserDisplayName
       UserId       = $Record.UserId
       UPN          = $Record.UserPrincipalName
       TenantName   = $ExternalTenantDisplayName
       TenantId     = $ExternalTenantId
       Resource     = $Record.ResourceDisplayName
       AppName      = $Record.AppDisplayName
       Type         = $Record.CrossTenantAccessType
       }     
   $Report.Add($ExternalData)
}

$Report | Sort {$_.Timestamp -as [datetime]} | Select Timestamp, User, TenantName, Resource, AppName | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
