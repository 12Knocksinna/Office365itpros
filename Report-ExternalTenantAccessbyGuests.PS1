# Report-ExternalTenantAccessbyGuests.PS1
# This script generates a report of member accounts who have accessed apps in external tenants in a Microsoft 365 environment.

# V1.0 4-Dec-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ExternalTenantAccessbyGuests.PS1

# Scopes: AuditLog.Read.All (application or delegated), Directory.Read.All
Connect-MgGraph -Scopes "AuditLog.Read.All","Directory.Read.All"

$StartDate = (Get-Date).AddDays(-30).ToString('yyyy-MM-ddT00:00:00Z')
$EndDate  = (Get-Date).ToString('yyyy-MM-ddT23:59:59Z')

# Find this tenant information
$ThisTenantData = Get-MgOrganization
$ThisTenantId = $ThisTenantData.Id

# Get all successful sign-in audit records within the specified date range
[array]$AuditRecords = Get-MgBetaAuditLogSignIn -All `
  -Filter "createdDateTime ge $StartDate and createdDateTime le $EndDate and status/errorCode eq 0"
# Filter to find the set of external tenant accesses by guest users
[array]$ExternalAccessRecords = $AuditRecords | Where-Object { $_.HomeTenantId -ne $_.ResourceTenantId -and $_.HomeTenantId -eq $ThisTenantId}
[array]$InboundAccessRecords = $AuditRecords | Where-Object { $_.HomeTenantId -ne $_.ResourceTenantId -and $_.HomeTenantId -ne $ThisTenantId}

# Summarize: which host tenants your users accessed
# Find the unique tenant ids
[array]$HostTenants = $ExternalAccessRecords | Select-Object -ExpandProperty ResourceTenantId -Unique

Write-Output "Extracting information for host tenants..."
$TenantData = [System.Collections.Generic.List[Object]]::new()
ForEach ($Tenant in $HostTenants) {
    $TenantDetails = Find-MgTenantRelationshipTenantInformationByTenantId -TenantId $Tenant
    $TenantRecords = $ExternalAccessRecords | Where-Object { $_.ResourceTenantId -eq $Tenant }
    $TenantUsers = $TenantRecords | Select-Object -ExpandProperty UserPrincipalName -Unique
    $TenantApps = $TenantRecords | Select-Object -ExpandProperty AppDisplayName -Unique
    $TenantDataLine = [PSCustomObject]@{
        TenantId                = $Tenant
        TenantName              = $TenantDetails.DisplayName
        'Number of Accesses'    = $TenantRecords.Count
        'Unique Users'          = $TenantUsers.Count   
        'Users'                 = ($TenantUsers -join '; ')
        'Apps Accessed'         = ($TenantApps -join '; ')
        Type                    = 'Outbound'  
    }
    $TenantData.Add($TenantDataLine)
}

Write-Output "Extracting information for inbound tenants..."
$TenantUsers = $InboundAccessRecords | Select-Object -ExpandProperty UserPrincipalName -Unique
$TenantApps = $InboundAccessRecords | Select-Object -ExpandProperty AppDisplayName -Unique
$TenantDataLine = [PSCustomObject]@{
    TenantId                = $ThisTenantData.Id
    TenantName              = $ThisTenantData.DisplayName
    'Number of Accesses'    = $InboundAccessRecords.Count
    'Unique Users'          = $TenantUsers.Count   
    'Users'                 = ($TenantUsers -join '; ')
    'Apps Accessed'         = ($TenantApps -join '; ')
    Type                    = 'Inbound'  
}
$TenantData.Add($TenantDataLine)



# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.