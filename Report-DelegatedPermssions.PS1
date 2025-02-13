# Report-DelegatedPermissions.PS1
# A script to report delegated permission grants for Entra ID service principals and user accounts

Connect-MgGraph -Scopes Directory.Read.All -NoWelcome

# Details of delegated permissions https://learn.microsoft.com/en-us/graph/api/resources/oauth2permissiongrant?view=graph-rest-1.0
Write-Host "Finding OAuth2 permissions granted to service principals and users..."
[array]$DelegatedPermissions = Get-MgOauth2PermissionGrant -All -PageSize 500

If (!$DelegatedPermissions) {
    Write-Host "No delegated permissions found"
    Exit
}   

# Create a hash table to store common resource identifiers
$ResourceHash = @{}
$ResourceHash.Add("14a3c489-ed6c-4005-96d1-be9c5770f7a3", "Microsoft Graph")
$ResourceHash.Add("2539d435-a696-47ed-a41f-093ecb3f8fbd", "PowerApps Service")
$ResourceHash.Add("25cbf210-02e5-4a82-9f5c-f41befd2681a", "Microsoft Rights Management Service")
$ResourceHash.Add("2be71509-6ab9-44d7-bfd8-eff4e50bfc7c", "Windows Azure Active Directory")
$ResourceHash.Add("6a527dab-43d4-4c6e-b9eb-9af3e8d59a90", "Windows Azure Service Management API")
$ResourceHash.Add("48e8cc32-e0cf-44f3-938e-128fa7a88786", "Dataverse")
$ResourceHash.Add("6bfb2e2c-6932-4d65-b136-a9e7f792ad0b", "Microsoft Information Protection Sync Service")
$ResourceHash.Add("a0073afc-e9aa-4c6e-bb58-c6936694718b", "Office 365 SharePoint Online")
$ResourceHash.Add("dacf6086-a190-467a-aadd-d519472b8d1d", "Office 365 Exchange Online")
$ResourceHash.Add("b7f5b95e-3f26-4f15-8637-db5b3f213326", "Microsoft Forms")
$ResourceHash.Add("e63b8447-bf48-4e3f-93bb-d13258c8abda", "Power BI Service")
$ResourceHash.Add("c9b99704-9637-4dfa-8854-67c1505454d2", "Graph Connector Service")
$ResourceHash.Add("ebf6e6bd-3a61-469d-9acf-514b74f182c4", "Office 365 Management APIs")

$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Permission in $DelegatedPermissions) {
    
    Switch ($Permission.ConsentType) {
        "AllPrincipals" {
            $ConsentType = "All users"
            $Principal = "Service principal"
        }
        "Principal" {
            $ConsentType = "Specific user"
            $Principal = (Get-MgUser -UserId $Permission.PrincipalId).UserPrincipalName
        }
    }

    # Check the resource hash table to see if we can find a friendly name for the resource
    $ResourceId = $Permission.ResourceId
    $Resource = $ResourceHash[$ResourceId]
    If (!$Resource) {
        # If not, search the service principals to see if we can find a match
        $Resource = $ServicePrincipals | Where-Object {$_.Id -eq $ResourceId}
    }
    
    # Find the details of the service principal that holds the permission
    $Client = Get-MgServicePrincipal -ServicePrincipalId $Permission.ClientId

    $ReportLine = [PSCustomObject][Ordered]@{
        ClientId            = $Client.DisplayName
        ClientType          = $Client.ServicePrincipalType
        ResourceId          = $ResourceId
        Resource            = $Resource
        PrincipalId         = $Permission.PrincipalId
        Principal           = $Principal
        ConsentType         = $ConsentType
        Scope               = $Permission.Scope.trim()
    }
    $Report.Add($ReportLine)
}

# Another way of reporting is to find service principals and list their application and delegated roles (otherwise called permissions). These permission are
# set by the application developers
[array]$ServicePrincipals = Get-MgServicePrincipal -All -PageSize 500 -Property * | `
    Sort-Object DisplayName | 
    Select-Object DisplayName, Id, AppId, `
    @{name='Application Roles'; expression={($_.AppRoles | ForEach-Object Value) -join ", "}},
    @{name='Delegated Roles'; expression={($_.Oauth2PermissionScopes | ForEach-Object Value) -join ", "}}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.