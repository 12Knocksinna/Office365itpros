# Report-DelegatePermissions.PS1
# A script to demonstrate how to report delegated permissions held by Entra ID user accounts
# V1.0 4-June-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-DelegatedPermissions.PS1

Connect-MgGraph -NoWelcome -Scopes Directory.Read.All
# Find licensed users
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
    -ConsistencyLevel eventual -CountVariable Records -All -Sort displayName
Write-Host ("{0} licensed user accounts found" -f $Users.count)

If (!($Users)) {
    Write-Host "No licensed users found - exiting!"; break
}
# Define file name for CSV output
$CSVOutputFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\DelegatedPermissions.csv"
# Scopes to ignore for reporting purposes
[array]$IgnoredScopes = "openid", "profile", "offline_access"
# Create hash tables for lookup of client and resource names
$ClientIds = @{}
$ResourceIds = @{}

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    Write-Host ("Checking delegated permissions for {0}" -f $User.UserPrincipalName)
    [array]$Permissions = Get-MgUserOauth2PermissionGrant -UserId $User.Id -All
    ForEach ($Permission in $Permissions) {
        # Try to look up client and resource names in the hash tables. If we find an entry, use the 
        # display name, else run Get-MgServicePrincipal to find the display name and store it
        $ClientDisplayName = $ClientIds[$Permission.ClientId]
        If ($null -eq $ClientDisplayName) {
            $ClientDisplayName = (Get-MgServicePrincipal -ServicePrincipalId $Permission.ClientId).displayName
            $ClientIds.Add($Permission.ClientId, $ClientDisplayName)
        }
        $ResourceDisplayName = $ResourceIds[$Permission.ResourceId]
        If ($null -eq $ResourceDisplayName) {
            $ResourceDisplayName = (Get-MgServicePrincipal -ServicePrincipalId $Permission.ResourceId).displayName
            $ResourceIds.Add($Permission.ResourceId, $ResourceDisplayName)
        }
        # Find the set of assigned scopes, ignoring some of the common scopes
        [array]$Scopes = $Permission.scope.Split(" ")
        [array]$FoundScopes = $null
        ForEach ($Scope in $Scopes) {
            If ($Scope -in $IgnoredScopes -or [string]::isNullOrWhiteSpace($Scope)) {
                Continue
            }
            $FoundScopes += $Scope
        }
        # Generate the output
        $ReportLine = [PSCustomObject][Ordered]@{
            'Consent type'      = $Permission.consentType
            UserPrincipalName   = $User.UserPrincipalName
            Client              = $ClientDisplayName
            Resource            = $ResourceDisplayName
            Scopes              = $FoundScopes -join ", "
            ClientId            = $Permission.ClientId
        }
        $Report.Add($ReportLine)
    }
}

# Now handle the AppPrincipals delegated permissions
Write-Host "Checking delegated permissions for all user accounts (AllPrincipals)"
[array]$AllPermissions = Get-MgOauth2PermissionGrant -filter "consentType eq 'AllPrincipals'" -All 
Write-Host ("{0} delegated permissions found for all user accounts" -f $AllPermissions.count)
[int]$i = 0
ForEach ($AllPermission in $AllPermissions) {
    $i++
    # Try to look up client and resource names in the hash tables. If we find an entry, use the 
    # display name, else run Get-MgServicePrincipal to find the display name and store it
    $ClientDisplayName = $ClientIds[$AllPermission.ClientId]
    If ($null -eq $ClientDisplayName) {
        $ClientDisplayName = (Get-MgServicePrincipal -ServicePrincipalId $AllPermission.ClientId).displayName
        $ClientIds.Add($AllPermission.ClientId, $ClientDisplayName)
    }
    Write-Host ("Checking permission for client {0} ({1}/{2})" -f $ClientDisplayName, $i, $AllPermissions.count)
    $ResourceDisplayName = $ResourceIds[$AllPermission.ResourceId]
    If ($null -eq $ResourceDisplayName) {
        $ResourceDisplayName = (Get-MgServicePrincipal -ServicePrincipalId $AllPermission.ResourceId).displayName
        $ResourceIds.Add($AllPermission.ResourceId, $ResourceDisplayName)
    }
    # Find the set of assigned scopes, ignoring some of the common scopes
    [array]$Scopes = $AllPermission.scope.Split(" ")
    [array]$FoundScopes = $null
    ForEach ($Scope in $Scopes) {
        If ($Scope -in $IgnoredScopes -or [string]::isNullOrWhiteSpace($Scope)) {
            Continue
        }
            $FoundScopes += $Scope
        }
    # Generate the output
    $ReportLine = [PSCustomObject][Ordered]@{
        'Consent type'      = $AllPermission.consentType
        UserPrincipalName   = "All User Accounts"
        Client              = $ClientDisplayName
        Resource            = $ResourceDisplayName
        Scopes              = $FoundScopes -join ", "
        ClientId            = $AllPermission.ClientId
    }
    $Report.Add($ReportLine)
}

$Report | Out-GridView -Title "Delegated Permissions Report"
$Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding UTF8
Write-Host ("CSV output file written to {0}" -f $CSVOutputFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.