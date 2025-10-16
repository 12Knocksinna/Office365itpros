# Report-AppsAndServicePrincipals.PS1

[Array]$Apps = Get-MgApplication -All
If (!$Apps) {
    Write-Output "No applications found"
    break
} Else {
    Write-Output ("{0} applications found" -f $Apps.Count)  
}   

$AppReport = [System.Collections.Generic.List[Object]]::new()

ForEach ($App in $Apps) {
    $AppOwners = $null; $AppOwnersString = $null; $SPObjectId = $null; $SPOwners = $null; $ARAs = $null
    $AppOwners = (Get-MgApplicationOwner -ApplicationId $App.Id -All)
    If ($AppOwners) {
        $AppOwnersString = $AppOwners.additionalProperties.displayName -join "; "
    }
    $AppIdentifierUris = $App.IdentifierUris -join ";"
    $AppReplyUrls = $App.Web.RedirectUris -join ";"
    
    If ($App.ModifiedDateTime) {
        $AppModifiedDateTime = Get-Date $App.ModifiedDateTime -Format "mm-MMM-yyyy HH:mm:ss"
    } Else {
        $AppModifiedDateTime = $null
    }
    $AppRequiredResourceAccessString = $AppRequiredResourceAccess -join ";"
    
    # Get service principal for the app
    $SP = Get-MgServicePrincipal -Filter "appId eq '$AppId'" -All
    If (!$SP) {
        Write-Output ("No service principal found for application {0}" -f $App.DisplayName)
        Continue
    } Else {
        Write-Output ("Service principal found for application {1}" -f $App.DisplayName)  
        
            $SPOwners = (Get-MgServicePrincipalOwner -ServicePrincipalId $SP.Id -All | Select-Object -ExpandProperty UserPrincipalName) -join ";"
            $SPAppRoleAssignmentRequired = $SP.AppRoleAssignmentRequired
            # Get app role assignments for the service principal
            [Array]$ARAs = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $SP.Id -All
            If (!$ARAs) {
                Write-Output ("No app role assignments found for service principal {0}" -f $SP.DisplayName)
                Continue
            } Else {
                Write-Output ("{0} app role assignments found for service principal {1}" -f $ARAs.Count, $SP.DisplayName)  
            }
        
    }
    $AppReportLine = [PSCustomObject]@{
        AppName                      = $App.DisplayName
        AppId                        = $App.AppId
        AppObjectId                  = $App.Id
        AppOwners                    = $AppOwnersString
        AppCreatedDateTime           = Get-Date $App.CreatedDateTime -Format "dd-MMM-yyyy HH:mm:ss"
        AppModifiedDateTime          = $AppModifiedDateTime
        AppIdentifierUris            = $AppIdentifierUris
        AppReplyUrls                 = $AppReplyUrls
        
        ServicePrincipalObjectId     = $SPObjectId
        ServicePrincipalOwners       = $SPOwners
        ServicePrincipalCreatedDateTime = Get-Date $SP.additionalProperties['createdDateTime'] -Format "mm-MMM-yyyy HH:mm:ss"
        ServicePrincipalAppRoleAssignmentRequired = $SPAppRoleAssignmentRequired
        ServicePrincipalAppRoleAssignmentsCount = $ARAs.Count
    }   
    $AppReport.Add($AppReportLine)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.