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
    $AppOwners = $null
    $AppOwners = (Get-MgApplicationOwner -ApplicationId $App.Id -All)
    If ($AppOwners) {
        $AppOwnersString = $AppOwners.additionalProperties.displayName -join "; "
    }
    $AppCreatedDateTime = $App.CreatedDateTime
    $AppUpdatedDateTime = $App.ModifiedDateTime
    $AppIdentifierUris = $App.IdentifierUris -join ";"
    $AppReplyUrls = $App.Web.RedirectUris -join ";"
    $AppRequiredResourceAccess = @()
    ForEach ($RRA in $App.RequiredResourceAccess) {
        $ResourceAppId = $RRA.ResourceAppId
        ForEach ($RA in $RRA.ResourceAccess) {
            $RAEntry = "{0}:{1}" -f $ResourceAppId, $RA.Id
            $AppRequiredResourceAccess += $RAEntry
        }
    }
    $AppRequiredResourceAccessString = $AppRequiredResourceAccess -join ";"
    
    # Get service principals for the app
    [Array]$SPs = Get-MgServicePrincipal -Filter "appId eq '$AppId'" -All
    If (!$SPs) {
        Write-Output ("No service principals found for application {0}" -f $AppName)
        Continue
    } Else {
        Write-Output ("{0} service principals found for application {1}" -f $SPs.Count, $AppName)  
    }   
    
    ForEach ($SP in $SPs) {
        $SPObjectId = $SP.Id
        $SPOwners = (Get-MgServicePrincipalOwner -ServicePrincipalId $SPObjectId -All | Select-Object -ExpandProperty UserPrincipalName) -join ";"
        $SPCreatedDateTime = $SP.CreatedDateTime
        $SPUpdatedDateTime = $SP.AppRoleAssignmentRequired
        # Get app role assignments for the service principal
        [Array]$ARAs = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $SPObjectId -All
        If (!$ARAs) {
            Write-Output ("No app role assignments found for service principal {0}" -f $SP.DisplayName)
            Continue
        } Else {
            Write-Output ("{0} app role assignments found for service principal {1}" -f $ARAs.Count, $SP.DisplayName)  
        }
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.