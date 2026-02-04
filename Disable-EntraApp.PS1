# Disable-EntraApp.PS1
# Demonstrate how to disable (deactivate) an Entra application by setting the isDisabled property using Microsoft Graph. This is only possible
# for applications registered in the tenant. Enterprise applications owned by other tenants cannot be disabled.

# V1.0 04-Feb-2026
# GitHub Link:  https://github.com/12Knocksinna/Office365itpros/blob/master/Disable-EntraApp.PS1

# Connect to Microsoft Graph
If (!(Get-MgContext).Account) {
    Write-Host "Connecting to Microsoft Graph..."
    Connect-MgGraph -NoWelcome -Scopes Application.ReadWrite.All 
    #  Application.ReadWrite.All needed to update app properties
}   

$Application = Read-Host "Enter the name of the Entra application to deactivate"
$Application = $Application.Trim()
If ($Application -eq "") {
    Write-Host "No application name provided - exiting"
    Break
}

# Using a Graph request to find the application because Get-MgApplication or Get-MgBetaApplication don't return the isDisabled property
# When the Graph SDK is updated, we should be able to use:
# $ApplicationObject = Get-MgBetaApplication -Filter "displayName eq '$Application'"
$Uri = ("https://graph.microsoft.com/beta/applications/?`$filter=displayname eq '{0}'" -f $Application)
[array]$Data = Invoke-MgGraphRequest -Uri $Uri -Method Get -OutputType PsObject 

If ($Data.Value.Count -gt 1) {
    Write-Host ("Multiple applications found with the name {0}. Please refine your search." -f $Application)
    Break
} Else {
    $ApplicationObject = $Data.Value
}

If (!$ApplicationObject) {
    Write-Host ("No application found with the name {0}." -f $Application)
    Break
} Else {
    # Is the application already disabled?
    If ($ApplicationObject.IsDisabled -eq $true) {
        Write-Host ("The {0} application is already disabled." -f $ApplicationObject.DisplayName)
        Break
    }
    Write-Host ("Deactivating application {0} (ObjectId: {1})..." -f $ApplicationObject.DisplayName, $ApplicationObject.Id)
}

# Propulate variables for the update
$NewDisplayName = $ApplicationObject.displayName + " (disabled)"
$NewNotes = ($ApplicationObject.notes.trim()) + "`nApp disabled on " + (Get-Date -format "dd-MMM-yyyy HH:mm") + " by " + (Get-MgContext).account

# Create the request body
$RequestBody = @{}
$RequestBody.Add("isDisabled", $true)
$RequestBody.Add("displayName", $NewDisplayName)
$RequestBody.Add("notes", $NewNotes)

# First, remove application owners
[array]$Owners = Get-MgApplicationOwner -ApplicationId $ApplicationObject.Id -All 
ForEach ($Owner in $Owners) {
    Try {
        Remove-MgApplicationOwnerDirectoryObjectByRef -ApplicationId $ApplicationObject.Id -DirectoryObjectId $Owner.Id
        Write-Host ("Removed owner {0} from application..." -f $Owner.DisplayName)
    } Catch {
        Write-Host ("Failed to remove owner {0} from application. Error details:" -f $Owner.DisplayName)
        Write-Host $_.Exception.Message
    }
}   

# Update the application
Try {
    Update-MgApplication -ApplicationId $ApplicationObject.Id -BodyParameter $RequestBody  
    Write-Host ("{0} application deactivated." -f $ApplicationObject.DisplayName)  
} Catch {
    Write-Host "Failed to disable application. Error details:"
    Write-Host $_.Exception.Message
} 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 