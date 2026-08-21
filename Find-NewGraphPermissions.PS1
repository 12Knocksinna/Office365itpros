# Find-NewGraphPermissions.PS1

# V1.0 21-Aug-2026
# GitHub Link: 

# Check the app roles defined for the Graph application and compare them to previous permissions. The input file for the comparison is
# generated from the current set of Graph permissions. It is checked against files from the previous run that are stored in SharePoint Online.
# If any new app roles or OAuth scopes are found, the current files are uploaded to SharePoint Online to replace the previous files
# and a report is generated that's emailed to tenant admins. 
# The script is intended to be run on a regular basis to check for changes in Graph permissions. The script is best run as a scheduled
# Azure Automation runbook.

# Permissions required to run the script:
# Application.Read.All - Read the Graph service principal to fetch app roles
# Sites.Read.All - Read access to SharePoint site to fetch previous comparison files 
# Files.ReadWrite.All - Write access to SharePoint site to upload new comparison files if changes are found

Connect-MgGraph -Scopes "Application.Read.All", "Sites.Read.All", "Files.ReadWrite.All" -NoWelcome

Write-Output "Checking Microsoft Graph permissions for new app roles and OAuth scopes..."
# Get Graph service principal
$GraphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

[array]$AppRoles = $GraphSp.AppRoles | Select-Object Value, DisplayName, Id
[array]$OauthScopes = $GraphSp.Oauth2PermissionScopes | Select-Object Value, AdminConsentDisplayName, Id
    
# Create CSV file names for the current Graph App Roles and OAuth scopes in the Downloads folder
$OutputAppRolesCsv = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\OutputAppRoles.csv"
$OutputOauthScopesCsv = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\OutputOauthScopes.csv"

# Remove any existing files before exporting new data
Remove-Item -Path $OutputAppRolesCsv -ErrorAction SilentlyContinue
Remove-Item -Path $OutputOauthScopesCsv -ErrorAction SilentlyContinue
$AppRoles | Export-Csv -Path $OutputAppRolesCsv -NoTypeInformation -Encoding UTF8
$OauthScopes | Export-Csv -Path $OutputOauthScopesCsv -NoTypeInformation -Encoding UTF8

Write-Output ("Graph App Roles: " + $AppRoles.Count)
Write-Output ("Graph OAuth Scopes: " + $OauthScopes.Count)  

# Change the SharePoint site URL to match whatever site is used to hold the previous comparison files. The site must be accessible to the
# signed-in account running the script (not a problem for app-only or Azure Automation).
Write-Host "Fetching data from previous comparison from SharePoint site..."
$Uri = "https://redmondassociates.sharepoint.com/sites/O365ExchPro"
$SiteId = $Uri.Split('//')[1].split("/")[0] + ":/sites/" + $Uri.Split('//')[1].split("/")[2]

# Get site and default document library
Try {
    $Site = Get-MgSite -SiteId $SiteId -ErrorAction Stop
} Catch {
    Write-Error "Failed to retrieve site information: $_"
    Return
}
# Get the default document library for the site
Try {
    $Drive = Get-MgSiteDefaultDrive -SiteId $Site.Id -ErrorAction Stop
} Catch {
    Write-Error "Failed to retrieve default document library: $_"
    Return
}    

# Make sure that c:\temp exists
If (-Not (Test-Path -Path "C:\Temp" -PathType Container)) {
    Try {
        New-Item -Path "C:\Temp" -ItemType Directory -ErrorAction Stop | Out-Null
    } Catch {
        Write-Error "Failed to create C:\Temp directory: $_"
        Return
    }
}

# Delete old download files if they exist
Remove-Item -Path "C:\Temp\OutputAppRoles.csv" -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Temp\OutputOauthScopes.csv" -ErrorAction SilentlyContinue

# Download the App Roles and OAuth Scopes files from the SharePoint site to the local machine for comparison. First get file properties
# because we want to display the information and need the item identifier to upload a new version if there are changes. 

# Change this for the folder in the SharePoint site where the previous comparison files are stored.
$FolderName = "The Base of All Good in Office 365"
$OldAppRolesFileName = "OutputAppRoles.csv"
$OldOauthScopesFileName = "OutputOauthScopes.csv"

$Uri = ("https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}/{2}" -f $Drive.Id, $FolderName, $OldAppRolesFileName)
$AppRolesItem = Invoke-MgGraphRequest -Uri $Uri -Method GET -ErrorAction Stop

$Uri = ("https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}/{2}:/content" -f $Drive.Id, $FolderName, $OldAppRolesFileName)
Try {
    Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath "C:\Temp\OutputAppRoles.csv" -ErrorAction Stop
} Catch {
    Write-Error "Failed to download OutputAppRoles.csv: $_"
    Return
}

$Uri = ("https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}/{2}" -f $Drive.Id, $FolderName, $OldOauthScopesFileName)
$OauthScopesItem = Invoke-MgGraphRequest -Uri $Uri -Method GET -ErrorAction Stop

$Uri = ("https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}/{2}:/content" -f $Drive.Id, $FolderName, $OldOauthScopesFileName)
Try {
    Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath "C:\Temp\OutputOauthScopes.csv" -ErrorAction Stop
} Catch {
    Write-Error "Failed to download OutputOauthScopes.csv: $_"
    Return
}

Write-Output ("Importing files with information about Graph roles dated {0}" -f (Get-Date $AppRolesItem.lastModifiedDateTime -format "dd-MMM-yyyy HH:mm:ss"))

# Import into arrays for comparison
[array]$OldAppRoles = Import-Csv -Path "C:\Temp\OutputAppRoles.csv"
[array]$OldOauthScopes = Import-Csv -Path "C:\Temp\OutputOauthScopes.csv"

# Find if any new app roles have been added since the last comparison
$NewAppRoles = Compare-Object -ReferenceObject $OldAppRoles -DifferenceObject $AppRoles -Property Value, DisplayName, Id -PassThru | Where-Object { $_.SideIndicator -eq "=>" }
$NewOauthScopes = Compare-Object -ReferenceObject $OldOauthScopes -DifferenceObject $OauthScopes -Property Value, AdminConsentDisplayName, Id -PassThru | Where-Object { $_.SideIndicator -eq "=>" }            

If ($NewAppRoles.Count -gt 0 -or $NewOauthScopes.Count -gt 0) {
    Write-Output "New Graph permissions have been added since the last comparison."
    Write-Output ("New App Roles: " + $NewAppRoles.Count)
    $NewAppRoles | Format-Table Value, Id, DisplayName -AutoSize
    Write-Output ("New OAuth Scopes: " + $NewOauthScopes.Count)
    $NewOauthScopes | Format-Table Value, Id, AdminConsentDisplayName -AutoSize
    Write-Output "Uploading the current Graph App Roles and OAuth Scopes to SharePoint site..."
} Else {
    Write-Output "No new Graph permissions have been added since the last comparison."
}

If ($NewAppRoles.Count -gt 0) {
    Try {
        Set-MgDriveItemContent -DriveId $Drive.Id -DriveItemId $AppRolesItem.Id -InFile $OutputAppRolesCsv -ErrorAction Stop | Out-Null
        Write-Output "Updated OutputAppRoles.csv in SharePoint site"
    } Catch {
        Write-Error "Failed to update OutputAppRoles.csv in SharePoint site: $_"
        Return
    }
}

If ($NewOauthScopes.Count -gt 0) {
    Try {
        Set-MgDriveItemContent -DriveId $Drive.Id -DriveItemId $OauthScopesItem.Id -InFile $OutputOauthScopesCsv -ErrorAction Stop | Out-Null
    } Catch {
        Write-Error "Failed to update OutputOauthScopes.csv in SharePoint site: $_"
        Return
    }
}

If ($NewAppRoles.Count -eq 0 -and $NewOauthScopes.Count -eq 0) {
   Continue
}

$NewAppRolesHtml = $NewAppRoles | Select-Object Value, Id, DisplayName | ConvertTo-Html -As Table -Fragment
$NewOauthScopesHtml = $NewOauthScopes | Select-Object Value, Id, AdminConsentDisplayName | ConvertTo-Html -As Table -Fragment

# Send email with the results - define your own email address here and the address to receive the report.
$MsgFrom = "Site.Announcements@office365itpros.com"
$DestinationEmailAddress = "Developers@office365itpros.com"
# Build the array of a single TO recipient detailed in a hash table - change this to the appropriate recipient for your tenant
$ToRecipient = @{}
$ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
[array]$MsgTo = $ToRecipient
# Define the message subject
$MsgSubject = "Important: New Graph Permissions Detected in Your Tenant - Action Required"
# Create the HTML content
$HtmlMsg = "</body></html><p>The following <b>New Graph permissions have been detected in your tenant</b> based on the latest comparison.</p>" 
$HtmlMsg += "<h3>New Graph App Roles</h3>"
$HtmlMsg += $NewAppRolesHtml
$HtmlMsg += "<h3>New Graph OAuth Scopes</h3>"
$HtmlMsg += $NewOauthScopesHtml
$HtmlMsg += "<p>Please review these new permissions and take any necessary actions to ensure your tenant's security and compliance.</p></body></html>"

$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlMsg)")
$MsgBody.Add('ContentType','html')
# Build the parameters to submit the message
$Message = @{}
$Message.Add('subject', $MsgSubject)
$Message.Add('toRecipients', $MsgTo)
$Message.Add('body', $MsgBody)
$Message.Add("attachments", $MsgAttachments)

$EmailParameters = @{}
$EmailParameters.Add('message', $Message)
$EmailParameters.Add('saveToSentItems', $true)
$EmailParameters.Add('isDeliveryReceiptRequested', $true)
Write-Output ""
Write-Output "Sending email..."
Write-Output ""

# Send the message
Try {
    Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters -ErrorAction Stop
    Write-Output ("New Graph permissions notification emailed to {0}" -f $ToRecipient.emailAddress.address)
} Catch {
    Write-Output "Unable to send email"
    Write-Output $_.Exception.Message
}

Write-Output "All done - check email for the report on new Graph permissions."

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.