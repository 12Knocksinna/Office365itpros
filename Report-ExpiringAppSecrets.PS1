# Report-ExpiringAppSecrets.PS1
# An example of using a script to check Entra ID registered apps to see if any have app secrets that are due to expire
# and inform admins of the fact.
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ExpiringAppSecrets.PS1
# V1.0 2-Jan-2023

Connect-MgGraph -Scopes 'Application.Read.All' -NoWelcome

$CheckDate = Get-Date
# Define the warning period to check for app secrets that are about to expire
[int]$ExpirationWarningPeriod = 30

# Find registered Entra ID apps that are limited to our organization (not multi-organization)
[array]$RegisteredApps = Get-MgApplication -All | Where-Object SignInAudience -match "AzureADMyOrg" | Sort-Object DisplayName
# Remove SharePoint helper apps https://learn.microsoft.com/en-us/answers/questions/1187017/sharepoint-online-client-extensibility-web-applica
$RegisteredApps = $RegisteredApps | Where-Object DisplayName -notLike "SharePoint Online Client Extensibility Web Application Principal*"

If (!($RegisteredApps)) {
    Write-Host "Can't retrieve details of any Entra ID registered apps - exiting"
    Break
} Else {
    Write-Host ("{0} registered applications found - proceeeding to analyze app secrets" -f $RegisteredApps.count)
}

$Report = [System.Collections.Generic.List[Object]]::new() 
ForEach ($App in $RegisteredApps) {
    Write-Host ("Processing {0} app" -f $App.DisplayName)
    $Uri = ("https://graph.microsoft.com/v1.0/applications/{0}/passwordCredentials" -f $App.Id)
    [array]$AppSecrets = Invoke-MgGraphRequest -Method GET -Uri $Uri
    [array]$AppSecrets = $AppSecrets.Value

    ForEach ($AppSecret in $AppSecrets) {
        $ExpirationDays = $null; $Status = $null
        If ($null -ne $AppSecret.endDateTime) {
            $ExpirationDays = (New-TimeSpan -Start $CheckDate -End $AppSecret.endDateTime).Days
            If ($ExpirationDays -lt 0) {
                $Status = "Expired"
            } ElseIf ($ExpirationDays -gt 0 -and $ExpirationDays -le $ExpirationWarningPeriod) {
                $Status = "Secret Expires Soon"
            } Else {
                $Status = "Active"
            }
            $DataLine = [PSCustomObject] @{
                "App Name"          = $App.DisplayName
                "App Id"            = $App.Id
                "Secret"            = $AppSecret.DisplayName
                "Secret creation"   = $AppSecret.startDateTime
                "Secret KeyId"      = $AppSecret.KeyId
                "Secret expiration" = $AppSecret.endDateTime
                "Days Until Expiry" = $ExpirationDays
                Status              = $Status
                }
            $Report.Add($DataLine)
        }
    }
}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 