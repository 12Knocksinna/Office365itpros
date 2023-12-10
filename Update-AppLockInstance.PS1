# Update-AppLockInstance.PS1
# Update Entra ID applications that don't have app instance lock set
# V1.0 8-Dec-2023
# https://github.com/12Knocksinna/Office365itpros/blob/master/Update-AppLockInstance.PS1

Connect-MgGraph -NoWelcome -Scopes Application.ReadWrite.All

Write-Host "Finding Entra ID applications..."
[array]$Apps = Get-MgApplication -All | Sort-Object DisplayName
Write-Host ("Found {0} applications. Now checking the app instance lock" -f $Apps.count)
# Create the hash table with the properties to update
$AppInstanceLockConfiguration = @{}
$AppInstanceLockConfiguration.Add("isEnabled",$true)
$AppInstanceLockConfiguration.Add("AllProperties",$true)

$Report = [System.Collections.Generic.List[Object]]::new()
# Loop through the apps and update each one that doesn't have app instance lock set
ForEach ($App in $Apps) {
  $ServiceLock = $App | Select-Object -ExpandProperty ServicePrincipalLockConfiguration
  Write-Host ("Now processing {0}" -f $App.displayName)
  If ($ServiceLock.IsEnabled -eq $True) {
    Write-Host ("The {0} app is already enabled" -f $App.displayName) -ForegroundColor Red
  } Else {
    Write-Host ("App Instance Property Lock Not enabled for {0}; updating app" -f $App.displayName)
    Update-MgApplication -ApplicationId $App.Id -ServicePrincipalLockConfiguration $AppInstanceLockConfiguration 
    $AppData = [PSCustomObject][Ordered]@{
        Timestamp       = Get-Date -Format s
        App             = $App.displayName
        AppId           = $App.Id
        Created         = $App.CreatedDateTime
        SignInAudience  = $App.SignInAudience
        Credentials     = ($App.PasswordCredentials.DisplayName -join ", ")
    }                                                         
    $Report.Add($AppData)
  }
}

$Report | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
