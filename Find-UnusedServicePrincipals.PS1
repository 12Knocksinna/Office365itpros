# Find-UnusedServicePrincipals.PS1
# Find service principals that have not signed in to Microsoft 365 in the last year and generate some statistics
# about the service principals in the tenant.

# V1.0 20-Nov-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-UnusedServicePrincipals.PS1

Connect-MgGraph -Scopes AuditLog.Read.All, Application.Read.All

Write-Host "Finding service principals..."
[array]$ServicePrincipals = Get-MgServicePrincipal -All -PageSize 500 | Sort-Object AppId
$SPHash = @{}
ForEach ($SP in $ServicePrincipals) {
    $SPHash.Add($SP.AppId, $SP.DisplayName)
}

$CheckDate = (Get-Date).AddYears(-1).toString('yyyy-MM-ddTHH:mm:ssZ')
# Output report
Write-Host "Fetching service principal sign-in activity data..."
$Report = [System.Collections.Generic.List[Object]]::new()
[array]$SPSignInLogs = Get-MgBetaReportServicePrincipalSignInActivity -Filter "(lastSignInActivity/lastSignInDateTime lt $CheckDate)" -All -PageSize 500
If (!($SPSignInLogs)) {
    Write-Host "No sign-ins found for service principals"
    Break
}   Else {
    Write-Host ("Found {0} sign-ins for service principals" -f $SPSignInLogs.Count)
}

Write-Host "Analyzing data..."
ForEach ($SPSignIn in $SPSignInLogs) {
    $SPName = $SPHash[$SPSignIn.appId]
    $DaysSince = (New-TimeSpan $SPSignIn.lastSignInActivity.lastSignInDateTime).Days
    $ReportLine = [PSCustomObject]@{
        'Service Principal Name'    = $SPName
        AppId                       = $SPSignin.AppId
        LastSignIn                  = Get-Date $SPSignIn.lastSignInActivity.lastSignInDateTime -format 'dd-MMM-yyyy HH:mm:ss'
        'Days Since Last Sign-In'   = $DaysSince
    }
    $Report.Add($ReportLine)
}

$TenantReport = [System.Collections.Generic.List[Object]]::new()
$HomeTenant = (Get-MgOrganization).DisplayName
[array]$TenantIds = $ServicePrincipals | Sort-Object AppOwnerOrganizationId -Unique | Select-Object -ExpandProperty AppOwnerOrganizationId
ForEach ($TenantId in $TenantIds) {
    $NumberApps = ($ServicePrincipals | Where-Object {$_.AppOwnerOrganizationId -eq $TenantId}).Count
    $Uri = ("https://graph.microsoft.com/V1.0/tenantRelationships/findTenantInformationByTenantId(tenantId='{0}')" -f $TenantId.ToString())
    $TenantData = Invoke-MgGraphRequest -Uri $Uri -Method Get
    $ReportLine = [PSCustomObject]@{
        'Tenant Name'   = $TenantData.DisplayName
        'Tenant ID'     = $TenantId
        'Number of Apps'= $NumberApps
    }
    $TenantReport.Add($ReportLine)
}

$TenantReport = $TenantReport | Sort-Object 'Number of apps' -Descending
[array]$AppsNoName = $Report | Where-Object {$_.'Service Principal Name' -eq $null}

Write-Host ("Some notes about service principals for the {0} tenant" -f $HomeTenant)
Write-Host "------------------------------------------------------------------------"
Write-Host ""
Write-Host "Service Principals by owning tenant"
$TenantReport | Format-Table -AutoSize
Write-Host ""
Write-Host ("Total Service Principals {0}" -f $ServicePrincipals.Count)
Write-Host ("Service Principals with no sign-ins in the last year {0}" -f $Report.Count)
Write-Host ("Service Principals with sign-ins in the last year {0}" -f ($ServicePrincipals.Count - $Report.Count))
Write-Host ("Number of apps with no service principal {0}" -f $AppsNoName.Count)
Write-Host ""

# Generate some reports
$Report | Out-GridView -Title "Service Principals with no sign-ins in the last year"
$Report | Export-CSV -Path ServicePrincpalsNoSignIn.csv -NoTypeInformation -Encoding UTF8
Write-Host "Report detailing service principals with last sign-in longer than a year written to ServicePrincipalsNoSignIn.csv"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.