# Report-ExternalTenantAccessbyGuests.PS1
# This script generates a report of member accounts who have accessed apps in external tenants in a Microsoft 365 environment.

# V1.0 4-Dec-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ExternalTenantAccessbyGuests.PS1

# accessing Entra ID audit logs via Microsoft Graph requires an Entra P1 license

# Scopes required for script
# 	Organization.Read.All (get tenant information)
#   AuditLog.Read.All (read sign-in logs)
# 	CrossTenantInformation.ReadBasic.All (resolve tenant identifier)
[array]$RequiredScopes = "AuditLog.Read.All","Organization.Read.All","CrossTenantInformation.ReadBasic.All"
$Interactive = $false

# Determine if we're interactive or not
If ([Environment]::UserInteractive) { 
    # We're running interactively...
    Clear-Host
    Write-Host "Script running interactively... connecting to the Graph" -ForegroundColor Yellow
    Connect-MgGraph -NoWelcome -Scopes $RequiredScopes
    $Interactive = $true
} Else { 
    # We're not, so likely in Azure Automation
    Write-Output "Executing the runbook to create the last activity report for guest access outside the tenant..." 
    Connect-MgGraph -Identity -NoWelcome
}

# Check that we have the right permissions - in Azure Automation, we assume that the automation account has the right permissions
If ($Interactive) {
    [int]$RequiredScopesCount = $RequiredScopes.Count
    [string[]]$CurrentScopes = (Get-MgContext).Scopes
    [string[]]$RequiredScopes = $RequiredScopes

    $CheckScopes =[object[]][Linq.Enumerable]::Intersect($RequiredScopes,$CurrentScopes)
    If ($CheckScopes.Count -ne $RequiredScopesCount ) { 
        Write-Host ("To run this script, you need to connect to Microsoft Graph with the following scopes: {0}" -f $RequiredScopes) -ForegroundColor Red
        Break
    }
}

# Define the date range for the report (last 30 days) - Reduce this is it takes too long to fetch audit sign-in records
$StartDate = (Get-Date).AddDays(-30).ToString('yyyy-MM-ddT00:00:00Z')
$EndDate  = (Get-Date).ToString('yyyy-MM-ddT23:59:59Z')

# Find this tenant information
Try {
    $ThisTenantData = Get-MgOrganization -ErrorAction Stop
    $ThisTenantId = $ThisTenantData.Id
} catch {
    Write-Warning "Failed to retrieve tenant information for this organization: $($_.Exception.Message)"
    Return
}

# Build OData filter string safely
$Filter = "createdDateTime ge $StartDate and createdDateTime le $EndDate and status/errorCode eq 0"

Write-Output "Fetching sign-in audit records from $StartDate to $EndDate..."
# Get all successful sign-in audit records within the specified date range
Try {
    [array]$AuditRecords = Get-MgBetaAuditLogSignIn -All -Filter $Filter -ErrorAction Stop
} Catch {
    Write-Error "Failed to retrieve sign-in logs: $($_.Exception.Message)"
    Return
}

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
        'Access type'           = 'Outbound'
        'Number of Sign-ins'    = $TenantRecords.Count
        'Unique Users'          = $TenantUsers.Count   
        'Users'                 = ($TenantUsers -join '; ')
        'Apps Accessed'         = ($TenantApps -join '; ')
    }
    $TenantData.Add($TenantDataLine)
}

Write-Output "Extracting information for inbound tenants..."
$TenantUsers = $InboundAccessRecords | Select-Object -ExpandProperty UserPrincipalName -Unique
$TenantApps = $InboundAccessRecords | Select-Object -ExpandProperty AppDisplayName -Unique
$TenantDataLine = [PSCustomObject]@{
    TenantId                = $ThisTenantData.Id
    TenantName              = $ThisTenantData.DisplayName
    'Access type'           = 'Inbound'  
    'Number of Sign-ins'    = $InboundAccessRecords.Count
    'Unique Users'          = $TenantUsers.Count   
    'Users'                 = ($TenantUsers -join '; ')
    'Apps Accessed'         = ($TenantApps -join '; ')
}
$TenantData.Add($TenantDataLine)

Write-Output ""
Write-Output "Report of External Tenant Access by Guest Users"
$TenantData | Select-Object TenantName, Type, 'Number of Sign-ins', 'Unique Users', Users, 'Apps Accessed' | Format-Table -AutoSize

# Export the report as an Excel worksheet or CSV file
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $true
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\ExternalTenantAccessReport.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    } 
    $TenantData | Export-Excel -Path $ExcelOutputFile -WorksheetName "External Tenant Access" -Title ("External Tenant Access Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "ExternalTenantAccess" -AutoSize -AutoFilter
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\ExternalTenantAccessReport.CSV"
    $TenantData  | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}

If ($ExcelGenerated) {
    Write-Output ("Excel worksheet output written to {0}" -f $ExcelOutputFile)
} Else {
    Write-Output ("CSV output file written to {0}" -f $CSVOutputFile)
} 

Write-Output "All done - enjoy the External Tenant Access data!"



# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.