# Report-CopilotLicensedUsers.PS1
# A script to generate a comprehensive report of all users who have Microsoft 365 Copilot licenses assigned
# This script identifies users with Copilot licenses, shows their license assignment details, and outputs 
# a formatted report suitable for tenant administrators.
# V1.0 08-Dec-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-CopilotLicensedUsers.PS1

# Connect to Microsoft Graph
# User.Read.All - to read user account details and license information
# AuditLog.Read.All - to read sign-in activity data
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -NoWelcome -Scopes User.Read.All, AuditLog.Read.All

# Microsoft 365 Copilot SKU identifiers
# First SKU is the standard Microsoft 365 Copilot license
# Second SKU is for trial Copilot licenses
[array]$CopilotSKUs = @(
    @{SkuId = '639dec6b-bb19-468b-871c-c5c441c4b0cb'; Name = 'Microsoft 365 Copilot'},
    @{SkuId = 'ea2d19f9-23bc-4f7f-9c12-a677b26a8e2a'; Name = 'Microsoft 365 Copilot (Trial)'}
)

# Service plan identifiers for Copilot features
$ServicePlanNames = @{
    'fe6c28b3-d468-44ea-bbd0-a10a5167435c' = 'Copilot Studio'
    '0aedf20c-091d-420b-aadf-30c042609612' = 'Copilot for SharePoint'
    '82d30987-df9b-4486-b146-198b21d164c7' = 'Copilot Graph Connectors'
    '89f1c4c8-0878-40f7-804d-869c9128ab5d' = 'Copilot Power Platform Connectors'
    'a62f8878-de10-42f3-b68f-6149a25ceb97' = 'Copilot in Productivity Apps'
    'b95945de-b3bd-46db-8437-f2beb6ea2347' = 'Copilot for Teams'
    '3f30311c-6b1e-48a4-ab79-725b469da960' = 'Copilot with Graph-grounded Chat'
    '931e4a88-a67f-48b5-814f-16a5f1e6028d' = 'Copilot with Intelligent Search'
}

Write-Host "Searching for users with Microsoft 365 Copilot licenses..." -ForegroundColor Cyan

# Build filter to find users with any of the Copilot SKUs
$FilterParts = @()
ForEach ($Sku in $CopilotSKUs) {
    $FilterParts += "assignedLicenses/any(s:s/skuId eq $($Sku.SkuId))"
}
$Filter = $FilterParts -join ' or '

# Get all users with Copilot licenses
[array]$Users = Get-MgUser -Filter $Filter `
    -ConsistencyLevel Eventual -CountVariable LicenseCount -All `
    -Property Id, DisplayName, UserPrincipalName, Mail, Department, JobTitle, Country, `
              SignInActivity, AssignedLicenses, LicenseAssignmentStates, AccountEnabled `
    -Sort 'DisplayName' -PageSize 999

If (!$Users) {
    Write-Host "No users with Microsoft 365 Copilot licenses found in this tenant." -ForegroundColor Yellow
    Break
} Else {
    Write-Host ("{0} users with Microsoft 365 Copilot licenses found" -f $Users.Count) -ForegroundColor Green
}

Write-Host "Processing user license details..." -ForegroundColor Cyan

# Create a hash table to cache group information to avoid repeated API calls
$GroupCache = @{}

$Report = [System.Collections.Generic.List[Object]]::new()
$UserCount = 0

ForEach ($User in $Users) {
    $UserCount++
    Write-Host ("Processing user {0} of {1}: {2}" -f $UserCount, $Users.Count, $User.DisplayName) -NoNewline
    
    # Determine which Copilot SKU(s) the user has
    $UserCopilotLicenses = @()
    $UserCopilotSkuIds = @()
    ForEach ($Sku in $CopilotSKUs) {
        If ($User.AssignedLicenses.SkuId -contains $Sku.SkuId) {
            $UserCopilotLicenses += $Sku.Name
            $UserCopilotSkuIds += $Sku.SkuId
        }
    }
    
    # Get license assignment details
    $LicenseDetails = Get-MgUserLicenseDetail -UserId $User.Id | Where-Object {$_.SkuId -in $UserCopilotSkuIds}
    
    # Check if license is assigned directly or via group
    $AssignmentType = @()
    $AssigningGroups = @()
    ForEach ($SkuId in $UserCopilotSkuIds) {
        $LicenseState = $User.LicenseAssignmentStates | Where-Object {$_.SkuId -eq $SkuId}
        If ($LicenseState.AssignedByGroup) {
            $AssignmentType += "Group-based"
            $GroupId = $LicenseState.AssignedByGroup
            # Check cache first to avoid repeated API calls
            If ($GroupCache.ContainsKey($GroupId)) {
                $AssigningGroups += $GroupCache[$GroupId]
            } Else {
                Try {
                    $Group = Get-MgGroup -GroupId $GroupId -ErrorAction SilentlyContinue
                    $GroupName = If ($Group) { $Group.DisplayName } Else { $GroupId }
                    $GroupCache[$GroupId] = $GroupName
                    $AssigningGroups += $GroupName
                } Catch {
                    $GroupCache[$GroupId] = $GroupId
                    $AssigningGroups += $GroupId
                }
            }
        } Else {
            $AssignmentType += "Direct"
        }
    }
    
    # Get enabled service plans for Copilot
    $EnabledPlans = @()
    $DisabledPlans = @()
    ForEach ($License in $LicenseDetails) {
        ForEach ($Plan in $License.ServicePlans) {
            If ($ServicePlanNames.ContainsKey($Plan.ServicePlanId)) {
                If ($Plan.ProvisioningStatus -eq 'Success') {
                    $EnabledPlans += $ServicePlanNames[$Plan.ServicePlanId]
                } Else {
                    $DisabledPlans += $ServicePlanNames[$Plan.ServicePlanId]
                }
            }
        }
    }
    
    # Get sign-in information
    $LastSignIn = $null
    $DaysSinceSignIn = "N/A"
    If ($User.SignInActivity.LastSuccessfulSignInDateTime) {
        $LastSignIn = Get-Date $User.SignInActivity.LastSuccessfulSignInDateTime -format 'dd-MMM-yyyy HH:mm:ss'
        $DaysSinceSignIn = (New-TimeSpan $User.SignInActivity.LastSuccessfulSignInDateTime).Days
    } ElseIf ($User.SignInActivity.LastSignInDateTime) {
        $LastSignIn = Get-Date $User.SignInActivity.LastSignInDateTime -format 'dd-MMM-yyyy HH:mm:ss'
        $DaysSinceSignIn = (New-TimeSpan $User.SignInActivity.LastSignInDateTime).Days
    } Else {
        $LastSignIn = "Never"
    }
    
    # Determine account status
    $Status = If ($User.AccountEnabled) { "Enabled" } Else { "Disabled" }
    
    $ReportLine = [PSCustomObject][Ordered]@{
        DisplayName            = $User.DisplayName
        UserPrincipalName      = $User.UserPrincipalName
        Mail                   = $User.Mail
        Department             = $User.Department
        JobTitle               = $User.JobTitle
        Country                = $User.Country
        'Account Status'       = $Status
        'Copilot License(s)'   = ($UserCopilotLicenses -join '; ')
        'Assignment Type'      = ($AssignmentType | Select-Object -Unique) -join '; '
        'Assigned By Group(s)' = ($AssigningGroups -join '; ')
        'Enabled Plans'        = ($EnabledPlans | Select-Object -Unique | Sort-Object) -join '; '
        'Disabled Plans'       = ($DisabledPlans | Select-Object -Unique | Sort-Object) -join '; '
        'Plans Enabled Count'  = ($EnabledPlans | Select-Object -Unique).Count
        'Last Sign-In'         = $LastSignIn
        'Days Since Sign-In'   = $DaysSinceSignIn
    }
    $Report.Add($ReportLine)
    Write-Host " - Done" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Generating report output..." -ForegroundColor Cyan

# Generate summary statistics
$TotalUsers = $Report.Count
$DirectAssigned = ($Report | Where-Object {$_.'Assignment Type' -eq 'Direct'}).Count
$GroupAssigned = ($Report | Where-Object {$_.'Assignment Type' -eq 'Group-based'}).Count
$DisabledAccounts = ($Report | Where-Object {$_.'Account Status' -eq 'Disabled'}).Count
$NeverSignedIn = ($Report | Where-Object {$_.'Last Sign-In' -eq 'Never'}).Count

# Determine output file path - use Downloads folder
$OutputPath = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path)

# Try to export to Excel if ImportExcel module is available, otherwise use CSV
$ExcelGenerated = $False
If (Get-Module ImportExcel -ListAvailable) {
    Try {
        Import-Module ImportExcel -ErrorAction Stop
        $ExcelOutputFile = Join-Path $OutputPath "Copilot-Licensed-Users-Report.xlsx"
        
        # Remove existing file if it exists
        If (Test-Path $ExcelOutputFile) {
            Remove-Item $ExcelOutputFile -Force -ErrorAction SilentlyContinue
        }
        
        # Export to Excel with formatting
        $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Copilot Licensed Users" `
            -Title ("Microsoft 365 Copilot Licensed Users Report - {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm')) `
            -TitleBold -TableName "CopilotUsers" -AutoSize -FreezeTopRow -BoldTopRow
        
        $ExcelGenerated = $True
        Write-Host ("Excel report generated: {0}" -f $ExcelOutputFile) -ForegroundColor Green
    } Catch {
        Write-Host "Failed to generate Excel report. Falling back to CSV..." -ForegroundColor Yellow
    }
}

If (-not $ExcelGenerated) {
    $CSVOutputFile = Join-Path $OutputPath "Copilot-Licensed-Users-Report.csv"
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding UTF8
    Write-Host ("CSV report generated: {0}" -f $CSVOutputFile) -ForegroundColor Green
}

# Display summary on screen
Clear-Host
Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "   Microsoft 365 Copilot Licensed Users - Summary Report" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host ("Report generated: {0}" -f (Get-Date -format 'dd-MMM-yyyy HH:mm:ss')) -ForegroundColor White
Write-Host ""
Write-Host "LICENSING STATISTICS:" -ForegroundColor Yellow
Write-Host ("  Total users with Copilot licenses: {0}" -f $TotalUsers) -ForegroundColor White
Write-Host ("  Direct license assignments:        {0}" -f $DirectAssigned) -ForegroundColor White
Write-Host ("  Group-based assignments:           {0}" -f $GroupAssigned) -ForegroundColor White
Write-Host ""
Write-Host "ACCOUNT STATUS:" -ForegroundColor Yellow
Write-Host ("  Disabled accounts:                 {0}" -f $DisabledAccounts) -ForegroundColor White
Write-Host ("  Never signed in:                   {0}" -f $NeverSignedIn) -ForegroundColor White
Write-Host ""
Write-Host "OUTPUT:" -ForegroundColor Yellow
If ($ExcelGenerated) {
    Write-Host ("  Excel report: {0}" -f $ExcelOutputFile) -ForegroundColor White
} Else {
    Write-Host ("  CSV report: {0}" -f $CSVOutputFile) -ForegroundColor White
}
Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Display top 10 users for quick review
Write-Host "SAMPLE OF LICENSED USERS (First 10):" -ForegroundColor Yellow
$Report | Select-Object -First 10 DisplayName, UserPrincipalName, 'Copilot License(s)', 'Assignment Type', 'Plans Enabled Count', 'Last Sign-In' | Format-Table -AutoSize

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
