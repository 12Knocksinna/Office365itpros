# Report-NonMFASignIns.PS1
# A script to show how to use the Microsoft Graph PowerShell SDK to report on non-MFA sign-ins in Microsoft 365.
# Showing how to use https://learn.microsoft.com/en-us/graph/api/authenticationmethodsroot-list-usermfasigninsummary?view=graph-rest-beta

# version 1.0 11-Aug-2025
# Github link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-NonMFASignIns.PS1

Connect-MgGraph -Scopes "AuditLog.Read.All"

$Uri = 'https://graph.microsoft.com/beta/reports/authenticationMethods/userMfaSignInSummary' 
[array]$Data = Invoke-MgGraphRequest -Uri $Uri -Method Get -OutputType PSObject | Select-Object -ExpandProperty Value
If ($Data) {
    Write-Output "Processing MFA summary sign-in data"
} Else {
    Write-Output "No non-MFA sign-ins found" -ForegroundColor Green
    Break
}

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Item in $Data) {

    $NonMFAUsers = @()
    $NonMFAApps = @()   
    If ($Item.singleFactorSignIns -gt 0) {
        $StartDate = (Get-Date $Item.CreatedDateTime).toString('yyyy-MM-ddT00:00:00Z')  
        $EndDate = (Get-Date $Item.CreatedDateTime).toString('yyyy-MM-ddT23:59:59Z')
        Write-Host ("Checking sign-in records from {0} to {1}" -f $StartDate, $EndDate)
        [array]$SignInRecords = Get-MgBetaAuditLogSignIn -Filter "createdDateTime gt $StartDate and createdDateTime lt $EndDate and AuthenticationRequirement eq 'singleFactorAuthentication' and status/errorCode eq 0" -Sort "createdDateTime DESC"
        If ($SignInRecords.Count -gt 0) {
            Write-Host ("Found {0} non-MFA sign-in records for {1}" -f $SignInRecords.Count, (Get-Date $Item.CreatedDateTime -format 'dd-MMM-yyyy'))
            [array]$SignInRecords = $SignInRecords | Sort-Object CreatedDate
            [array]$NonMFAUsers = $SignInRecords | Group-Object UserPrincipalName -NoElement | Select-Object -ExpandProperty Name
            [array]$NonMFAApps = $SignInRecords | Group-Object AppDisplayName -NoElement | Select-Object -ExpandProperty Name
        } Else {
            Write-Host "No non-MFA sign-in records found" -ForegroundColor Green
        }

    }

    $ReportLine = [PSCustomObject][Ordered]@{
        Date            = Get-Date $Item.CreatedDateTime -format 'dd-MMM-yyyy'
        'Total Signins' = $Item.TotalSignIns
        'Non MFA'       = $Item.singleFactorSignIns
        'MFA Signins'   = $Item.multiFactorSignIns
        'Non MFA Percentage' = "{0:P2}" -f ($Item.singleFactorSignIns / $Item.TotalSignIns)
        'Non MFA Users' = $NonMFAUsers -Join ", "
        'Non MFA Apps'  = $NonMFAApps -Join ", "
    }
    $Report.Add($ReportLine)
    
}

$Report | Out-GridView -Title "Non-MFA Sign-In Report"

Write-Host "Generating an Excel or CSV report for non-MFA sign-ins"
# Generate reports
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    $ExcelTitle = ("Non-MFA sign-ins for period {0} to {1}" -f $StartDate, $EndDate)
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $OutputXLSXFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\NonMFASignIns.xlsx"
    If (Test-Path $OutputXLSXFile) {
        Remove-Item $OutputXLSXFile -ErrorAction SilentlyContinue
    }
    $Report | Export-Excel -Path $OutputXLSXFile -WorksheetName "Non-MFA Sign-Ins" -Title $ExcelTitle -TitleBold -TableName "NonMFASignIns" -AutoSize
} Else {
    $OutputCSVFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\NonMFASignIns.csv"
    $Report | Export-Csv -Path $OutputCSVFile -NoTypeInformation -Encoding Utf8
}
  
If ($ExcelGenerated) {
    Write-Host ("An Excel worksheet containing the report data is available in {0}" -f $OutputXLSXFile)
} Else {
    Write-Host ("A CSV file containing the report data is available in {0}" -f $OutputCSVFile)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.