# Find-TeamsForUser.PS1
# V1.0 26-Feb-2026

# An example of using the Microsoft Graph PowerShell SDK to generate a report of Teams memberships for selected users. 
# GitHub Link:

Connect-MgGraph -NoWelcome -Scopes Group.Read.All, Groups.Read.All, Team.ReadBasic.All

$GroupId = (Get-MgGroup -Filter "displayName eq 'Users to Monitor for Teams Membership'").Id
[array]$Users = Get-MgGroupMember -GroupId $GroupId

$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($User in $Users) {
    Write-Host ("Fetching details of Teams membership for user {0} ({1}):" -f $User.additionalProperties.displayName, $User.additionalProperties.userPrincipalName) -ForegroundColor Yellow
  
    [array]$Memberships = Get-MgUserJoinedTeam -UserId $User.Id | Sort-Object DisplayName
    If ($Memberships.Count -gt 0) {
        Write-Host ("Found {0} Teams group memberships for {1}." -f $Memberships.Count, $User.additionalProperties.displayName) -ForegroundColor Green
        ForEach ($Membership in $Memberships) {
            $ReportLine = [PSCustomObject][Ordered]@{  
                UserId              = $User.Id
                User                = $User.additionalProperties.displayName
                UserUPN             = $User.additionalProperties.userPrincipalName
                'Job Title'         = $User.additionalProperties.jobTitle
                'Office'            = $User.additionalProperties.officeLocation
                'Team'              = $Membership.DisplayName
                'TeamId'            = $Membership.Id
                'Timestamp'         = Get-Date -format 'dd-MMM-yyyy HH:mm:ss'
            }
            $Report.Add($ReportLine)
        }
    } Else {
        Write-Host "No Teams group memberships found for this user." -ForegroundColor Yellow
    }
    Write-Host ""
}

# Generate the report in either Excel worksheet or CSV format, depending on if the ImportExcel module is available
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Teams Memberships.xlsx"
    $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Teams Memberships" -Title ("Teams Memberships {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "TeamsMemberships" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Teams Memberships.CSV"
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}   

If ($ExcelGenerated -eq $true) {
    Write-Host ("Teams memberships report is available in Excel workbook {0}" -f $ExcelOutputFile)
} Else {
    Write-Host ("Teams memberships report is available in CSV file {0}" -f $CSVOutputFile)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.