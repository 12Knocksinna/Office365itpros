# FindPotentialDirectoryProblems.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindPotentialDirectoryProblems.PS1
# Quick and dirty script to highlight potential issues which might exist in an Office 365 tenant directory
# 
# Get list of user mailboxes in the directory
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
Write-Host "Finding User mailboxes..."
$Directory = Get-User -RecipientTypeDetails UserMailbox -ResultSize Unlimited

# Find people without an office and other potential directory problems
$NoOffice = $Directory | ? {([string]::IsNullOrEmpty($_.Office))}
ForEach ($C in $NoOffice) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Office" }               
    $Report.Add($ReportLine) 
}

$NoPhone = $Directory | ? {([string]::IsNullOrEmpty($_.Phone))}
ForEach ($C in $NoPhone) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Phone" }               
    $Report.Add($ReportLine) 
}

$NoCity = $Directory | ? {([string]::IsNullOrEmpty($_.City))}
ForEach ($C in $NoCity) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No City" }               
    $Report.Add($ReportLine) 
}

$NoCompany = $Directory | ? {([string]::IsNullOrEmpty($_.Company))}
ForEach ($C in $NoCompany) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Company" }               
    $Report.Add($ReportLine) 
}

$NoState = $Directory | ? {([string]::IsNullOrEmpty($_.StateOrProvince))}
ForEach ($C in $NoState) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No State" }               
    $Report.Add($ReportLine) 
}

$NoManager = $Directory | ? {([string]::IsNullOrEmpty($_.Manager))}
ForEach ($C in $NoManager) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Manager" }               
    $Report.Add($ReportLine) 
}

$NoZip = $Directory | ? {([string]::IsNullOrEmpty($_.PostalCode))}
ForEach ($C in $NoZip) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Postal code" }               
    $Report.Add($ReportLine) 
}

$NoTitle = $Directory | ? {([string]::IsNullOrEmpty($_.Title))}
ForEach ($C in $NoTitle) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Title" }               
    $Report.Add($ReportLine) 
}

$NoStreet = $Directory | ? {([string]::IsNullOrEmpty($_.StreetAddress))}
ForEach ($C in $NoStreet) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Street Address" }               
    $Report.Add($ReportLine) 
}

# Calculate percentages
$PercentNoOffice = ($NoOffice.Count/$Directory.Count).ToString("P")
$PercentNoCity =   ($NoCity.Count/$Directory.Count).ToString("P")
$PercentNoCompany = ($NoCompany.Count/$Directory.Count).ToString("P")
$PercentNoState = ($NoState.Count/$Directory.Count).ToString("P")
$PercentNoManager = ($NoManager.Count/$Directory.Count).ToString("P")
$PercentNoZip  = ($NoZip.Count/$Directory.Count).ToString("P")
$PercentNoTitle = ($NoTitle.Count/$Directory.Count).ToString("P")
$PercentNoPhone = ($NoPhone.Count/$Directory.Count).ToString("P")
$PercentNoStreet = ($NoStreet.Count/$Directory.Count).ToString("P")

CLS
Write-Host " "
Write-Host "Number of user mailboxes:                " $Directory.Count
Write-Host "--------------------------------------------------"
Write-Host " "
Write-Host "Analysis of potential directory problems"
Write-Host "----------------------------------------"
Write-Host "Number of mailboxes with no Office:      " $NoOffice.Count $PercentNoOffice
Write-Host "Number of mailboxes with no City         " $NoCity.Count $PercentNoCity
Write-Host "Number of mailboxes with no Company:     " $NoCompany.Count $PercentNoCompany
Write-Host "Number of mailboxes with no State:       " $NoState.Count $PercentNoState 
Write-Host "Number of mailboxes with no Manager:     " $NoManager.Count $PercentNoManager
Write-Host "Number of mailboxes with no Title:       " $NoTitle.Count $PercentNoTitle
Write-Host "Number of mailboxes with no Phone:       " $NoPhone.Count $PercentNoPhone
Write-Host "Number of mailboxes with no Address:     " $NoStreet.Count $PercentNoPhone
Write-Host "Number of mailboxes with no Postal Code: " $NoZip.Count $PercentNoZip

$Report | Sort User | Export-CSV c:\temp\DirectoryIssues.csv -NoTypeInformation
Write-Host " "
Write-Host "An output file containing details of missing directory properties is available in c:\temp\DirectoryIssues.csv"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
