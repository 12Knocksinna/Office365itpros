# FindPotentialDirectoryProblems.PS1
#
# Quick and dirty script to highlight potential issues which might exist in an Office 365 tenant directory
# 
# Get list of user mailboxes in the directory

If ($Null -eq (Get-ConnectionInformation)) {
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
Write-Host "Finding User mailboxes..."
[array]$Directory = Get-User -RecipientTypeDetails UserMailbox -ResultSize Unlimited
If (!($Directory)) { 
    Write-Host "Unable to find user accounts - exiting" ; break 
}

# Find people without an office and other potential directory problems
[array]$NoOffice = $Directory | Where-Object {([string]::IsNullOrEmpty($_.Office))}
ForEach ($C in $NoOffice) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Office" }               
    $Report.Add($ReportLine) 
}

[array]$NoPhone = $Directory | Where-Object {([string]::IsNullOrEmpty($_.Phone))}
ForEach ($C in $NoPhone) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Phone" }               
    $Report.Add($ReportLine) 
}

[array]$NoCity = $Directory | Where-Object {([string]::IsNullOrEmpty($_.City))}
ForEach ($C in $NoCity) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No City" }               
    $Report.Add($ReportLine) 
}

$NoCompany = $Directory | Where-Object {([string]::IsNullOrEmpty($_.Company))}
ForEach ($C in $NoCompany) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Company" }               
    $Report.Add($ReportLine) 
}

[array]$NoState = $Directory | Where-Obbject {([string]::IsNullOrEmpty($_.StateOrProvince))}
ForEach ($C in $NoState) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No State" }               
    $Report.Add($ReportLine) 
}

[array]$NoManager = $Directory | Where-Object {([string]::IsNullOrEmpty($_.Manager))}
ForEach ($C in $NoManager) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Manager" }               
    $Report.Add($ReportLine) 
}

$NoZip = $Directory | Where-Object {([string]::IsNullOrEmpty($_.PostalCode))}
ForEach ($C in $NoZip) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Postal code" }               
    $Report.Add($ReportLine) 
}

[array]$NoTitle = $Directory | Where-Object {([string]::IsNullOrEmpty($_.Title))}
ForEach ($C in $NoTitle) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Title" }               
    $Report.Add($ReportLine) 
}

[array]$NoStreet = $Directory | Where-Object {([string]::IsNullOrEmpty($_.StreetAddress))}
ForEach ($C in $NoStreet) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Street Address" }               
    $Report.Add($ReportLine) 
}

# No department
[array]$NoDepartment = $Directory | Where-Object {([string]::IsNullOrEmpty($_.Department))}
ForEach ($C in $NoDepartment) {
   $ReportLine = [PSCustomObject] @{
       User        = $C.UserPrincipalName
       Name        = $C.DisplayName
       Issue       = "No Department" }               
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
$PercentNoDepartment = ($NoDepartment.count/$Directory.count).toString("P")

Clear-Host
Write-Host " "
Write-Host ("Number of user mailboxes {0}" -f $Directory.Count)
Write-Host   "---------------------------"
Write-Host " "
Write-Host "Analysis of missing directory properties"
Write-Host "----------------------------------------"
Write-Host ("Number of mailboxes with no Office     {0} ({1})" -f $NoOffice.Count, $PercentNoOffice)
Write-Host ("Number of mailboxes with no City       {0} ({1})" -f $NoCity.Count, $PercentNoCity)
Write-Host ("Number of mailboxes with no Company    {0} ({1})" -f $NoCompany.Count, $PercentNoCompany)
Write-Host ("Number of mailboxes with no Department {0} ({1})" -f $NoDepartment.count, $PercentNoDepartment)
Write-Host ("Number of mailboxes with no State      {0} ({1})" -f $NoState.Count, $PercentNoState)
Write-Host ("Number of mailboxes with no Manager    {0} ({1})" -f $NoManager.Count, $PercentNoManager)
Write-Host ("Number of mailboxes with no Title      {0} ({1})" -f $NoTitle.Count, $PercentNoTitle)
Write-Host ("Number of mailboxes with no Phone      {0} ({1})" -f $NoPhone.Count, $PercentNoPhone)
Write-Host ("Number of mailboxes with no Address    {0} ({1})" -f $NoStreet.Count, $PercentNoStreet)
Write-Host ("Number of mailboxes with no Post Code  {0} ({1})" -f $NoZip.Count, $PercentNoZip)

$Report | Sort-Object User | Export-CSV c:\temp\DirectoryIssues.csv -NoTypeInformation
Write-Host " "
Write-Host "An output file containing details of missing directory properties is available in c:\temp\DirectoryIssues.csv"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
