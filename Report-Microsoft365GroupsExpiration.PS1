# Report-Microsoft365GroupsExpiration.PS1
# A script to show how to use the Microsoft Graph PowerShell SDK to report on the expiration of Microsoft 365 groups.

# V1.0 21-Nov-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-Microsoft365GroupsExpiration.PS1

# Connect to the Microsoft Graph - the signed-in account must have the necessary role to read group information
Connect-MgGraph -Scopes Group.Read.All

Write-Host "Looking for Microsoft 365 Groups with an expiration date set"
[array]$Groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'unified') and ExpirationDateTime ge 2014-01-01T00:00:00Z" `
    -All -PageSize 500 -ConsistencyLevel Eventual -CountVariable Var -Property DisplayName, CreatedDateTime, RenewedDateTime, ExpirationDateTime
    
If (!($Groups)) {   
    Write-Host "No groups with expiration dates found"
    Break
}   Else {
    Write-Host ("Found {0} groups with expiration dates" -f $Var)
}
$Today = (Get-Date)
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($G in $Groups) {
    $Days = (New-TimeSpan -Start $G.CreatedDateTime -End $Today).Days  # Age of group
    $DaysLeft = (New-TimeSpan -Start $Today -End $G.ExpirationDateTime).Days
    $ReportLine = [PSCustomObject]@{
       Group                   = $G.DisplayName
       Created                 = Get-Date($G.CreatedDateTime) -format 'dd-MMM-yyyy HH:mm'
      "Age in days"            = $Days
      "Last renewed"           = Get-Date($G.RenewedDateTime) -format 'dd-MMM-yyyy'
      "Next renewal"           = Get-Date($G.ExpirationDateTime) -format 'dd-MMM-yyyy'
      "Days before expiration" = $DaysLeft}
      $Report.Add($ReportLine)
} # End Foreach

$Report = $Report | Sort-Object "Days before expiration"
Clear-Host
Write-Host ("A total of {0} Microsoft 365 Groups are covered by expiration policy:" -f $Groups.Count)
Write-Host ""
$Report | Format-Table Group, 'Age in Days', Created, 'Last renewed', 'Next renewal', 'Days before expiration' -AutoSize
$Report | Export-CSV -Path "C:\Temp\GroupExpirationReport.csv" -NoTypeInformation -Encoding UTF8

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.