# FindNonSSPREnabledUsers.PS1
# An example of how to find Azure AD user accounts that aren't SSPR capable.

# Connect to the Graph SDK
Connect-MgGraph -Scope Directory.Read.All, UserAuthenticationMethod.Read.All, AuditLog.Read.All 
Select-MgProfile Beta

Write-Host "Finding licensed Azure AD accounts"
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All
# Populate a hash table with the details about user accounts
$UserTable = @{}
ForEach ($U in $Users) {
    $ReportLine  = [PSCustomObject] @{          
     Id                  = $U.Id
     DisplayName         = $U.DisplayName
     Department          = $U.Department
     Office              = $U.OfficeLocation  
     Country             = $U.Country
     }
    $UserTable.Add([String]$U.Id, $ReportLine) 
}

Write-Host "Finding user accounts not capable of Self-Service Password Reset (SSPR)"
[array]$SSPRUsers = Get-MgReportAuthenticationMethodUserRegistrationDetail | Where-Object {$_.userType -eq 'member' -and $_.IsSSPRCapable -eq $False} | Select-Object Id, userDisplayName, userPrincipalName, DefaultMfaMethod, IsAdmin, IsMfaCapable, IsMfaRegistered, IsPasswordlessCapable, IsSSPRCapable                                                                                  
Write-Host "Cross-checking against licensed users..."
$NonSSPRUsers = [System.Collections.Generic.List[Object]]::new() # Create merged output file
ForEach ($S in $SSPRUsers) {
  $Data = $UserTable.Item($S.Id) 
  If ($Data) { # We found a match
     $ReportLine  = [PSCustomObject] @{  
       Id = $Data.Id
       DisplayName = $Data.DisplayName
       Department  = $Data.Department
       Office      = $Data.Office
       Country     = $Data.Country }
     $NonSSPRUsers.Add($ReportLine) }
}
$PNonSSPR = ($NonSSPRUsers.count/$Users.Count).toString("P")
Write-Host ("{0} out of {1} licensed accounts ({2}) are not enabled for Self-Service Password Reset" -f $NonSSPRUsers.count, $Users.count, $PNonSSPR )
Write-Host ($NonSSPRUsers.DisplayName -join ", ")
$NonSSPRUsers | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
