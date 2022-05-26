# ReportLastAccountSignIn-Mg.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportLastAccountSignIn-Mg.PS1
# Demonstration of how to fetch user account sign in data using the Microsoft Graph SDK for PowerShell
#
# Connect to the Graph, specifing the tenant and profile to use - Add your tenant identifier here
Connect-MgGraph -TenantId xxxxx-axxad-acadd-ddaadda -Scope Auditlog.Read.All
Select-MgProfile beta # Beta needed to get more informtion from the Graph
$Details = Get-MgContext
$Scopes = $Details | Select -ExpandProperty Scopes
$Scopes = $Scopes -Join ", "
$ProfileName = (Get-MgProfile).Name
$OrgName = (Get-MgOrganization).DisplayName
CLS
Write-Host "Microsoft Graph Connection Information"
Write-Host "--------------------------------------"
Write-Host " "
Write-Host ("Connected to Tenant {0} ({1}) as account {2}" -f $Details.TenantId, $OrgName, $Details.Account)
Write-Host "+-------------------------------------------------------------------------------------------------------------------+"
Write-Host ("Profile set as {0}. The following permission scope is defined: {1}" -f $ProfileName, $Scopes)
Write-Host ""

Write-Host "Finding Azure AD Users"
[array]$Users = Get-MgUser -All
If ($Users.Count -eq 0) { Write-Host "Sorry, no Azure AD user info fetched"; break}
CLS
$ProgressDelta = 100/($Users.count); $PercentComplete = 0; $UserNumber = 0
Write-Host ("{0} accounts to process..." -f $Users.Count)
CLS;$Report = [System.Collections.Generic.List[Object]]::new();$CSVOutput = "C:\temp\LastAccountSignIn.CSV"
ForEach ($User in $Users) { # See if we can find the last sign in record for an account
   $UserNumber++
   $PercentComplete += $ProgressDelta
   $UserStatus = $User.DisplayName + " ["+ $UserNumber +"/" + $Users.Count + "]"
   Write-Progress -Activity "Processing user" -Status $UserStatus -PercentComplete $PercentComplete
   [array]$LastSignIn = Get-MgAuditLogSignIn -Filter "UserId eq '$($User.Id)'" -Top 1
   If ($LastSignIn) { # We found a sign in record, so report it
      $ReportLine  = [PSCustomObject][Ordered]@{ 
         Name           = $LastSignIn.UserDisplayName
         UPN            = $LastSignIn.UserPrincipalName
         Id             = $LastSignIn.UserId
         UserType       = $User.UserType
         AccountCreated = $User.CreatedDateTime
         SignInDate     = $LastSignIn.CreatedDateTime
         DaysSince      = ($LastSignIn.CreatedDateTime | New-TimeSpan).Days
         Location       = $LastSignIn.Location.City
         App            = $LastSignIn.AppDisplayName
         Client         = $LastSignIn.ClientAppUsed
         IPAddress      = $LastSignIn.IPAddress
         IsInteractive  = $LastSignIn.IsInterActive
         CAAccess       = $LastSignIn.ConditionalAccessStatus 
         RiskState      = $LastSignIn.RiskState }
      $Report.Add($ReportLine) } #End if
    Else { # No sign in record found, so we report that"
       $ReportLine  = [PSCustomObject][Ordered]@{ 
         Name           = $User.UserDisplayName
         UPN            = $User.UserPrincipalName
         Id             = $User.Id
         UserType       = $User.UserType
         AccountCreated = $User.CreatedDateTime
         SignInDate     = "No sign in data found"
         DaysSince      = "N/A"
         App            = "N/A"
         Client         = "N/A"
         IPAddress      = "N/A"
         IsInteractive  = "N/A"
         CAAccess       = "N/A" 
         RiskState      = "N/A" }
      $Report.Add($ReportLine) } #End Else
} #End ForEach

$Report | Sort {$_.SignInDate -as [datetime]} -descending | Out-GridView
$Report | Export-CSV -NoTypeInformation $CSVOutput
Write-Host "All done. Report is available in" $CSVOutput

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
