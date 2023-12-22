# Report-PermissionConsentRequests.PS1
# Script showing how to report the permission consent requests executed by users
# original code from ghdineshs@gmail.com
# V1.0 
# https://github.com/12Knocksinna/Office365itpros/edit/master/Report-PermissionConsentRequests.PS1

#Connect to Graph - I believe Directory.Read.All also suffices
Connect-MgGraph -NoWelcome -Scopes ConsentRequest.Read.All

[array]$HighPriorityPermissions = "User.Read.All", "User.ReadWrite.All", "Group.Read.All", "Group.ReadWrite.All", `
    "Directory.ReadWrite.All", "Sites.Manage.All", "Policy.Read.All", "Mail.Read", `
    "Application.Read.All","AuditLog.Read.All", "Mail.Send", "Organization.Read.All", "Sites.Read.All"
$Report = [System.Collections.Generic.List[Object]]::new()
$CSVFile = "C:\temp\AdministratorConsentRequests.CSV"

[array]$ConsentRequests = Get-MgIdentityGovernanceAppConsentRequest -All
If (!($ConsentRequests)) {
    Write-Host "No administrator consent requests for application permissions found"; break
}

Write-Host ("Processing {0} administrator consent requests..." -f $ConsentRequests.Count)
ForEach ($ConsentRequest in $ConsentRequests) {
    $RequestedUsers = $Null
    [array]$RequestedUsers = Get-MgIdentityGovernanceAppConsentRequestUserConsentRequest -AppConsentRequestId $ConsentRequest.Id
    ForEach ($RequestedUser in $RequestedUsers) {
        # Get details of requesting user
        $User = Get-MgUser -UserId $RequestedUser.CreatedBy.User.Id `
            -Property Id, displayName, department, country, userprincipalname, jobtitle
        # Get requested scopes (permissions)    
        [array]$RequestedScopes = $ConsentRequest.PendingScopes.DisplayName
        [array]$PotentialScopeProblems = $Null
        # Check if any of the requested scopes are high-priority and need further administrator review
        ForEach ($Scope in $RequestedScopes) {
            If ($Scope -in $HighPriorityPermissions) {
                $PotentialScopeProblems += $Scope
            }
        }
        # Generate output
        $Reportline = [PSCustomObject]@{
            AppDisplayName          = $ConsentRequest.AppDisplayName
            AppID                   = $ConsentRequest.AppId
            ConsentID               = $ConsentRequest.Id
            PendingScopes           = $RequestedScopes -join ", "
            'High priority scopes'  = ($PotentialScopeProblems -join ", ")
            RequestStatus           = $RequestedUser.Status
            CreatedDateTime         = $RequestedUser.CreatedDateTime
            'Requesting user'       = $RequestedUser.CreatedBy.User.DisplayName
            'Job title'             = $User.JobTitle
            'User Id'               = $RequestedUser.CreatedBy.User.Id
            Department              = $User.Department
            Country                 = $User.Country
        }
        $Report.Add($Reportline)
    }
}
$Report | Export-Csv $CSVFile -NoTypeInformation -Encoding UTF8
$Report | Out-GridView
Write-Host ("All done - report is available in {0}" -f $CSVFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
