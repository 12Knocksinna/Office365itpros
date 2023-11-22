# Reoport-ServicePrincipalAssignments.PS1
# Example script to demonstrate how to report the assignments of users and groups to Entra ID enterprise applications
# (service principals)
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ServicePrincipalAssignments.PS1
# 15-Nov-2023

Connect-MgGraph -NoWelcome -Scopes Directory.Read.All, Application.Read.All
$TenantId = (Get-MgOrganization).Id

[array]$ServicePrincipals = Get-MgServicePrincipal -All 

If ($ServicePrincipals) {
    # This filter finds enterprise apps where app role assignment is required. If you want to report
    # assignments for all enterprise apps, remove the clause that looks for $_.AppRoleAssignmentRequired to be True
    [array]$EnterpriseApps = $ServicePrincipals | Where-Object {$_.AppOwnerOrganizationId -ne $TenantId -and $_.AppRoleAssignmentRequired -eq $True} | `
        Sort-Object DisplayName
    Write-Host ("Analyzing assignments for {0} Enterprise Apps..." -f $EnterpriseApps.count)
    $Report = [System.Collections.Generic.List[Object]]::new() 
    $CSVOutputFile = "C:\temp\ServicePrincipalAssignments.CSV"
} Else {
    Write-Host "Can't find any service principals - exiting"; break
}
[int]$i = 0
ForEach ($App in $EnterpriseApps) {
    [array]$Assignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $App.Id | `
            Where-Object {$_.PrincipalType -ne 'ServicePrincipal'}
    If ($Assignments) {
        $i++
        Write-Host ("Found assignments for {0}" -f $App.DisplayName)
        ForEach ($Assignment in $Assignments) {
            $ReportLine = [PSCustomObject]@{
                TimeStamp   = $Assignment.CreatedDateTime  
                Id          = $Assignment.Id
                DisplayName = $Assignment.PrincipalDisplayName 
                UserId      = $Assignment.PrincipalId
                Type        = $Assignment.PrincipalType
                Resource    = $Assignment.ResourceDisplayName
                ResourceId  = $Assignment.ResourceId
            }
            $Report.Add($ReportLine)
        }
    }
}

$Report | Select-Object Resource, DisplayName, TimeStamp, Id, ResourceId | Out-GridView
Write-Host ("Found {0} assignments - data available in {1}" -f $i, $CSVOutputFile)
$Report | Export-CSV -NoTypeInformation $CSVOutputFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
