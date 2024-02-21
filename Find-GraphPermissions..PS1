# Find-GraphPermissions.PS1
# Return the Graph permissions needed to run a set of cmdlets (normally in a script)
[array]$InputCmdlets = "Get-MgUser", "Get-MgGroup", "New-MgPlannerTask"
[array]$RequiredPermissions = $null

ForEach ($Command in $InputCmdlets) {
    $Permissions = (Find-MgGraphCommand -Command $Command | Select-Object -ExpandProperty Permissions).Name
    ForEach ($Permission in $Permissions) {
        If ($Permission -notin $RequiredPermissions) {
            $RequiredPermissions += $Permission
        }
    }
}
Write-Host ("To use these {0} cmdlets, you need the following permissions: {1}" -f ($InputCmdlets -join ", "), ($RequiredPermissions -join ", "))

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.