# Create-EntraIDDynamicGroups.PS1
# Example script to illustrate how to create dynamic administrative units for every department in
# an organization
# https://github.com/12Knocksinna/Office365itpros/blob/master/Create-EntraIDDynamicGroups.PS1

Connect-MgGraph -NoWelcome -Scopes Directory.ReadWrite.All
Write-Host "Finding user accounts to analyze departments..."
[array]$Users = Get-MgUser -All -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
 -ConsistencyLevel eventual -CountVariable UsersFound -Property Id, UserPrincipalName, Department, DisplayName

# Get list of departments
[array]$Departments = $Users.Department | Sort-Object -Unique
# Define any exclusions we don't want to create groups for
[array]$DepartmentExclusions = "EMAIL", "Shared Mailbox"
$Departments = $Departments | Where-Object {$_ -notin $DepartmentExclusions}

# Retrieve current groups because we should check them before creating another dynamic group for 
# a department if one already exists
[array]$Groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'dynamicMembership') and groupTypes/any(x:x eq 'unified')" -All

Write-Host ("Checking dynamic Microsoft 365 groups for the following departments: {0}" -f ($Departments -Join ", "))

ForEach ($Dept in $Departments) {
    $NewGroup = $Null; $NewTeam = $Null
    Write-Host ("Checking groups for department {0}" -f $Dept)
    $Description = ("Dynamic Microsoft 365 group created for the {0} department on {1}" -f $Dept, (Get-Date))
    $DisplayName = ("{0} Dynamic group" -f $Dept)
    $MailNickName = ("Dynamic.{0}.Group" -f ($Dept -replace " ",""))
    $MembershipRule = '(User.Department -eq "' + $Dept +'")'

    If ($DisplayName -in $Groups.DisplayName) {
        Write-Host ("Group already exists for {0}" -f $Dept) -ForegroundColor Red
    } Else {
    # Create the new dynamic Microsoft 365 Group
    $NewGroup = New-MgGroup -DisplayName $DisplayName -Description $Description -MailEnabled:$True -SecurityEnabled:$False `
        -MailNickname $MailNickName -GroupTypes "DynamicMembership", "Unified" -MembershipRule $MembershipRule -MembershipRuleProcessingState "On"
    }
   
    # If the create worked, team-enable the new group
    If ($NewGroup) {
       Write-Host ("Team-enabling {0}..." -f $NewGroup.DisplayName) -ForegroundColor Yellow
       # We sleep to allow background synchronization to happen
       Start-Sleep -Seconds 15
       $GroupUri = "https://graph.microsoft.com/v1.0/groups('" + $NewGroup.Id + "')"
       $NewTeamParams = @{
          "template@odata.bind"="https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
          "group@odata.bind"="$($GroupUri)"
        }
        $NewTeam = New-MgTeam -BodyParameter $NewTeamParams
        If ($NewTeam) {
           Write-Host ("Successfully team-enabled the {0}" -f $NewGroup.DisplayName)
        }
    }
 } # End Foreach department

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
