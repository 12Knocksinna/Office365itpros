# Create-EntraIDDynamicAUs.PS1
# Example script to illustrate how to create dynamic administrative units for every department in
# an organization
# https://github.com/12Knocksinna/Office365itpros/blob/master/Create-EntraIDDynamicAUs.PS1

Connect-MgGraph -NoWelcome -Scopes Directory.ReadWrite.All
Write-Host "Finding user accounts to analyze departments..."
[array]$Users = Get-MgUser -All -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
 -ConsistencyLevel eventual -CountVariable UsersFound -Property Id, UserPrincipalName, Department 
 [array]$Departments = $Users.Department | Sort-Object -Unique

 Write-Host ("Creating dynamic administrative units for the following departments: {0}" -f ($Departments -Join ", "))
 # Retrieve current AUs because we should check them before creating another dynamic AU for 
 # a department if one already exists
 [array]$CurrentAUs = Get-MgBetaAdministrativeUnit

 ForEach ($Department in $Departments) {
    $Description = ("Dynamic administrative unit created for the {0} department created {1}" -f $Department, (Get-Date))
    $DisplayName = ("{0} dynamic administrative unit" -f $Department)

    If ($DisplayName -in $CurrentAUs.DisplayName) {
        Write-Host ("Administrative unit already exists for {0}" -f $DisplayName)
    } Else {
    # Create the new AU
    $NewAUParameters = @{
        displayName = $DisplayName
        description = $Description
        isMemberManagementRestricted = $false
       }
       $NewAdminUnit = (New-MgBetaAdministrativeUnit -BodyParameter $NewAUParameters)
    }
    # If the create worked, update the new AU to make it a dynamic AU with a membership rule
    If ($NewAdminUnit) {
       # Define the membership rule
       $MembershipRule = '(user.department -eq "' + $Department + '" -and user.usertype -eq "member")'
       # Create hash table with the parameters
       $UpdateAUParameters = @{
	      membershipType = "Dynamic"
	      membershipRuleProcessingState = "On"
	      membershipRule = $MembershipRule
        }
        Try {
            Update-MgBetaAdministrativeUnit -AdministrativeUnitId $NewAdminUnit.Id -BodyParameter $UpdateAUParameters
        } Catch {
            Write-Host ("Error updating {0} with dynamie properties" -f $NewAdminUnit.DisplayName )
        }
        Write-Host ("Created dynamic administrative unit for the {0} department called {1}" -f $Department, $NewAdminUnit.DisplayName)
        
    }
 } # End Foreach department


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
