# Update-Managers.PS1
# A script to explore how to keep manager-employee assignments updated.
# https://github.com/12Knocksinna/Office365itpros/blob/master/Update-Managers.PS1

Connect-MgGraph -Scopes User.ReadWrite.All, Directory.Read.All
Write-Host "Setting up manager data..."
$InputManagerData = "c:\temp\ManagerData.csv"
[int]$ErrorCount = 0
$ManagerHash = @{}

# Parse the manager data
[array]$ManagerData = Import-Csv -Path $InputManagerData

ForEach ($M in $ManagerData) {
   $ManagerName = $M.Manager
   $Manager = Get-MgUser -UserId $M.Manager -ErrorAction SilentlyContinue
   If ($null -eq $Manager) {
      Write-Host ("Can't find a user account for {0} for the {1} department" -f $ManagerName, $M.Department)
      $ErrorCount++ 
   } Else {
      $ManagerHash.Add([string]$M.Department,[string]$M.Manager)
   }
}

If ($ErrorCount -gt 0) { Write-Host "Please fix the manager input file before proceeeding..." -foregroundcolor Red; break }

# Find user accounts with assigned licenses
Write-Host "Finding licensed Azure AD accounts..."
[array]$Users = Get-MgUser -All -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable UsersFound -Property Manager | Sort-Object DisplayName
If (!($Users)) { Write-Host "Couldn't find any user accounts with assigned licenses!" ; break }

Write-Host "Checking managers assigned to user accounts..."
[array]$UsersNoDepartment = $Users | Where-Object {$Null -eq $_.department}
[array]$UsersWithDepartment = $Users | Where-Object {$Null -ne $_.department}

# Get default manager to assign people to when we can't find a department manager
$DefaultManager = $ManagerData | Where-Object {$_.Department -eq "Default" } | Select-Object -ExpandProperty Manager

$Report = [System.Collections.Generic.List[Object]]::new() 

ForEach ($User in $UsersWithDepartment) {
  $DefaultFlag = $Null; $DeptManager = $Null
  $DepartmentManager = $ManagerHash[$User.Department]
  # Check if we can identify a department manager for this user account
  If (($DepartmentManager) -and ($DepartmentManager -ne $User.UserPrincipalName)) {
       Write-Host ("User account {0} has manager {1}" -f $User.displayName, $DepartmentManager)
  } ElseIf (($DepartmentManager) -and ($DepartmentManager -eq $User.UserPrincipalName)) {
      # The user is a department manager, so we update with the level 2 manager
      $DeptManager = $True
      $DepartmentManager = $ManagerData | Where-Object {$_.Department -eq $User.Department} | Select-Object -ExpandProperty Level2Manager
      If ($DepartmentManager) {
         Write-Host ("User account {0} is a manager, so we use level 2 manager {1}" -f $User.displayName, $DepartmentManager)
      }
  }
  
  # Use the default manager if we can't find a department manager for some reason, but don't update the CEO
  If (!($DepartmentManager) -and ($User.JobTitle -ne "Chief Executive Officer")) { 
     $DefaultFlag = "Yes"
     $DepartmentManager = $DefaultManager }

  # Update the account
 If ($DepartmentManager) {
   Write-Host ("Updating user account {0} with manager {1}" -f $User.displayName, $DepartmentManager)
   $Id = ("https://graph.microsoft.com/v1.0/users/{0}" -f $DepartmentManager) 
   Set-MgUserManagerByRef -UserId $User.Id -AdditionalProperties @{ "@odata.id" = $Id }
   $ReportLine = [PSCustomObject]@{
     Account                    = $User.UserPrincipalName
     User                       = $User.DisplayName
     Department                 = $User.Department
     Manager                    = $DepartmentManager
    "Department Manager"       = $DeptManager
    "Assigned Default Manager" = $DefaultFlag }
   $Report.Add($ReportLine)
  }

} # End ForEach User


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
