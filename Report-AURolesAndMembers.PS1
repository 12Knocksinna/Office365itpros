# Report-AURolesAndMembers.PS1
# Show how to use the Microsoft Graph PowerShell SDK (V2 or later) to report administrative units, users with role assignments to manage the AUs
# and the membership of the AUs
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-AURolesAndMembers.PS1
# Updated 14-Nov-2023 to remove some beta cmdlets and replace them with V1.0 cmdlets (based on SDK V2.9)

Connect-MgGraph -Scopes Directory.Read.All -NoWelcome

$OutputFile = "c:\temp\AdminUnitAssignments.csv"  
$Version = "V1.1"
$HtmlReportFile = "c:\temp\AdminUnitAssignments.html"  

Write-Host "Looking for administrative units and role assignments..."
[array]$AdminUnits = Get-MgDirectoryAdministrativeUnit | Sort-Object DisplayName
[array]$Assignments = Get-MgRoleManagementDirectoryRoleAssignment
Write-Host ("Now examining details of {0} administrative units and {1} role assignments..." -f $AdminUnits.count, $Assignments.count)

[int]$i=0
$Report = [System.Collections.Generic.List[Object]]::new()  
ForEach ($AU in $AdminUnits) {
   $i++
   Write-Output ("Processing admin unit: {0} ({1}/{2})..." -f $AU.DisplayName, $i, $AdminUnits.count)
   $AUAssignments = $Null; $AUAssignments = $Null

   $AUId = "/administrativeUnits/" + $AU.Id
   If ($AU.IsMemberManagementRestricted -eq $True) {
      $Scope = "Restricted"    
   } Else {
      $Scope = "Directory" 
   }

   $AUAssignments = $Assignments | Where-Object {$_.directoryscopeid -eq $AuId}
   $RoleData = $Null;  [array]$Roles = $Null
   ForEach ($AUAssignment in $AUAssignments) {
      # Check if assignment is for a user account
      $Result = Get-MgUser -UserId $AUAssignment.PrincipalId -ErrorAction SilentlyContinue
      $Type = "User"
      # Check if it's a group
      If (!($Result)) {
         $Result = Get-MgGroup -GroupId $AUAssignment.PrincipalId -ErrorAction SilentlyContinue
         $Type = "Group"
      }
      If (!($Result)) { # No user, so try service principal
         $Result = Get-MgServicePrincipal -ServicePrincipalId $AuAssignment.PrincipalId -ErrorAction SilentlyContinue 
         $Type = "Service Principal" 
      }
      $Role =  Get-MgDirectoryRoleTemplate -DirectoryRoleTemplateId $AuAssignment.RoleDefinitionId 
      If ($Result -and $Type -ne "Group") {
        $RoleData = ("{0}/{1} ({2})" -f $Result.DisplayName, $Type, $Role.DisplayName)
        $Roles += $RoleData 
      } ElseIf ($Result -and $Type -eq "Group") {
        $AuGroupMembers = Get-MgGroupMember -GroupId $AuAssignment.PrincipalId -All
        ForEach ($Member in $AuGroupMembers.additionalProperties.displayName) {
            $RoleData = ("{0}/{1} ({2})" -f $Member, $Type, $Role.DisplayName)
            $Roles += $RoleData }
      }
   } 
   # Check if the membership is dynamic or static
   If ($AU.additionalProperties.membershipType) {
      $AuMembershipType = "Dynamic" 
      $AuMembershipRule = $AU.additionalProperties.membershipRule
   } Else {
      $AuMembershipType = "Static" 
      $AuMembershipRule = $Null
   }

   # Get membership
   [array]$AuMembers = Get-MgBetaAdministrativeUnitMember -AdministrativeUnitId $Au.Id -All
   [array]$MemberUsers = $AuMembers.additionalProperties | Where-Object {$_.'@odata.type' -eq "#microsoft.graph.user"}
   [array]$MemberGroups = $AuMembers.additionalProperties | Where-Object {$_.'@odata.type' -eq "#microsoft.graph.group"}
   $AuMembersNames = $AuMembers.additionalproperties.displayName -Join ", "

   If ($Roles) {
      $RoleDisplayNames = $Roles -join ", "
   } Else {
      $RoleDisplayNames = "Administrator Roles"
   }

   $ReportLine = [PSCustomObject]@{
        DisplayName     = $AU.DisplayName
        Description     = $AU.Description
        Id              = $AU.Id
        Scope           = $Scope
        Assignments     = $RoleDisplayNames
        Members         = $AuMembersNames
        'User members'  = $MemberUsers.count
        'Group members' = $MemberGroups.count
        MembershipType  = $AuMembershipType
        MembershipRule  = $AuMembershipRule
        }
   $Report.Add($ReportLine)

} #End ForEach AU

[string]$Organization = (Get-MgOrganization).DisplayName

# Generate the report files
$HtmlHeading ="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.pass{background: #B7EB83;}
	   td.warn{background: #FFF275;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Microsoft Entra ID Administrative Units Report</h1></p>
           <p><h2>for $Organization</h2></p>
           <p><h3>Generated: " + (Get-Date -format 'dd-MMM-yyyy hh:mm tt') + "</h3></p></div>"

$HtmlReport = $HtmlHeading
$HtmlData = $Report | ConvertTo-html -Fragment

$Htmltail = "<p><p>Report created for: " + ($Organization.DisplayName) + "</p><p>" +
             "<p>Number of Entra ID Administrative Units:     " + $AdminUnits.count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------" +
             "<p>Microsoft Entra ID Administrative Units<b>" + $Version + "</b>"	
$HtmlReport = $HtmlHeading + $HtmlData + $HtmlTail
$HtmlReport | Out-File $HtmlReportFile  -Encoding UTF8

$Report | Export-CSV -NoTypeInformation $OutputFile
Write-Host ("All done. Reports available in {0} and {1}" -f $OutputFile, $HtmlReportFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 