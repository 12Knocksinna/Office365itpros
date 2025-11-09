# Report-ComplianceRoleGroups.PS1
# An example of how to create a report about the compliance role groups in a Microsoft 365 tenant. Each role group is listed
# with the assigned roles and role group members.

# V1.0 9 November 2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ComplianceRoleGroups.PS1

Function Get-AssignmentInfo {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$RoleHolderId
    )
    
    # Function to return details about a role holder (user or service principal)
    $Id = $null; $Object = $null; $ObjectId = $null; $ObjectName = $null; $ObjectCompany = $null; $ObjectDepartment = $null; $ObjectJobTitle = $null; $UserAssignments = $null
    Try {
        $Id = Get-MgUser -UserId $RoleHolder -Property DisplayName, UserPrincipalName, companyName, department, jobtitle -ErrorAction Stop
        $Object = 'User account'
        $ObjectId = $Id.UserPrincipalName
        $ObjectName = $Id.DisplayName
        $ObjectCompany = $Id.companyName
        $ObjectDepartment = $Id.department
        $ObjectJobTitle = $Id.jobtitle
        $ObjectMemberId = $RoleHolderId
    }   Catch {
        Write-Host "Unable to retrieve user with ID $RoleHolder" -ForegroundColor Red
        $DeletedObject = $AssignmentReport | Where-Object { $_.MemberIdentity -eq $RoleHolder } | Select-Object -First 1
        $Object = "Deleted Role Holder"
        $ObjectId = $DeletedObject.Member
        $ObjectName = $RoleHolder
    }
    If ($null -eq $Id) {
        # Maybe it's a service principal
        Try {
            $Id = Get-MgServicePrincipal -ServicePrincipalId $RoleHolder -Property DisplayName, AppId -ErrorAction Stop
            $Object = 'Service principal'
            $ObjectName = $Id.displayName
            $ObjectId = $RoleHolderId
            $ObjectMemberId = $RoleHolderId
        } Catch {
            Write-Host "Unable to retrieve service principal with ID $RoleHolder" -ForegroundColor Red
            $DeletedObject = $AssignmentReport | Where-Object { $_.MemberIdentity -eq $RoleHolder } | Select-Object -First 1
            $Object = "Deleted Role Holder"
            $ObjectId = $DeletedObject.Member
            $ObjectName = $RoleHolder
            $ObjectMemberId = $RoleHolderId
        }
    } 
    If ($null -eq $Id) {
        $DeletedObject = $AssignmentReport | Where-Object { $_.MemberIdentity -eq $RoleHolder } | Select-Object -First 1
        $Object = "Unknown Object Type"
        $ObjectName = $DeletedObject.Member
        $ObjectId = $RoleHolder
    }

    $UserAssignments = $AssignmentReport | Where-Object { $_.MemberIdentity -eq $RoleHolder }

    $ReportLine = [PSCustomObject]@{
        Object          = $Object
        ObjectId        = $ObjectId
        ObjectName      = $ObjectName
        Company         = $ObjectCompany    
        Department      = $ObjectDepartment
        JobTitle        = $ObjectJobTitle
        Assignments     = $UserAssignments.Count
        MemberIdentity  = $ObjectMemberId
    }
    $Global:UserAssignmentInfo.Add($ReportLine)

    Return 
}

# Connect to the Graph (scopes needed to fetch user account, directory, and tenant information)
Connect-MgGraph -NoWelcome -Scopes User.Read.All, Organization.Read.All, Directory.Read.All
$TenantDetails = Get-MgOrganization

If ("ExchangeOnlineManagement" -notin (Get-Module | Select-Object -ExpandProperty Name)) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowBanner:$False -ErrorAction Stop
    Write-Host "Connecting to Security & Compliance Center..." -ForegroundColor Yellow
    Connect-IPPSSession -ShowBanner:$False -ErrorAction Stop
}

# Get compliance role groups
[array]$RoleGroups = Get-RoleGroup | Sort-Object DisplayName
If ($RoleGroups) {
    Write-Output "Processing $($RoleGroups.Count) compliance role groups..."
}   Else {
    Write-Host "No compliance role groups found." -ForegroundColor Yellow
    break
}

[int]$i = 0
$Report = [System.Collections.Generic.List[Object]]::new()
$AssignmentReport = [System.Collections.Generic.List[Object]]::new()

ForEach ($Group in $RoleGroups) {
    $i++
    Write-Output "Processing group $i of $($RoleGroups.Count): $($Group.DisplayName)"

    # Get the members of the role group
    [array]$Members = Get-RoleGroupMember -Identity $Group.ExchangeObjectId
    # Log each member assignment so that we can collate them on a per-member basis later
    ForEach ($Member in $Members) {
        $DirectoryObject = $null
        # Determine the type of member assigned to the role group (everything is otherwise reported as a 'MailUser')
        Try {
            $DirectoryObject = (Get-MgDirectoryObject -DirectoryObjectId $Member.ExchangeObjectId.Guid -ErrorAction Stop)
            $MemberType = $DirectoryObject.additionalProperties.'@odata.type'.split("#microsoft.graph.")[1]
        } Catch {
            $MemberType = "Object not found in directory"
        }

        # Get some more definition about the type of member holds the role. Especially interested in 
        # Managed identities vs application service principals
        Switch ($MemberType) {
            'user' { $MemberType = 'User account' }
            'servicePrincipal' { 
                If ($DirectoryObject.additionalProperties.servicePrincipalType -eq 'Application') {
                    $MemberType = 'Application service principal'
                } Else {
                    $MemberType = 'Managed identity service principal'
                }
            }
            'group' { $MemberType = 'Security group' }
            Default { $MemberType = "Object not found in directory" }
        }

        # Report what we found about the role assignment
        $ReportLine = [PSCustomObject]@{
            Group          = $Group.DisplayName
            GroupName      = $Group.Name
            Member         = $Member.Name
            MemberType     = $MemberType
            MemberIdentity = $Member.ExchangeObjectId.Guid
        }
        $AssignmentReport.Add($ReportLine)
    }

    # Get the names of the assigned roles
    [array]$Roles = $Group | Select-Object -ExpandProperty Roles
    $ExtractedRoles = @()
    ForEach ($Item in $Roles) {
        $text = $Item.ToString()
        $RoleName = $text.Split('/')[3].Trim()
        $ExtractedRoles += $RoleName
    }

    # Report the role group information
    $ReportLine = [PSCustomObject]@{
        DisplayName     = $Group.DisplayName
        Description     = $Group.Description
        GroupType       = $Group.RoleGroupType
        Members         = $Members.DisplayName -join "; "
        Roles           = $ExtractedRoles -join "; "
        'Last Changed'  = Get-Date $Group.WhenChanged -format 'dd-MMM-yyyy HH:mm'
        GroupName       = $Group.Name
    }
    $Report.Add($ReportLine)
}

# Find the set of unique users with assignments for compliance role groups
[array]$ListofIds = $AssignmentReport | Select-Object -ExpandProperty MemberIdentity -Unique

$Global:UserAssignmentInfo = [System.Collections.Generic.List[Object]]::new()
Write-Output "Checking assignments for $($ListofIds.Count) unique role holders..."
ForEach ($RoleHolder in $ListofIds) {
    # See if a user holds the role
    Get-AssignmentInfo -RoleHolderId $RoleHolder
}

# Sort the user assignment info by Object ascending, then ObjectName descending - this makes it ready for reporting
$UserAssignmentInfo = $UserAssignmentInfo |
    Sort-Object -Property @{Expression='Object'; Ascending=$false}, @{Expression='ObjectName'; Ascending=$true}

# Find role groups with assignments (all we're going to report on)
[array]$RoleGroupsWithAssigments = $AssignmentReport | Sort-Object GroupName -Unique | Select-Object GroupName

# To generate the report, we:
# Loop through the role groups with assignments and report the assignments for each role group, and then
# Loop through the user assignment info to find users with assignments and report those assignments. This way
# administrators get to see role groups and members and users with role group assignments.

# Create the HTML report
$HTMLHeader = "<html>
	<style>
	BODY{font-family: Arial; font-size: 8pt;}
	H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	TD{border: 1px solid #969595; padding: 5px }
	</style>
	<body>
        <div align=center>
        <p><h1>Microsoft Purview Role Groups Report</h1></p>
        <p><h2><b>For the " + $TenantDetails.DisplayName + " tenant</b></h2></p>
        <p><h3>Generated: " + (Get-Date -format 'dd-MMM-yyyy HH:mm') + "</h3></p></div>"

# Add the role assignment information
$HTMLBody = $HTMLHeader + "<h2>Compliance Role Groups and Assigned Members</h2><p>"
ForEach ($RoleGroupWithAssignment in $RoleGroupsWithAssigments) {
    $GroupName = $RoleGroupWithAssignment.GroupName
    $HTMLBody += "<h3>Role Group: " + $GroupName + "</h3>"
    $GroupInfo = $Report | Where-Object { $_.GroupName -eq $GroupName }
    $HTMLBody += $GroupInfo | Select-Object DisplayName, Description, GroupType, Roles, 'Last Changed' | ConvertTo-Html -Fragment
    $HTMLBody += "<p><b>Assigned Members:</b></p>"
    $GroupMembers = $AssignmentReport | Where-Object { $_.GroupName -eq $GroupName } | Select-Object Member, MemberType
    $HTMLBody += $GroupMembers | ConvertTo-Html -Fragment
    $HTMLBody += "<hr>"
}   

# Add the user assignment information
$HTMLBody += "<h2>Users and Service Principals with Compliance Role Group Assignments</h2><p>"
ForEach ($User in $UserAssignmentInfo) {
    $HTMLBody += "<h3><b>" + $User.ObjectName + "</b> (" + $User.ObjectId + ") - " + $User.Object + "</h3>"
    $HTMLBody += "<p><b>Company:</b> " + $User.Company + " <b>Department:</b> " + $User.Department + " <b>Job Title:</b> " + $User.JobTitle + "</p>"
    $HTMLBody += "<p><b>Total Role Group Assignments:</b> " + $User.Assignments + "</p>"
    $UserAssignments = $AssignmentReport | Where-Object { $_.MemberIdentity -eq $User.MemberIdentity } | Select-Object Group, MemberType
    $HTMLBody += "<p><b>Role Group Assignments:</b></p>"
    $HTMLBody += $UserAssignments | ConvertTo-Html -Fragment
    $HTMLBody += "<hr>"
}

$HTMLTail = "<p>Microsoft Purview Role Group Report <b>V1.0</b> </p>"
$HTMLTail += "</body></html>"

$HTMLReport = $HTMLBody + $HTMLTail
$HTMLOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Microsoft Purview Role Groups Report.html"
$HTMLReport | Out-File $HTMLOutputFile -Encoding UTF8

Write-Output "Report saved to $HTMLOutputFile"


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.