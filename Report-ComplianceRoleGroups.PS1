# Report-ComplianceRoleGroups.PS1
# An example of how to create a report about the compliance role groups in a Microsoft 365 tenant. Each role group is listed
# with the assigned roles and role group members.

# V1.0 June 2024
# GitHub Link:


Function Get-AssignmentInfo {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$RoleHolderId
    )
    
    $Id = $null; $Object = $null; $ObjectId = $null; $ObjectName = $null; $ObjectCompany = $null; $ObjectDepartment = $null; $ObjectJobTitle = $null; $UserAssignments = $null
    Try {
        $Id = Get-MgUser -UserId $RoleHolder -Property DisplayName, UserPrincipalName, companyName, department, jobtitle -ErrorAction Stop
        $Object = 'User account'
        $ObjectId = $Id.UserPrincipalName
        $ObjectName = $Id.DisplayName
        $ObjectCompany = $Id.companyName
        $ObjectDepartment = $Id.department
        $ObjectJobTitle = $Id.jobtitle
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
        } Catch {
            Write-Host "Unable to retrieve service principal with ID $RoleHolder" -ForegroundColor Red
            $DeletedObject = $AssignmentReport | Where-Object { $_.MemberIdentity -eq $RoleHolder } | Select-Object -First 1
            $Object = "Deleted Role Holder"
            $ObjectId = $DeletedObject.Member
            $ObjectName = $RoleHolder
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
        }
    $Global:UserAssignmentInfo.Add($ReportLine)

    Return 
}

# Connect to the Graph (needed to fetch user account information)
Connect-MgGraph -NoWelcome -Scopes User.Read.All

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
        $ReportLine = [PSCustomObject]@{
            Group          = $Group.DisplayName
            GroupName      = $Group.Name
            Member         = $Member.Name
            MemberType     = $Member.RecipientType
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
ForEach ($RoleHolder in $ListofIds) {
    # See if a user holds the role
    Get-AssignmentInfo -RoleHolderId $RoleHolder
}

# Sort the user assignment info by Object ascending, then ObjectName descending - this makes it ready for reporting
$UserAssignmentInfo = $UserAssignmentInfo |
    Sort-Object -Property @{Expression='Object'; Ascending=$false}, @{Expression='ObjectName'; Ascending=$true}

# Find role groups with assignments (all we're going to report on)
[array]$RoleGroupsWithAssigments = $AssignmentReport | Sort-Object GroupName -Unique | Select-Object GroupName

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.