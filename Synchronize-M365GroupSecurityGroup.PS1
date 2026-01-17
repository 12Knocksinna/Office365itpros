<# 
.SYNOPSIS
Synchronize-M365GroupSecurityGroup.PS1
Synchronizes the membership of a Microsoft 365 Group with that of a Security Group.

.PARAMETER M365GroupName
Display name of the Microsoft 365 Group to synchronize.

.PARAMETER SecurityGroupName
Display name of the Security Group to synchronize with the Microsoft 365 Group.

.EXAMPLE
.\Synchronize-M365GroupSecurityGroup.ps1 -M365GroupName "eDiscovery Administrators" -SecurityGroupName "eDiscoveryAdmins" 

.NOTES
Requires the Microsoft Graph PowerShell SDK and  scopes:
GroupMember.ReadWrite.All, Group.ReadWrite.All, User.ReadBasic.All

Version 1.0 17-Jan-2026
GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Synchronize-M365GroupSecurityGroup.PS1

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $M365GroupName,

    [Parameter(Mandatory)]
    [string] $SecurityGroupName

)

function Connect-GraphSession {
    $requiredScopes = @(
        "GroupMember.ReadWrite.All",
        "Group.ReadWrite.All", "User.ReadBasic.All"
    )

    if (-not (Get-MgContext)) {
        Write-Host "Connecting to Microsoft Graph and requesting delegated scopes..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        
        return
    }

    # Ensure we have all scopes
    [array]$Scopes = Get-MgContext | Select-Object -ExpandProperty Scopes
    $MissingScopes = $RequiredScopes | Where-Object { $_ -notin $Scopes }
    if ($MissingScopes) {
        Write-Host "Disconnecting from Graph to allow the addition of missing scopes: $($MissingScopes -join ', ')" -ForegroundColor Yellow
        Disconnect-MgGraph
        Break
        
    }
}

# --- Main ---

Connect-GraphSession
[int]$Removals = 0
[int]$Additions = 0

# Get the Microsoft 365 Group and check that it is a unified group with static membership
[array]$M365Group = Get-MgGroup -Filter "displayName eq '$M365GroupName'" -ConsistencyLevel eventual -CountVariable Count
If (!($M365Group)){
    Write-Output "Cannot find a unique Microsoft 365 Group named '$M365GroupName'" -ForegroundColor Red
    Break
}

If ("Unified" -notin $M365Group.GroupTypes){
    Write-Output "'$M365GroupName' is not a Microsoft 365 Group." -ForegroundColor Red
    Break
}

If ("DynamicMembership" -in $M365Group.GroupTypes){
    Write-Output "'$M365GroupName' is a dynamic Microsoft 365 Group and cannot be synchronized." -ForegroundColor Red
    Break
}

# Get the Security Group and check that it is security and mail-enabled
[array]$SecurityGroup = Get-MgGroup -Filter "displayName eq '$SecurityGroupName' and securityEnabled eq true" -ConsistencyLevel eventual -CountVariable Count
If (!($SecurityGroup)){
    Write-Output "Cannot find a unique Security Group named '$SecurityGroupName'" -ForegroundColor Red
    Break
}

If ($SecurityGroup.MailEnabled -ne $true -and $SecurityGroup.SecurityEnabled -ne $true){
    Write-Output "'$SecurityGroupName' is not a Security Group." -ForegroundColor Red
    Break
}

Write-Output "Synchronizing membership of M365 Group '$M365GroupName' with Security Group '$SecurityGroupName'"

# Get the members of both groups
[array]$M365GroupMembers = Get-MgGroupMember -GroupId $M365Group.Id -All | Select-Object -ExpandProperty Id
[array]$SecurityGroupMembers = Get-MgGroupMember -GroupId $SecurityGroup.Id -All | Select-Object -ExpandProperty Id

Write-Output "M365 Group has $($M365GroupMembers.Count) members and Security Group has $($SecurityGroupMembers.Count) members"
Write-Output "Comparing memberships..."
$Report = [System.Collections.Generic.List[Object]]::new()

# The security group is the master, so we must add security group members to the M365 group if they're not already present
[array]$NotInM365Group = $SecurityGroupMembers | Where-Object { $_ -notin $M365GroupMembers }
ForEach ($Member in $NotInM365Group)  {
    Write-Output "Adding member $Member to $M365GroupName"
    Try {
        New-MgGroupMember -GroupId $M365Group.Id -DirectoryObjectId $Member -ErrorAction Stop
        $Additions++
        $ReportLine = [PSCustomObject][Ordered]@{
            UserId                  = $Member
            Name                    = Get-MgUser -UserId $Member -Select displayName | Select-Object -ExpandProperty DisplayName
            Action                  = "Added"
            'Microsoft 365 Group'   = $M365GroupName
            Timestamp               = (Get-Date).ToString("u")
        }
        $Report.Add($ReportLine)
    } Catch {
        Write-Output "Failed to add member $Member to $M365GroupName. Error: $_" -ForegroundColor Red
    }
}   

# And remove members from the M365 group who are not in the security group
[array]$NotInSecurityGroup = $M365GroupMembers | Where-Object { $_ -notin $SecurityGroupMembers }
ForEach ($Member in $NotInSecurityGroup)  {
    Write-Output "Removing member $Member from $M365GroupName"
    Try {
        Remove-MgGroupMemberByRef -GroupId $M365Group.Id -DirectoryObjectId $Member -ErrorAction Stop
        $Removals++
        $ReportLine = [PSCustomObject][Ordered]@{
            UserId                  = $Member
            Name                    = Get-MgUser -UserId $Member -Select displayName | Select-Object -ExpandProperty DisplayName
            Action                  = "Removed"
            'Microsoft 365 Group'   = $M365GroupName
            Timestamp               = (Get-Date).ToString("u")
        }
        $Report.Add($ReportLine)
    } Catch {
        Write-Output "Failed to remove member $Member from $M365GroupName. Error: $_" -ForegroundColor Red
    }
}

Write-Host "Synchronization complete. Added $Additions members and removed $Removals members." -ForegroundColor Green
$Report

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.