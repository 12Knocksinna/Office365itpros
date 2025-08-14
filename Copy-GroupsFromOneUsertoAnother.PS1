# Copy-GroupsFromOneUsertoAnother.ps1
# Description: This script copies all groups from one user to another in Microsoft 365. 
# .\Copy-GroupsFromOneUserToAnother -Source James.Ryan@contoso.com -Target Linda.Smith@contoso.com -RemoveOriginalUser:$true

# V1.0 1-Mar-2025
# V2.0 7-July-2025. Added support for removing the original user from the groups after copying them to the target user.
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Copy-GroupsFromOneUsertoAnother.PS1
# See article https://office365itpros.com/2025/07/08/copy-group-membership-powershell/

param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the UPN for the source user for group memberships to copy from")]
    [string]$Source,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the UPN for the user to copy group memberships to")]
    [string]$Target,
    [Parameter(Mandatory = $false, HelpMessage = "Remove the original user from the groups after copying them to the target user (true/false) (default: false)")]
    [bool]$RemoveOriginalUser = $false
)

function Remove-UserFromM365Group {
    # Remove a user account from a Microsoft 365 group, checking if the user is an owner first. If the user is an owner and they are not the last owner,
    # the function removes the user from as both an owner and a member of the group.
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        [Parameter(Mandatory = $true)]
        [string]$GroupId,
         [Parameter(Mandatory = $true)]
        [string]$GroupName
    )
    
    $Outcome = $null
    [array]$Owners = Get-MgGroupOwner -GroupId $GroupId | Select-Object -ExpandProperty Id
    if ($UserId -in $Owners) {
        if ($Owners.count -eq 1) {
            Write-Host ("The {0} group has only one owner - can't remove the last owner" -f $GroupDisplayName) -Foregroundcolor Red
            $Outcome = "Failed - can't remove the last owner of a group"
            return $Outcome
        }
        Write-Host ("User {0} is an owner of group {1} - removing both owner and member links" -f $TargetDisplayName, $GroupDisplayName) -ForegroundColor Yellow
        try {
            Remove-MgGroupOwnerByRef -DirectoryObjectId $UserId -GroupId $GroupId
        } catch {
            Write-Host ("Failed to remove user {0} as owner from group {1}: {2}" -f $TargetDisplayName, $GroupDisplayName, $_.Exception.Message) -ForegroundColor Red
            $Outcome = "Failed to remove owner"
            return $Outcome
        }
    }
    try {
        Remove-MgGroupMemberByRef -DirectoryObjectId $UserId -GroupId $GroupId
    } catch {
        Write-Host ("Failed to remove user {0} as member from group {1}: {2}" -f $TargetDisplayName, $GroupDisplayName, $_.Exception.Message) -ForegroundColor Red
        $Outcome = "Failed to remove member"
        return $Outcome
    }

    $Outcome = ("User {0} removed from group {1}" -f $SourceDisplayName, $GroupDIsplayName)
    return $Outcome
}

function Remove-UserFromDistributionList {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        [Parameter(Mandatory = $true)]
        [string]$GroupId,
         [Parameter(Mandatory = $true)]
        [string]$GroupName
    )
    try {
        Remove-DistributionGroupMember -Identity $GroupId -Member $UserId -Confirm:$false -ErrorAction Stop
        Write-Host ("User {0} removed from distribution list {1}" -f $TargetDisplayName, $GroupDisplayName) -ForegroundColor Yellow
        return
    } catch {
        Write-Host ("Failed to remove mailbox {0} from distribution list {1}: {2}" -f $TargetDisplayName, $GroupDisplayName, $_.Exception.Message) -ForegroundColor Red
        return $_.Exception.Message
    }
}

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "GroupMember.ReadWrite.All" -NoWelcome
# Connect to Exchange Online if necessary
[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ($Modules -notcontains "ExchangeOnlineManagement") {
    Write-Host "Connecting to Exchange Online" -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
}

Write-Host "Making sure that the source and target users exist" -ForegroundColor Cyan
# Check if the source user exists  
Try {
    $SourceUser = Get-MgUser -UserId $Source -ErrorAction Stop -Property DisplayName, UserPrincipalName, Id
    $Global:SourceDisplayName = $SourceUser.DisplayName
} Catch {
    Write-Host "Source user not found. Please check the User Principal Name." -ForegroundColor Red
    Break
}
# Check if the target user exists
Try {      
    $TargetUser = Get-MgUser -UserId $Target -ErrorAction Stop -Property DisplayName, UserPrincipalName, Id
    $Global:TargetDisplayName = $TargetUser.DisplayName
} catch {
    Write-Host "Target user not found. Please check the User Principal Name." -ForegroundColor Red
    Break
}

Write-Host ("Checking groups for user {0} to copy to {1}" -f $SourceUser.DisplayName, $TargetUser.DisplayName)
Write-Host "Note: Groups with dynamic membership are not copied." -ForegroundColor Yellow
[array]$SourceGroups = Get-MgUserMemberOf -UserId $SourceUser.Id -All -PageSize 500 | `
    Where-Object {
        ($_.additionalProperties.'@odata.type' -eq '#microsoft.graph.group') -and
        (
            -not ($_.additionalProperties.groupTypes -contains "DynamicMembership")
        )
    } | Select-Object -ExpandProperty Id
If ($null -eq $SourceGroups) {
    Write-Host "No groups found for user $($SourceUser.DisplayName)." -ForegroundColor Yellow
    Break
}

# Check what groups the target user is already a member of
[array]$CurrentTargetGroups = Get-MgUserMemberOf -UserId $TargetUser.Id -All -PageSize 500 | `
    Where-Object {$_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group'} | `
    Select-Object -ExpandProperty Id

$GroupsToProcess = [System.Collections.Generic.List[Object]]::new()
ForEach ($GroupId in $SourceGroups) {
    If ($GroupId -notin $CurrentTargetGroups) {
        $GroupsToProcess.Add($GroupId)
    }
}    

# If there are no groups to copy fromt the source user to the target user, exit the script
If ($GroupsToProcess.Count -eq 0) {
    Write-Host "No new groups to copy from $($SourceUser.DisplayName) to $($TargetUser.DisplayName)." -ForegroundColor Yellow
    Break
}

Write-Host ("Found {0} groups to copy from {1} to {2}" -f $GroupsToProcess.Count, $SourceUser.DisplayName, $TargetUser.DisplayName)
Write-Host "Starting to copy groups..." -ForegroundColor Yellow
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($GroupId in $GroupsToProcess) {
    Try {
        $Group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop -Property DisplayName, Id, GroupTypes, MailEnabled, SecurityEnabled
        $GroupDisplayName = $Group.DisplayName
    } Catch {
        Write-Host ("Failed to retrieve group {0}: {1}" -f $GroupId, $_.Exception.Message) -ForegroundColor Red
        $ReportLine = [PSCustomObject]@{
            GroupId = $GroupId
            Status  = "Failed to fetch group details"
            Error   = $_.Exception.Message
        }
        $Report.Add($ReportLine)
        Continue
    }
    # If you don't wnat to process Viva Engage communities, uncomment these lines
    #If ($Group.additionalProperties.creationOptions -contains 'YammerProvisioning') {
    #    Write-Host ("Skipping Viva Engage group {0}" -f $Group.DisplayName) -ForegroundColor Yellow
    #    $ReportLine = [PSCustomObject]@{
    #        GroupId = $GroupId
    #        Name    = $Group.DisplayName
    #        Type    = "Viva Engage Community"
    #        Status  = "Skipped Viva Engage community"
    #        Error   = $null
    #    }
    #    $Report.Add($ReportLine)
    #    Continue
    #}
    If ($Group.groupTypes -contains "Unified") { # Microsoft 365 group 
        Write-Host ("Adding {0} to Microsoft 365 group {1}" -f $TargetUser.displayName, $GroupDisplayName) -ForegroundColor Cyan
        Try {
            New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $TargetUser.Id -ErrorAction Stop
            $ReportLine = [PSCustomObject]@{
                GroupId = $Group.Id
                Name    = $Group.DisplayName
                Type    = "Microsoft 365 Group"
                Status  = "Success"
                Error   = $null
            }
            $Report.Add($ReportLine)
            If ($RemoveOriginalUser) {
                $Outcome = Remove-UserFromM365Group -UserId $SourceUser.Id -GroupId $Group.Id -GroupName $GroupDisplayName
                Write-Host $Outcome -ForegroundColor Yellow
            }
        } Catch {
            Write-Host ("Failed to add {0} to Microsoft 365 group {1}: {2}" -f $TargetUser.DisplayName, $GroupDisplayName, $_.Exception.Message) -ForegroundColor Red
            $ReportLine = [PSCustomObject]@{
                GroupId = $Group.Id
                Name    = $GroupDisplayName
                Type    = "Microsoft 365 Group"
                Status  = "Failed"
                Error   = $_.Exception.Message
            }
            $Report.Add($ReportLine)
            Continue
        }
    }
    If ($Group.SecurityEnabled -and $null -eq $Group.MailEnabled) { # Security group
        Write-Host ("Adding {0}} to securty group {1}" -f $TargetUser.DisplayName, $Group.DisplayName) -ForegroundColor Cyan
        Try {
            New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $TargetUser.Id -ErrorAction Stop
            $ReportLine = [PSCustomObject]@{
                GroupId = $Group.Id
                Name    = $GroupDisplayName
                Type    = "Security Group"
                Status  = "Success"
                Error   = $null
            }
            $Report.Add($ReportLine)
            If ($RemoveOriginalUser) {
                $Outcome = Remove-UserFromM365Group -UserId $SourceUser.Id -GroupId $Group.Id -GroupName $GroupDisplayName
                Write-Host $Outcome -ForegroundColor Yellow
            }
        } Catch {
            Write-Host ("Failed to add {0} to security group {1}: {2}" -f $TargetUser.DisplayName, $GroupDisplayName, $_.Exception.Message) -ForegroundColor Red
              $ReportLine = [PSCustomObject]@{
                GroupId = $Group.Id
                Name    = $Group.DisplayName
                Type    = "Security Group"
                Status  = "Failed"
                Error   = $_.Exception.Message
            }
            $Report.Add($ReportLine)
            Continue
        }
    }
    If ($Group.SecurityEnabled -eq $false -and $Group.MailEnabled -eq $true -and -not($Group.groupTypes -contains "Unified")) { # distribution list
        Write-Host ("Adding {0} to distribution list {1}" -f $TargetUser.DisplayName, $GroupDisplayName) -ForegroundColor Cyan
        Try { 
            Add-DistributionGroupMember -Identity $Group.Id -Member $TargetUser.Id -ErrorAction Stop
            $ReportLine = [PSCustomObject]@{
                GroupId = $Group.Id
                Name    = $GroupDisplayName
                Type    = "Distribution List"
                Status  = "Success"
                Error   = $null
            }
            $Report.Add($ReportLine)
            If ($RemoveOriginalUser) {
                $Outcome = Remove-UserFromDistributionList -UserId $SourceUser.Id -GroupId $Group.Id -GroupName $GroupDisplayName
            }
        } Catch {
            Write-Host ("Failed to add {0} to distribution list {1}: {2}" -f $TargetUser.DisplayName, $Group.DisplayName, $_.Exception.Message) -ForegroundColor Red
            $ReportLine = [PSCustomObject]@{
                GroupId = $Group.Id
                Name    = $Group.DisplayName
                Type    = "Distribution List"
                Status  = "Failed"
                Error   = $_.Exception.Message
            }
            $Report.Add($ReportLine)
            Continue
        }
    }
    If ($Group.SecurityEnabled -eq $true -and $Group.MailEnabled -eq $true -and -not($Group.groupTypes -contains "Unified")) { # mail-enabled security group
        Write-Host ("Adding {0} to mail-enabled security group {1}" -f $TargetUser.DisplayName, $Group.DisplayName) -ForegroundColor Cyan
        Try {
            Add-DistributionGroupMember -Identity $Group.Id -Member $TargetUser.Id -ErrorAction Stop
              $ReportLine = [PSCustomObject]@{
                GroupId = $Group.Id
                Name    = $Group.DisplayName
                Type    = "Mail-enabled Security Group"
                Status  = "Success"
                Error   = $null
            }
            $Report.Add($ReportLine)
            If ($RemoveOriginalUser) {
                $Outcome = Remove-UserFromDistributionList -UserId $SourceUser.Id -GroupId $Group.Id -GroupName $GroupDisplayName
            }
        } Catch {
            Write-Host ("Failed to add {0} to mail-enabled security group {1}: {2}" -f $TargetUser.DisplayName, $Group.DisplayName, $_.Exception.Message) -ForegroundColor Red
            $ReportLine = [PSCustomObject]@{
                GroupId = $Group.Id
                Name    = $Group.DisplayName
                Type    = "Mail-enabled Security Group"
                Status  = "Failed"
                Error   = $_.Exception.Message
            }
            $Report.Add($ReportLine)
            Continue
        }
    }
}

Write-Host ("Finished copying groups from {0} to {1}." -f $SourceUser.DisplayName, $TargetUser.DisplayName) -ForegroundColor Green
Write-Host ("Membership was added to a total of {0} groups." -f $Report.Count) -ForegroundColor Green

$Report | Group-Object -Property Type | ForEach-Object {
    $Type = $_.Name
    $Count = $_.Count
    Write-Host ("{0} groups of type {1} were processed." -f $Count, $Type) -ForegroundColor Cyan
}

$Report | Out-GridView -Title ("Group memberships added for {0}" -f $TargetUser.DisplayName)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
