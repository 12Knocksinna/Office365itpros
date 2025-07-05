# Copy-GroupsFromOneUsertoAnother.ps1
# Description: This script copies all groups from one user to another in Microsoft 365. 

# V1.0 1-Mar-2025
# GitHub link:

param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the UPN for the source user")]
    [string]$Source,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the UPN for target user")]
    [string]$Target
)

Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "GroupMember.ReadWrite.All" -NoWelcome

# Check if the source user exists  
Try {
    $SourceUser = Get-MgUser -UserId $Source -ErrorAction Stop -Property DisplayName, UserPrincipalName, Id
} catch {
    Write-Host "Source user not found. Please check the User Principal Name." -ForegroundColor Red
    Break
}
# Check if the target user exists
Try {      
    $TargetUser = Get-MgUser -UserId $Target -ErrorAction Stop -Property DisplayName, UserPrincipalName, Id
} catch {
    Write-Host "Target user not found. Please check the User Principal Name." -ForegroundColor Red
    Break
}

Write-Host ("Checking groups for user {0} to copy to {1}" -f $SourceUser.DisplayName, $TargetUser.DisplayName)
[array]$SourceGroups = Get-MgUserMemberOf -UserId $SourceUser.Id -All -PageSize 500 | `
    Where-Object {$_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group'} | `
    Select-Object -ExpandProperty Id
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

ForEach ($GroupId in $GroupsToProcess) {
    Try {
        $Group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop -Property DisplayName, Id, GroupTypes, MailEnabled, SecurityEnabled
    } Catch {
        Write-Host ("Failed to retrieve group {0}: {1}" -f $GroupId, $_.Exception.Message) -ForegroundColor Red
        Continue
    }
    If ($Group.GroupTypes -contains "Unified") { # Microsoft 365 group 
        Write-Host ("Adding user to Microsoft 365 group {0}" -f $Group.DisplayName) -ForegroundColor Cyan
        Try {
            # New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $TargetUser.Id -ErrorAction Stop
        } Catch {
            Write-Host ("Failed to add {0} to Microsoft 365 group {1}: {2}" -f $TargetUser.DisplayName, $Group.DisplayName, $_.Exception.Message) -ForegroundColor Red
        }
        Continue
    }
    If ($Group.SecurityEnabled -and $null -eq $Group.MailEnabled) { # Security group
        Write-Host ("Adding user to securty group {0}" -f $Group.DisplayName) -ForegroundColor Cyan
        Try {
           # New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $TargetUser.Id -ErrorAction Stop
        } Catch {
            Write-Host ("Failed to add {0} to security group {1}: {2}" -f $TargetUser.DisplayName, $Group.DisplayName, $_.Exception.Message) -ForegroundColor Red
        }
        Continue
    }
    If ($Group.SecurityEnabled -eq $false -and $Group.MailEnabled -eq $true) { # distribution list
        Write-Host ("Adding user to distribution list {0}" -f $Group.DisplayName) -ForegroundColor Cyan
        # Add-DistributionGroupMember -Identity $Group.DisplayName -Member $TargetUser.Id -ErrorAction Stop
        Continue
    }
    If ($Group.SecurityEnabled -eq $true -and $Group.MailEnabled -eq $true) { # mail-enabled security group
        Write-Host ("Adding user to mail-enabled security group {0}" -f $Group.DisplayName) -ForegroundColor Cyan
        # Add-DistributionGroupMember -Identity $Group.DisplayName -Member $TargetUser.Id -ErrorAction Stop
        Continue
    }

}


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.