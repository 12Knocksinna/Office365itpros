# Copy-GroupsFromOneUsertoAnother.ps1
# Description: This script copies all groups from one user to another in Microsoft 365. 

# V1.0 1-Mar-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Copy-GroupsFromOneUsertoAnother.PS1

param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the UPN for the source user")]
    [string]$Source,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the UPN for target user")]
    [string]$Target
)

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
} Catch {
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
        Write-Host ("Adding {0} to Microsoft 365 group {1}" -f $TargetUser.displayName, $Group.DisplayName) -ForegroundColor Cyan
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
        } Catch {
            Write-Host ("Failed to add {0} to Microsoft 365 group {1}: {2}" -f $TargetUser.DisplayName, $Group.DisplayName, $_.Exception.Message) -ForegroundColor Red
            $ReportLine = [PSCustomObject]@{
                GroupId = $Group.Id
                Name    = $Group.DisplayName
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
                Name    = $Group.DisplayName
                Type    = "Security Group"
                Status  = "Success"
                Error   = $null
            }
            $Report.Add($ReportLine)
        } Catch {
            Write-Host ("Failed to add {0} to security group {1}: {2}" -f $TargetUser.DisplayName, $Group.DisplayName, $_.Exception.Message) -ForegroundColor Red
              $ReportLine = [PSCustomObject]@{
                GroupId = $Group.Id
                Name    = $Group.DisplayName
                Type    = "Microsoft 365 Group"
                Status  = "Failed"
                Error   = $_.Exception.Message
            }
            $Report.Add($ReportLine)
            Continue
        }
    }
    If ($Group.SecurityEnabled -eq $false -and $Group.MailEnabled -eq $true -and -not($Group.groupTypes -contains "Unified")) { # distribution list
        Write-Host ("Adding {0} to distribution list {1}" -f $TargetUser.DisplayName, $Group.DisplayName) -ForegroundColor Cyan
        Try { 
            Add-DistributionGroupMember -Identity $Group.Id -Member $TargetUser.Id -ErrorAction Stop
            $ReportLine = [PSCustomObject]@{
                GroupId = $Group.Id
                Name    = $Group.DisplayName
                Type    = "Distribution List"
                Status  = "Success"
                Error   = $null
            }
            $Report.Add($ReportLine)
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
                Type    = "Distribution List"
                Status  = "Success"
                Error   = $null
            }
            $Report.Add($ReportLine)
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


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.