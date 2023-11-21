# Remove-PermissionsSharedMailbox.PS1
# Sample script to show how to remove permissions for one or more users from shared mailboxes.
# https://github.com/12Knocksinna/Office365itpros/blob/master/Remove-PermissionsSharedMailbox.PS1
# V1.0 20-Nov-2023

# Two simple parameters - the set of user mailboces ($Mailboxes) to check against a set of shared mailboxes
# ($TargetSharedMailboxes). # Both can be 'All' meaning # that the script will check all of the objects of that type. 
# Or they can be a comma-separated lists of mailbox identities that the script wll resolve and then process.

param (
    [Parameter(Mandatory=$true)]
    [string]$Mailboxes,

    [Parameter(Mandatory=$true)]
    [string]$TargetSharedMailboxes
)

# Need a connection to Exchange Online
If (Get-ConnectionInformation) {
    Write-Host "Checking the target shared mailboxes and user accounts to process..." 
} Else {
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

$CSVOutputFile = "C:\temp\SharedMailboxPermissions.CSV"

# Check whether we need to process all shared mailboxes or just a selected set passed in a comma-seperated list in 
# the $TargetSharedMailboxes parameter
If ($TargetSharedMailboxes) {
    # Easy processing - get all shared mailboxes
    If ($TargetSharedMailboxes -eq "All") {
        [array]$TargetSM = Get-ExoMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited `
            | Select-Object ExternalDirectoryObjectId, DisplayName
    } Else {
        # Slightly more complicated. First, split the comma-separated list up to form an array
        $SharedMbx = $TargetSharedMailboxes.Split(",")
        $TargetSM = [System.Collections.Generic.List[Object]]::new()
        # Now try to get the details of each shared mailbox and add them to the array of mailboxes to process
        ForEach ($M in $SharedMbx) {
            $Mbx = Get-ExoMailbox -Identity $M.Trim() -ErrorAction SilentlyContinue
            If ($Mbx) {
                $SharedMailboxDetails = [PSCustomObject]@{
                    ExternalDirectoryObjectId = $Mbx.ExternalDirectoryObjectId
                    DisplayName               = $Mbx.DisplayName
                }
                $TargetSM.Add($SharedMailboxDetails)
            } Else {
                Write-Host ("Can't find the {0} shared mailbox" -f $M)
            }
        }
    }
}

$UsersToProcess = [System.Collections.Generic.List[Object]]::new() 
# Check that the set of input users are all OK
If ($Mailboxes -ne "All") {
    $UsersToCheck = $Mailboxes.Split(",")
    ForEach ($U in $UsersToCheck) {
        $TargetUser = Get-ExoMailbox -Identity $U.Trim() -ErrorAction SilentlyContinue
        If (!($TargetUser)) {
            Write-Host ("Can't find the target user {0} to process" -f $TargetUser)
        } Else {
            $UserDetails = [PSCustomObject]@{
                ExternalDirectoryObjectId = $TargetUser.ExternalDirectoryObjectId
                DisplayName               = $TargetUser.DisplayName
                UserPrincipalName         = $TargetUser.UserPrincipalName
                Alias                     = $TargetUser.Alias
                Name                      = $TargetUser.Name
            }
            $UsersToProcess.Add($UserDetails)
        }
    }    
} Else {
    Write-Host "Fetching all user mailboxes to check against..."
    [array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Sort-Object DisplayName
    If ($Mbx) {
        ForEach ($Mailbox in $Mbx) {
            $UserDetails = [PSCustomObject]@{
                ExternalDirectoryObjectId = $Mailbox.ExternalDirectoryObjectId
                DisplayName               = $Mailbox.DisplayName
                UserPrincipalName         = $Mailbox.UserPrincipalName
                Alias                     = $Mailbox.Alias
                Name                      = $Mailbox.Name
            }
            $UsersToProcess.Add($UserDetails)
        }
    }
}

# After all that, we should tell the administrator what are we going to do
Write-Host ""
Write-Host "Processing permission removal from shared mailboxes"
Write-Host "---------------------------------------------------"
Write-Host ""
Write-Host ("Permissions will be removed from the following shared mailboxes: {0}" -f ($TargetSM.DisplayName -join ", "))
Write-Host ("for the following mailboxes:                                     {0}" -f ($UsersToProcess.DisplayName -join ", "))
Write-Host ""
Write-Host "Starting..."

# Loop to go through each shared mailbox and remove the permissions for the users
[int]$i = 0
$RemovalResults = [System.Collections.Generic.List[Object]]::new() 
[int]$RemovedFullAccessCount = 0; [int]$RemovedSendOnBehalfOfCount = 0; [int]$RemovedSendAsCount = 0
ForEach ($SM in $TargetSM) {
    $i++
    Write-Host ("Checking permissions for shared mailbox {0} ({1}/{2})" -f $SM.DisplayName, $i, $TargetSM.count)
    $Permissions = Get-MailboxPermission -Identity $SM.ExternalDirectoryObjectId `
        | Where-Object {$_.User -ne 'NT AUTHORITY\SELF' -and $_.IsInherited -ne $true }
    ForEach ($U in $UsersToProcess) {
        $RemovedSendAs = $false; $RemovedSendOnBehalfOf = $False; $RemovedFullAccess = $False
        # Check if this user has permission for this mailbox
        If ($U.UserPrincipalName -in $Permissions.User) {
            Write-Host ("Found user {0} with full access permission for shared mailbox {1}" -f $U.DisplayName, $SM.DisplayName) -ForegroundColor Red
            Remove-MailboxPermission -Identity $SM.ExternalDirectoryObjectId -User $U.UserPrincipalName -AccessRights FullAccess -Confirm:$false
            $RemovedFullAccess = $True
            $RemovedFullAccessCount++
        }
        # Remove Send on behalf of permission if granted
        [array]$SendOnBehalfOf = Get-ExoMailbox -Identity $SM.UserPrincipalName -Properties GrantSendOnBehalfTo | Select-Object -ExpandProperty GrantSendOnBehalfTo
        If ($U.Alias -in $SendOnBehalfOf) {
            Write-Host ("Removing Send on Behalf Of permission for {0} from {1}" -f $U.DisplayName, $SM.DisplayName) -ForegroundColor Yellow
            Set-Mailbox -Identity $SM.ExternalDirectoryObjectId -GrantSendOnBehalfTo @{remove="$U.UserPrincipalName"} `
               -ErrorAction SilentlyContinue
            $RemovedSendOnBehalfOf = $True
            $RemovedSendOnBehalfOfCount++
        }
        # Remove Send As permission if granted
        If (Get-ExoRecipientPermission -Identity $SM.ExternalDirectoryObjectId -Trustee $U.UserPrincipalName -AccessRights SendAs) {
            Write-Host ("Removing Send As permission for {0} from {1}" -f $U.DisplayName, $SM.DisplayName) -ForegroundColor Yellow
            Remove-RecipientPermission -Identity $SM.ExternalDirectoryObjectId -AccessRights SendAs -Trustee $U.UserPrincipalName -Confirm:$False -ErrorAction SilentlyContinue
            $RemovedSendAs = $True
            $RemovedSendAsCount++
        }
        # Capture record if a permission was removed
        If (($RemovedFullAccess -eq $True) -or ($RemovedSendAs -eq $True) -or ($RemovedSendOnBehalfOf -eq $True)) {
            $RemovalDetails = [PSCustomObject]@{
                User                       = $U.UserPrincipalName
                Name                       = $U.DisplayName
                'Shared Mailbox'           = $SM.DisplayName
                'Remove Full Access'       = $RemovedFullAccess
                'Remove Send As'           = $RemovedSendAs
                'Remove Send on Behalf of' = $RemovedSendOnBehalfOf
                TimeStamp                  = (Get-Date -format s)
            }
            $RemovalResults.Add($RemovalDetails)
        }
    } # End ForEach User Mailbox
} # End ForEach Shared Mailbox

Write-Host ""
Write-Host ("Number of Full access permissions removed:       {0}" -f $RemovedFullAccessCount)
Write-Host ("Number of Send As permissions removed:           {0}" -f $RemovedSendAsCount)
Write-Host ("Number of Send On Behalf Of permissions removed: {0}" -f $RemovedSendOnBehalfOfCount)
Write-Host ""
$RemovalResults | Format-Table User, 'Shared mailbox', 'Remove Full Access', 'Remove Send As', 'Remove Send on Behalf Of' -AutoSize
$RemovalResults | Export-CSV -NoTypeInformation $CSVOutputFile
Write-Host
Write-Host ("CSV file available with detailed permission removal results in {0}" -f $CSVOutputFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.