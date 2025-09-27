# Remove-InactiveUserAccounts.PS1
# An Azure Automation runbbook to remove inactive user accounts from Entra ID
# 
# V1.0 27-Sep0-2025
# Github Link:


# Define theidentifier for the group holding details of accounts we want to review for inactivity. This could be a dynamic group.
$TargetGroupId = '1cfc9ab3-d230-45f6-a81f-8e32de9ad95b' 

# If no target group, process all licensed users
Try {
    [array]$TargetAccounts = Get-MgGroupMember -GroupId $TargetGroupId -All -PageSize 500 | Select-Object -ExpandProperty Id
} Catch {
    Write-Output "Error retrieving group containing target accounts: $_"
    Write-Output "Checking for all user accounts instead"
    Try {
        [array]$TargetAccounts = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member' and accountEnabled ne false" -ConsistencyLevel eventual `
        -CountVariable UsersFound -Property id, displayName, userprincipalname, usertype, signInActivity, SignInSessionsValidFromDateTime, LastPasswordChangeDateTime, passwordPolicies `
        -All -PageSize 500 -Sort displayName -ErrorAction Stop | Select-Object -ExpandProperty Id
    } Catch {
        Write-Output "Error retrieving all user accounts: $_"
        Break    
    }
}

Write-Output "Found $($TargetAccounts.Count) target accounts to check for inactivity"
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Id in $TargetAccounts) {

    $DaysSinceLastSignIn = $null; $DaysSinceLastSuccessfulSignIn = $null
    $DaysSincePasswordChange = $null; $PasswordPoliciesOutput = $null
    $DaysSinceLastSignIn = "N/A"; $DaysSinceLastSuccessfulSignIn = "N/A"
    
    If (!([string]::IsNullOrWhiteSpace($User.signInActivity.lastSuccessfulSignInDateTime))) {
        [datetime]$LastSuccessfulSignIn = $User.signInActivity.lastSuccessfulSignInDateTime
         $DaysSinceLastSuccessfulSignIn = (New-TimeSpan $LastSuccessfulSignIn).Days 
    }
    If (!([string]::IsNullOrWhiteSpace($User.signInActivity.lastSignInDateTime))) {
        [datetime]$LastSignIn = $User.signInActivity.lastSignInDateTime
        $DaysSinceLastSignIn = (New-TimeSpan $LastSignIn).Days
    }    



}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.