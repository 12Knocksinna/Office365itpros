# Report-LastSuccessfulSignIn.PS1
# Demo script to show the use of the new (from 1-Dec-2023) lastSuccessfulSignInDateTime property currently available in beta
# 6-Dec-2023
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-LastSuccessfulSignIn.PS1

# Connect to the Graph SDK with the correct permissions
Connect-MgGraph -NoWelcome -Scopes AuditLog.Read.All, Directory.Read.All

# Find licensed user accounts
$Headers = @{ConsistencyLevel="Eventual"}  
$Uri = "https://graph.microsoft.com/beta/users?`$count=true&`$filter=(assignedLicenses/`$count ne 0 and userType eq 'Member')&$`top=999&`$select=id, displayName, usertype, signInActivity"
[array]$Data = Invoke-MgGraphRequest -Uri $Uri -Headers $Headers
[array]$Users = $Data.Value

If (!($Users)) {
    Write-Host "Can't find any users... exiting!" ; break
}

# Paginate until we have all the user accounts
While ($Null -ne $Data.'@odata.nextLink') {
    Write-Host ("Fetching more user accounts - currently at {0}" -f $Users.count)
    $Uri = $Data.'@odata.nextLink'
    [array]$Data = Invoke-MgGraphRequest -Uri $Uri -Headers $Headers
    $Users = $Users + $Data.Value
 }
 Write-Host ("All available user accounts fetched ({0}) - now processing sign in report" -f $Users.count)

 # And report what we've found
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    $DaysSinceLastSignIn = $Null; $DaysSinceLastSuccessfulSignIn = $Null
    $DaysSinceLastSignIn = "N/A"; $DaysSinceLastSuccessfulSignIn = "N/A"
    $LastSuccessfulSignIn = $User.signInActivity.lastSuccessfulSignInDateTime
    $LastSignIn = $User.signInActivity.lastSignInDateTime
    If (!([string]::IsNullOrWhiteSpace($LastSuccessfulSignIn))) {
        $DaysSinceLastSuccessfulSignIn = (New-TimeSpan $LastSuccessfulSignIn).Days 
    }
    If (!([string]::IsNullOrWhiteSpace($LastSignIn))) {
        $DaysSinceLastSignIn = (New-TimeSpan $LastSignIn).Days
    }    
    $DataLine = [PSCustomObject][Ordered]@{
        User = $User.displayName
        UserId = $User.ID
        'Last successful sign in'        = $LastSuccessfulSignIn
        'Last sign in'                   = $LastSignIn
        'Days since successful sign in'  = $DaysSinceLastSuccessfulSignIn
        'Days since sign in'             = $DaysSinceLastSignIn
    }
    $Report.Add($DataLine)
}

$Report | Sort-Object 'Days since sign in' | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
