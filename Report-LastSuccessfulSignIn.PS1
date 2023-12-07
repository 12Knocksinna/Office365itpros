# Report-LastSuccessfulSignIn.PS1
# Demo script to show the use of the new (from 1-Dec-2023) lastSuccessfulSignInDateTime property currently available in beta
# 6-Dec-2023
#

$Headers = @{ConsistencyLevel="Eventual"}  
$Uri = "https://graph.microsoft.com/beta/users?`$count=true&`$filter=(assignedLicenses/`$count ne 0 and userType eq 'Member')&$`top=999&`$select=id, displayName, usertype, signInActivity"
[array]$Data = Invoke-MgGraphRequest -Uri $Uri -Headers $Headers
[array]$Users = $Data.Value

$Data = Invoke-MgGraphRequest -Uri $Uri -Method Get -Headers @{"ConsistencyLevel" = "eventual"}
[array]$Users = $Data.Value
If (!($Users)) {
    Write-Host "Can't find any users... exiting!" ; break
}

$Report = [System.Collections.Generic.List[Object]]::new()
Write-Host ("Processing {0} users..." -f $Users.count)

ForEach ($User in $Users) {
    $DaysSinceLastLogOn = $Null; $DaysSinceLastSuccessfulSignIn = $Null
    $LastSuccessfulSignIn = $User.signInActivity.lastSuccessfulSignInDateTime
    $LastLogin = $User.signInActivity.lastSignInDateTime
    If (!([string]::IsNullOrWhiteSpace($LastSuccessfulSignIn))) {
        $DaysSinceLastSuccessfulSignIn = (New-TimeSpan $LastSuccessfulSignIn).Days 
        $DaysSinceLastLogOn = (New-TimeSpan $LastLogin).Days
    }    
    $DataLine = [PSCustomObject][Ordered]@{
        User = $User.displayName
        UserId = $User.ID
        'Last successful sign in'   = $LastSuccessfulSignIn
        'Last successful log on'    = $LastLogin
        'Days since sign in'        = $DaysSinceLastSuccessfulSignIn
        'Days since log on'         = $DaysSinceLastLogOn
    }
    $Report.Add($DataLine)
}

$Report | Sort-Object 'Days since sign in' | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.