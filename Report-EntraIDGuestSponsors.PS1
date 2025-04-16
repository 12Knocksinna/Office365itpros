# Report-EntraIDGuestSponsors.PS1
# Report the sponsors assigned to Entra ID guest accounts

# V1.0 15-Apr-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-EntraIDGuestSponsors.PS1

Connect-MgGraph -NoWelcome -Scopes User.ReadWrite.All

[int]$Threshold = 500

Write-Host "Finding guest accounts to analyze..." -ForegroundColor Green
[array]$Guests = Get-MgUser -Filter "userType eq 'Guest'" -All `
    -Property Id, DisplayName, Sponsors, CreatedDateTime, SignInActivity, Mail -ExpandProperty Sponsors | Sort-Object DisplayName
If (!($Guests)) { 
    Write-Host "No guest accounts found." -ForegroundColor Red
}

Write-Host ("Checking {0} guest accounts..." -f $Guests.Count) -ForegroundColor Green
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Guest in $Guests) {
    $SponsorNames = $null
    If ($Null -eq $Guest.Sponsors.Id) {
        $SponsorNames = "No sponsor assigned"
    } Else {
        $SponsorNames = $Guest.Sponsors.additionalProperties.displayName -join ", "
    }

    $SignInDate = $null
    If ([string]::IsNullOrEmpty($Guest.SignInActivity.LastSuccessfulSignInDateTime)) {
        $SignInDate = "No sign-in activity"
        [int]$DaysSinceSignIn = (New-TimeSpan $Guest.CreatedDateTime).Days
    } Else {
        $SignInDate = Get-Date($Guest.SignInActivity.LastSuccessfulSignInDateTime) -format 'dd-MMM-yyyy HH:mm'  
        [int]$DaysSinceSignIn = (New-TimeSpan $SignInDate).Days
    }

    $ReportLine = [PSCustomObject] @{
        Name                 = $Guest.DisplayName
        Email                = $Guest.Mail
        'Sponsor Names'      = $SponsorNames
        Created              = Get-Date($Guest.CreatedDateTime) -format 'dd-MMM-yyyy HH:mm'
        'Last Sign In'       = $SignInDate
        'Days Since Sign In' = $DaysSinceSignIn.ToString()
    }
    $Report.Add($ReportLine)
}

$Report | Out-GridView -Title "Entra ID Guest Account Sponsors" 

# List all the guest accounts (and their sponsors) that haven't signed in for more than the threshold number of days
$OldGuests = $Report | Where-Object {$_.'Days Since Sign In' -as [int] -gt $Threshold}
Write-Host ""
Write-Host ("The following guest accounts have not signed in for more than {0}} days:" -f $Threshold) -ForegroundColor Red
Write-Host ""
$OldGuests | Format-Table Name, 'Sponsor Names', 'Days Since Sign In', 'Last Sign In' -AutoSize

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 