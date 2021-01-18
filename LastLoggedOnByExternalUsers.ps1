# LastLoggedInByExternalUsers
# Find the last time that external users (guest accounts) logged into our Office 365 tenant
# https://github.com/12Knocksinna/Office365itpros/blob/master/LastLoggedOnByExternalUsers.ps1

$Guests = (Get-AzureADUser -Filter "UserType eq 'Guest'" -All $True| Select Displayname, Mail, RefreshTokensValidFromDateTime | Sort RefreshTokensValidFromDateTime)
Write-Host $Guests.Count "guest accounts found. Checking last connections..."
$StartDate = (Get-Date).AddDays(-90)
$StartDate2 = (Get-Date).AddDays(-10)
$EndDate = (Get-Date).AddDays(+1)
$Active = 0
$EmailActive = 0
$Inactive = 0
$TeamsSpo = 0

ForEach ($G in $Guests) {
    Write-Host "Checking" $G.DisplayName  
    $Recs = $Null
    $UserId = $G.Mail
    # Handle account whose guest invitation is not redeemed
    If ($Userid -eq $Null) {$UserId = "NullString"}
    $Recs = (Search-UnifiedAuditLog -UserIds $UserId -Operations UserLoggedIn, TeamsSessionStarted -StartDate $StartDate -EndDate $EndDate)
    If ($Recs -eq $Null) {
       Write-Host "No connections found in the last 90 days for" $G.DisplayName "created on" $G.RefreshTokensValidFromDateTime -Foregroundcolor Red
       # Check email tracking logs because guests might receive email from Groups. Account must be fully formed for the check. We can only go back 10 days
       If ($UserId -ne "NullString") {
          $EmailRecs = (Get-MessageTrace –StartDate $StartDate2 –EndDate $EndDate -Recipient $G.Mail)
            If ($EmailRecs.Count -gt 0) {
            Write-Host "Email traffic found for " $G.DisplayName "at" $EmailRecs[0].Received -foregroundcolor Yellow
            $Active++ 
            $EmailActive++ }}
    }
    Elseif ($Recs[0].CreationDate -ne $Null) {
       Write-Host "Last connection for" $G.DisplayName "on" $Recs[0].CreationDate "as" $Recs[0].Operations -Foregroundcolor Green
       $Active++
       $TeamsSpo++ }
    
}
Write-Host ""
Write-Host "Statistics"
Write-Host "----------"
Write-Host "Guest Accounts          " $Guests.Count
Write-Host "Active Guests           " $Active
Write-Host "Active on Teams and SPO " $TeamsSPO
Write-Host "Active on Email         " $EmailActive
Write-Host "InActive Guests         " ($Guests.Count - $Active)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
