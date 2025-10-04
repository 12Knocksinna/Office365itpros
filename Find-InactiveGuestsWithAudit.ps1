# Find-InactiveGuestsWithAudit.PS1
# A script to find inactive Entra ID guests and report what they've been doing



# Find all guests - a complex query is used to sort the retrieved results
[array]$Guests = Get-MgUser -Filter "usertype eq 'Guest'" -PageSize 500 -All -Sort displayName -ConsistencyLevel eventual -CountVariable GuestCount
If ($Guests.Count -eq 0) {
    Write-Host "No guest users found."
    break
} Else {
    Write-Host ("Found {0} guest users" -f $Guests.Count)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.