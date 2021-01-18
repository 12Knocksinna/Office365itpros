# UpdateUserPhotos.PS1
# A script to illustrate how to update user photos in Microsoft 365 accounts
# 
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateUserPhotos.PS1
#
Write-Host "Checking for connection to Exchange Online"
$Check = Get-Module -Name ExchangeOnlineManagement
If (!$Check) {
      Write-Host "Your PowerShell session is not connected to Exchange Online."
      Write-Host "Please connect to the Exchange Online Management module using an administrator account and retry."; Break }

# The value of $PhotoLocation should be changed to point to the place where user photos are stored
# We expect to find JPG files there named after the first and last names of the UPN assigned to accounts
# For example, the photo for Kim.Akers@Office365itpros.com is in Kim.Akers.jpg
$PhotoLocation = "c:\Temp\UserPhotos\" 
If (!(Test-Path ($PhotoLocation))) {
    Write-Host "Can't find $PhotoLocation - please check if this is the right place to find user photos"; break }
$i=0
# Find mailboxes
Write-Host "Finding user mailboxes"
$Users = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited
CLS
$ProgDelta = 100/($Users.Count); $CheckCount = 0; $UserNumber = 0
ForEach ($User in $Users) {
   $PhotoExists = $Null
   $UserNumber++
   $UserStatus = $User.DisplayName + " ["+ $UserNumber +"/" + $Users.Count + "]"
   Write-Progress -Activity "Checking photo for" -Status $UserStatus -PercentComplete $CheckCount
   $CheckCount += $ProgDelta
   # Is EXODS happy with the user photo information for the account?
   $CheckPhoto = Get-UserPhoto -Identity $User.Alias -ErrorAction SilentlyContinue
   If (!$CheckPhoto) { # No photo found in mailbox
      $UserPhoto = $PhotoLocation + $User.UserPrincipalName.Split("@")[0]+".jpg"
      If (Test-Path $UserPhoto) { # Update the photo because we have a file
        Write-Host "Updating photo for" $User.DisplayName -Foregroundcolor Red
        Set-UserPhoto -Identity $User.Alias -PictureData ([System.IO.File]::ReadAllBytes($UserPhoto)) -Confirm:$False 
        $i++; CLS }
      Else { # No photo file available
        Write-Host "No photo file available for" $User.DisplayName }
    } 
}
Write-Host "All done. $i User Photos updated"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
