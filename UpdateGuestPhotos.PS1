# UpdateGuestPhotos.PS1
# A script to demonstrate how to update photos for Azure AD Guest Accounts. The photos are stored in 
# c:\UserPhotos\ (definable location) and are named after the guest account (as in Tony Redmond.jpg).
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateGuestPhotos.PS1

Connect-MgGraph -Scopes User.ReadWrite.All
Select-MgProfile Beta

$PhotoLocation = "c:\UserPhotos\"
If (!(Test-Path -Path $PhotoLocation)) {
    Write-Host "Can't find $PhotoLocation - please check if this is the right place to find user photos"; break }

$DefaultGuestPhoto = "C:\UserPhotos\DefaultGuestPicture.jpg"
If (!(Test-Path -PathType Leaf -Path $DefaultGuestPhoto)) {
   Write-Host ("Can't find the default guest photo file {0} - please check if this is the right place to find user photos" -f $DefaultGuestPhoto) ; break }

[int]$i = 0; [int]$GuestNumber = 0; [int]$z=0; [int]$u = 0
[array]$NewGuestPhotos = $Null; [array]$UpdatedGuestPhotos = $Null
Write-Host "Finding guest accounts..."
[array]$Guests = Get-MgUser -Filter "Usertype eq 'Guest'" -All | Sort-Object DisplayName
CLS
$ProgDelta = 100/($Guests.Count); $CheckCount = 0;
ForEach ($Guest in $Guests) {
   $GuestNumber++
   $GuestStatus = $Guest.DisplayName + " ["+ $GuestNumber +"/" + $Guests.Count + "]"
   Write-Progress -Activity "Checking photo for guest account " -Status $GuestStatus -PercentComplete $CheckCount
   $CheckCount += $ProgDelta
   $PhotoExists = $False; $PhotoAvailable = $False; $NewerPhotoAvailable = $False
   $PhotoExists = Get-MgUserPhoto -UserId $Guest.Id -ErrorAction SilentlyContinue
   # Calculate the name of the photo file in the library and check the date of the file (if available)
   If ($Guest.DisplayName.IndexOf("(") -gt 0) {
      $PhotoFile = $PhotoLocation + $Guest.DisplayName.Split("(").trim()[0]+".jpg" 
   } Else {
      $PhotoFile = $PhotoLocation + $Guest.DisplayName + ".jpg" }
   If (Test-Path -Path $PhotoFile -PathType Leaf -ErrorAction SilentlyContinue) { # Check if a photo is available for this guest account
      $PhotoAvailable = $True
      [datetime]$LastTimePhotoFileUpdated = (Get-ChildItem $PhotoFile).LastWriteTime
      $UpdateText = ($Guest | Select-Object -ExpandProperty OnPremisesExtensionAttributes).ExtensionAttribute10
        If ($UpdateText) { # if an update ever occurred for the account, extract the date from the text
         [datetime]$LastTimeAccountPhotoUpdated = $UpdateText.Substring(0,19) 
         If ($LastTimeAccountPhotoUpdated -lt $LastTimePhotoFileUpdated)  { # Check if library file is newer than account photo
             $NewerPhotoAvailable = $True }
      } ElseIf ($UpdateText -eq $Null) { $NewerPhotoAvailable = $True }
    }  

   If (!($PhotoExists)) { # a photo is not present in this guest account
      If ($PhotoAvailable) {  
        Write-Host ("Adding a user photo to guest account {0} from {1}" -f $Guest.DisplayName, $PhotoFile) -ForeGroundColor Red
        Try {
           Set-MgUserPhotoContent -UserId $Guest.Id -Infile $PhotoFile; $i++
           $Now = Get-Date -format s; $PhotoUpdateText = ("{0} User Photo Updated" -f $Now)
           # Update account with date and time when we updated the photo
           Update-MgUser -UserId $Guest.Id -OnPremisesExtensionAttributes @{'extensionAttribute10' = "$($PhotoUpdateText)"} -ShowInAddressList
           $NewGuestPhotos += $Guest.DisplayName }
        Catch { 
           Write-Host ("Error adding photo to guest account {0}" -f $Guest.DisplayName) }
        }  Else { 
        Write-Host "Photo does not exist for" $Guest.DisplayName "- updating with default guest logo"
        Set-MgUserPhotoContent -UserId $Guest.Id -Infile $DefaultGuestPhoto; $z++
      }
    } #End if photo does not exist in account
  
   If (($PhotoExists) -and ($NewerPhotoAvailable)) { #Newer photo available in the library, so let's update it
       Write-Host ("Updating the user photo for guest account {0} from {1}" -f $Guest.DisplayName, $PhotoFile) -ForeGroundColor Yellow
       Try {
          Set-MgUserPhotoContent -UserId $Guest.Id -Infile $PhotoFile; $u++
          $Now = Get-Date -format s; $PhotoUpdateText = ("{0} User Photo Updated" -f $Now)
          # Update account with date and time when we updated the photo
          Update-MgUser -UserId $Guest.Id -OnPremisesExtensionAttributes @{'extensionAttribute10' = "$($PhotoUpdateText)"} -ShowInAddressList 
          $UpdatedGuestPhotos += $Guest.DisplayName }
       Catch 
         { Write-Host ("Error updating photo for guest account {0}" -f $Guest.DisplayName) }
   }

    Sleep -Seconds 1 
} #End ForEach Guest

Write-Host ""
Write-Host "Processing run complete:"
Write-Host "------------------------"
Write-Host "New photos added to accounts:    " $i
Write-Host "Updated photos for accounts:     " $u
Write-Host "Default image added to accounts: " $z
Write-Host ""

Write-Host "New photos added to: " ($NewGuestPhotos -Join ", ")
Write-Host "Photos updated for:  " ($UpdatedGuestPhotos -Join ", ")

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.

