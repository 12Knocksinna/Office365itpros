# Update-GroupPhotos.PS1
# Example script to show how to check the photos for Microsoft 365 Groups against a list of approved photos
# and update them if necessary
# https://github.com/12Knocksinna/Office365itpros/blob/master/Update-GroupPhotos.PS1
Connect-ExchangeOnline
# Define groups we want to find

$LimitedAccessId = "d6cfd185-f31c-4508-ae40-229ff18a9919"
$ConfidentialAccessId = "c99e52c6-f5ff-4050-9313-ca6a3a35710f"
$PhotoDirectory = "c:\temp\"

[array]$Groups = Get-UnifiedGroup -ResultSize Unlimited 
$Groups = $Groups | Where-Object {$_.SensitivityLabel -eq $LimitedAccessId -or $_.SensitivityLabel -eq $ConfidentialAccessId} | Sort-Object DisplayName
Write-Host ("Scanning {0} groups to check photos..." -f $Groups.count)

# Read in data about approved photos
[array]$GroupPhotos = Import-csv c:\temp\GroupPhotos.csv

ForEach ($Group in $Groups) {
  $ExistingPhotoData = $Null
  # Do we have some photo data?
  $Photo = $GroupPhotos | Where-Object {$_.ExternalDirectoryObjectId -eq $Group.ExternalDirectoryObjectId}
  If ($Photo) { # We do!
     $PhotoThumbPrint = $Photo.Thumbprint    
     If ($Photo.Photo) {
        $PhotoFile = $PhotoDirectory + $Photo.Photo 
     } Else {
        $PhotoFile = $Null 
        Write-Host ("Group {0} requires an approved photo but no entry is available in the photo list" -f $Group.DisplayName)
    }
     # Check if a photo file is where we expect it to be 
     If ((Test-Path $PhotoFile) -eq $False) { 
        Write-Host ("Group {0} requires an approved photo but the expected file is not available in {1}" -f $Group.DisplayName, $PhotoFile)
         $PhotoFile = $Null 
     }
     Write-Host ("Checking photo for group {0}" -f $Group.DisplayName)
     $ExistingPhotoData = Get-UserPhoto -Identity $Group.ExternalDirectoryObjectId -GroupMailbox -ErrorAction SilentlyContinue
     If ($ExistingPhotoData.Thumbprint -eq $PhotoThumbprint) {
        Write-Host ("Group {0} has the approved photo" -f $Group.DisplayName)
     }
     If ($ExistingPhotoData.Thumbprint -ne $PhotoThumbprint -and $Null -ne $PhotoFile) { # Thumbprints don't match, so update with approved image
        Write-Host ("Thumbprint mismatch: Updating photo for {0} with {1}" -f $Group.DisplayName, $PhotoFile) -ForegroundColor Red
        Set-UserPhoto -GroupMailbox -Identity $Group.ExternalDirectoryObjectId -ErrorAction SilentlyContinue -PictureData ([System.IO.File]::ReadAllBytes($PhotoFile)) -Confirm:$False
        $ExistingPhotoData = Get-UserPhoto -Identity $Group.ExternalDirectoryObjectId -GroupMailbox -ErrorAction SilentlyContinue
        $Photo | Add-Member -NotePropertyName Thumbprint -NotePropertyValue $ExistingPhotoData.Thumbprint -force
     } ElseIf ($ExistingPhotoData.Thumbprint -ne $PhotoThumbprint -and $Null -eq  $PhotoFile) {
        Write-Host ("Thumbprint mismatch detected for group {0} but no photo file available" -f $Group.displayName) -ForegroundColor Yellow
     } # End if for thumbprint check
       
 } # End if photo
} # End Foreach group

# Export group data with adjusted thumbprints
$GroupPhotos | Export-CSV -NoTypeInformation c:\temp\GroupPhotos.csv -Force

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
