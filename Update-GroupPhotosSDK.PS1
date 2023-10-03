# Update-GroupPhotos.PS1
# Example script to show how to check the photos for Microsoft 365 Groups against a list of approved photos
# and update them if necessary
# https://github.com/12Knocksinna/Office365itpros/blob/master/Update-GroupPhotosSDK.PS1
# V2 Upgraded 3-Oct-2023 for the Microsoft Graph PowerShell SDK

# Define the sensitivity labels for the groups that we want to find
# In this example, we use the identifiers for the Limited Access and Confidential Access labels
[array]$LabelsToCheck = "d6cfd185-f31c-4508-ae40-229ff18a9919", "c99e52c6-f5ff-4050-9313-ca6a3a35710f"
$PhotoDirectory = "c:\temp\GroupPhotos\"

Connect-MgGraph -Scopes Directory.ReadWrite.All -NoWelcome

[int]$i=0
[array]$Groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'unified')" -All
# Filter groups to extract those with the desired labels
$Groups = $Groups | Where-Object {$_.AssignedLabels.LabelId -in $LabelsToCheck} | Sort-Object DisplayName
Write-Host ("Scanning {0} groups to check photos..." -f $Groups.count)

# Read in data about approved photos
[array]$GroupPhotos = Import-csv c:\temp\GroupPhotos.csv

ForEach ($Group in $Groups) {
  $ExistingPhotoData = $Null
  # Do we have some photo data?
  $Photo = $GroupPhotos | Where-Object {$_.Id -eq $Group.Id}
  If ($Photo) { # We do!
     $PhotoFile = ("{0}{1}.jpg" -f $PhotoDirectory, $Photo.Id)     
     # Check if a photo file is where we expect it to be 
     If ((Test-Path $PhotoFile) -eq $False) { 
        Write-Host ("Group {0} requires an approved photo but the expected file is not available in {1}" -f $Group.DisplayName, $PhotoFile)
         $PhotoFile = $Null 
     }
     Write-Host ("Checking photo for group {0}" -f $Group.DisplayName)
     $ExistingPhotoData = Get-MgGroupPhoto -GroupId $Group.Id -ErrorAction SilentlyContinue
     If ($ExistingPhotoData) {
        Write-Host ("Group {0} has a photo" -f $Group.DisplayName)
     } Else {
        Set-MgGroupPhotoContent -GroupId $Group.Id -InFile $PhotoFile
        $i++
        $Photo | Add-Member -NotePropertyName Photo -NotePropertyValue $PhotoFile -Force
     } 
   } Else { # No photo available
      Write-Host ("Group {0} requires an approved photo but no entry is available in the photo list" -f $Group.DisplayName)
   }

} # End Foreach group

Write-Host ("All done. {0} photos updated." -f $i)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
