# A script to fetch the daily Bing photo and put it into the Teams background images folder. Images are fetched for the last
# seven days. See https://www.codeproject.com/Tips/1044421/Use-Bing-Photo-of-the-Day-in-Your-Application for more information
# https://github.com/12Knocksinna/Office365itpros/blob/master/GetBingImagesTeamsBackgrounds.PS1
# Define target folder for Teams background images and Bing market code for the images (could be something like fr-FR for France)
$TeamsBackgroundFiles = $env:APPDATA + "\Microsoft\Teams\Backgrounds\Uploads\" 
$Market = "en-US" 
# Check that the Teams background images folder exists. If not, create it
If (-not (Test-Path -LiteralPath $TeamsBackgroundFiles)) {
    Try {
        New-Item -Path $TeamsBackgroundFiles -ItemType Directory -ErrorAction Stop | Out-Null
    }
    Catch {
        Write-Error -Message "Unable to create directory '$TeamsBackgroundFiles'. Error was: $_" -ErrorAction Stop }
    Write-Host "Folder to store Teams background image files created: '$TeamsBackgroundFiles'" }
Else {
     Write-Host "Folder for Teams background images exists"
}
# Download the last seven days of Bing images
CLS ; For ($i=0; $i -le 7; $i++) {
  $BingUri = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=$i&n=1&mkt=$Market"
  $BingResponse = Invoke-WebRequest -Method Get -Uri $BingUri
  $BingContent = ConvertFrom-Json -InputObject $BingResponse.Content # Unpack content
  $BingBackgroundFile = "https://www.bing.com/"+$BingContent.Images.Url
  $BingFileName = $BingContent.Images.UrlBase.Split(".")[1]; $BingFileName = $BingFileName.Split("_")[0]+".jpg" 
  $TeamsBackgroundFile = $TeamsBackgroundFiles + "Bing-" + $BingFileName
  If (([System.IO.File]::Exists($TeamsBackgroundFile) -eq $False)) { 
     # File isn't there, so we can download
     Try {
         Invoke-WebRequest -Method Get -Uri $BingBackgroundFile -OutFile $TeamsBackgroundFile 
         Write-Host "Downloaded new Bing image" $TeamsBackgroundFile
         }
     Catch {
        Write-Host "Error occurred when downloading image from Bing" }
     } #End If
} #End loop

# Clean up Bing images older than 30 days so we keep the number of background images to a reasonable number
$RemoveDate = (Get-Date).AddDays(-30)
$BingFiles = Get-ChildItem $TeamsBackgroundFiles | ? {($_.LastWriteTime -lt $RemoveDate) -and ($_.Name.Substring(0,5) -eq "Bing-")}
If ($BingFiles) {
   Write-Host "Cleaning up old Bing background images"
   ForEach ($File in $BingFiles) { Del $File.FullName -Force}
}

Write-Host "All done"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
