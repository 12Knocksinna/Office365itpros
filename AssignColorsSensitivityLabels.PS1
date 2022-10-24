# AssignColorsSensitivityLabels.PS1
# A script to assign suitable traffic light colors to sensitivity labels
# The script works by assigning thresholds for Green, Yellow, and Red label colors based on the priority order of the label. The code then
# checks existing colors assigned to labels used with items and assigns the approriaate label color.

# Define thresholds
[int]$Yellow = 5
[int]$Red = 13

# Define colors to use 

$GreenLabel = "#00FF00"
$YellowLabel = "#FFFF00"
$RedLabel =  "#FF0000"

# Create a hash table to translate color hex values to display values.
$Colors = @{
  "000000" = "Black"
  "393939" = "Charcoal"
  "0000FF" = "Blue"
  "3A96DD" = "Light Blue"
  "7160E8" = "Lavender"
  "C239B3" = "Berry"
  "7A7574" = "Beige"
  "FF0000" = "Red"
  "FFC0CB" = "Pink"
  "F7630C" = "Orange"
  "EAA300" = "Marigold"
  "A80000" = "Dark Red"
  "8b0000" = "Darker Red"
  "A4262C" = "Burgandy"
  "00FF00" = "Green"
  "13A10E" = "Light Green"
  "317100" = "Dark Green"
  "0078D7" = "Dark Blue"
  "8a2be2" = "Bright Violet"
  "FFFF00" = "Yellow"
  "859599" = "Silver"
}

Connect-ExchangeOnline
Connect-IPPSession

# Find the set of sensitivity labels and filter out those that can handle items (files, messages)
[array]$Labels = Get-Label
$ItemLabels = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Label in $Labels) { 
    If ($Label.ContentType -Like "*File, Email*") { # It's a label for items
    $ColorFound = $Null; $ColorDisplay = "No color defined"
    $ColorFound = ($Label.Settings | ? {$_ -match "color"})
    If ($ColorFound) {
    Try { 
        $ColorCode = $ColorFound.ToString().Split("#")[1].Split("]")[0] ; $ColorDisplay = $Colors[$ColorCode]
     }
     Catch {
        Write-Host "Error reading configuration for label" $L.DisplayName  
     }}
    $DataLine = [PSCustomObject] @{
       LabelId     = $Label.ImmutableId
       DisplayName = $Label.DisplayName
       Priority    = $Label.Priority 
       Color       = $ColorDisplay } 
      $ItemLabels.Add($DataLine) } 
}

Write-Host "Current Sensitivity Labels Defined for Items"
Write-Host "--------------------------------------------"
Write-Host ""
$ItemLabels | Format-Table DisplayName, Priority, Color
Write-Host ""
Write-Host ("{0} sensitivity labels found for item assignments. Updating them with new colors" -f $ItemLabels.Count)

ForEach ($Label in $ItemLabels) {
  Switch ($Label.Priority) 
   {
     ({$PSItem -le $Yellow})
      {
       Write-Host ("Setting label {0} to Green" -f $Label.DisplayName)
       Set-Label -Identity $Label.LabelId -AdvancedSettings @{color=$GreenLabel}
      }
     ({$PSItem -gt $Yellow -and $PSItem -le $Red})
      {
       Write-Host ("Setting label {0} to Yellow" -f $Label.DisplayName )
       Set-Label -Identity $Label.LabelId -AdvancedSettings @{color=$YellowLabel}
      }
     ({$PSItem -ge $Red}) 
      {
       Write-Host ("Setting Label {0} to Red" -f $Label.DisplayName )
       Set-Label -Identity $Label.LabelId -AdvancedSettings @{color=$RedLabel}
      }
   } # End Switch
} # End ForEach Label


Write-Host "All done. Labels now have traffic-light colors"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
