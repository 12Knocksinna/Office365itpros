# UpdateOffice365PowerShellModules.PS1
# Mentioned in Chapter 4 of Office 365 for IT Pros
# V2.0 4-Dec-2022
# V2.1 9-Apr-2023
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateOffice365PowerShellModules.PS1
# Very simple script to check for updates to a defined set of PowerShell modules used to manage Office 365 services
# If an update for a module is found, it is downloaded and applied.
# Once all modules are checked for updates, we remove older versions that might be present on the workstation. V2.1 improves the processing of Microsoft Graph SDK sub-modules
#  
# Define the set of modules installed and updated from the PowerShell Gallery that we want to maintain - edit this set of modules to include the modules 
# you want to process.
$O365Modules = @("MicrosoftTeams", "Microsoft.Graph", "ExchangeOnlineManagement", "Microsoft.Online.Sharepoint.PowerShell", "ORCA", "Az.Accounts", "Az.Automation", "Az.Keyvault")

# We're installing from the PowerShell Gallery so make sure that it's trusted
Set-PSRepository -Name PsGallery -InstallationPolicy Trusted

# Check and update all modules to make sure that we're at the latest version
ForEach ($Module in $O365Modules) {
   Write-Host "Checking and updating module" $Module
   $CurrentVersion = (Find-Module -Name $Module | Select-Object -ExpandProperty Version)
   # Check current version on workstation and report what we find. If no module is installed, tell the user
   If ($CurrentVersion) {
     $CurrentVersion = $CurrentVersion.Major.toString() + "." + $CurrentVersion.Minor.toString() + "." + $CurrentVersion.Build.toString()
     Write-Host ("Current version of the {0} module in the PowerShell Gallery is {1}" -f $Module, $CurrentVersion)
   } Else {
     $CurrentVersion = "0.0.0"
     Write-Host ("No module found for {0}" -f $Module) }
   
   $PCVersion = Get-InstalledModule -Name $Module -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Version
   If ($PCVersion) {
      $PCVersion = $PCVersion.Major.toString() + "." + $PCVersion.Minor.toString() + "." + $PCVersion.Build.toString()
      If ($PCVersion -eq $CurrentVersion) { 
         Write-Host ("Version {0} is installed on this PC - no need to update" -f $PCVersion) 
      } Else {
         Write-Host ("Updating {0} module to version {1}" -f $Module, $CurrentVersion)
         Remove-Module $Module -ErrorAction SilentlyContinue
         Update-Module $Module -Force -Confirm:$False
      } # End if 
    } Else { # Module is not installed, so install it
         Write-Host ("Installing module {0}..." -f $Module)
         Install-Module $Module -Scope AllUsers -Confirm:$False
     } #End if
} # End ForEach Module

# Check and remove older versions of the modules from the PC
Write-Host "Beginning clean-up phase..."
[array]$InstalledModules = Get-InstalledModule
[array]$GraphModules = $InstalledModules | Where-Object {$_.Name -Like "*Microsoft.Graph*"} | Select-Object -ExpandProperty Name
$ModulesToProcess = $O365Modules + $GraphModules | Sort-Object -Unique

ForEach ($Module in $ModulesToProcess) {
   Write-Host "Checking for older versions of" $Module
   [array]$AllVersions = Get-InstalledModule -Name $Module -AllVersions -ErrorAction SilentlyContinue
   If ($AllVersions) {
     $AllVersions = $AllVersions | Sort-Object PublishedDate -Descending 
     $MostRecentVersion = $AllVersions[0].Version
     $MostRecentVersion = $MostRecentVersion.Major.toString() + "." + $MostRecentVersion.Minor.toString() + "." + $MostRecentVersion.Build.toString()
     $PublishedDate = (Get-Date($AllVersions[0].PublishedDate) -format g)
     Write-Host ("Most recent version of {0} is {1} published on {2}" -f $Module, $MostRecentVersion, $PublishedDate)
     If ($AllVersions.Count -gt 1 ) { # More than a single version installed
      ForEach ($Version in $AllVersions) { #Check each version and remove old versions
        If ($Version.Version -ne $MostRecentVersion)  { # Old version - remove
           Write-Host ("Uninstalling version {0} of module {1}" -f $Version.Version, $Module) -foregroundcolor Red 
           Uninstall-Module -Name $Module -RequiredVersion $Version.Version -Force
         } #End if
      } #End If
    } #End ForEach Module
  } #End If
} #End ForEach

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
