# UpdateOffice365PowerShellModules.PS1
# Mentioned in Chapter 4 of Office 365 for IT Pros
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateOffice365PowerShellModules.PS1
# Very simple script to check for updates to a defined set of PowerShell modules used to manage Office 365 services
# If an update for a module is found, it is downloaded and applied.
# Once all modules are checked for updates, we remove older versions that might be present on the workstation
#
# Define the set of modules installed and updated from the PowerShell Gallery that we want to maintain
$O365Modules = @("MicrosoftTeams", "MSOnline", "AzureADPreview", "ExchangeOnlineManagement", "Microsoft.Online.Sharepoint.PowerShell", "ORCA")

# Check and update all modules to make sure that we're at the latest version
ForEach ($Module in $O365Modules) {
   Write-Host "Checking and updating module" $Module
   Update-Module $Module -Force -Scope AllUsers }

# Check and remove older versions of the modules from the PC
ForEach ($Module in $O365Modules) {
   Write-Host "Checking for older versions of" $Module
   $AllVersions = Get-InstalledModule -Name $Module -AllVersions
   $AllVersions = $AllVersions | Sort PublishedDate -Descending 
   $MostRecentVersion = $AllVersions[0].Version
   Write-Host "Most recent version of" $Module "is" $MostRecentVersion "published on" (Get-Date($AllVersions[0].PublishedDate) -format g)
   If ($AllVersions.Count -gt 1 ) { # More than a single version installed
      ForEach ($Version in $AllVersions) { #Check each version and remove old versions
        If ($Version.Version -ne $MostRecentVersion)  { # Old version - remove
           Write-Host "Uninstalling version" $Version.Version "of Module" $Module -foregroundcolor Red 
           Uninstall-Module -Name $Module -RequiredVersion $Version.Version -Force
         } #End if
      } #End ForEach
  } #End If
} #End ForEach

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
