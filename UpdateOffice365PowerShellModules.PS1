# UpdateOffice365PowerShellModules.PS1
# Mentioned in Chapter 4 of Office 365 for IT Pros
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
