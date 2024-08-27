# UpdateOffice365PowerShellModules.PS1
# Mentioned in Chapter 4 of Office 365 for IT Pros
# V2.0 4-Dec-2022
# V2.1 9-Apr-2023
# V2.2 28-Apr-2023
# V2.3 9-Jul-2023 (include support for Microsoft Graph PowerShell SDK V2.0)
# V2.4 12-Aug (make sure that everything works well for PowerShell V5 and V7 )
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateOffice365PowerShellModules.PS1
# Very simple script to check for updates to a defined set of PowerShell modules used to manage Office 365 services
# If an update for a module is found, it is downloaded and applied.
# Once all modules are checked for updates, we remove older versions that might be present on the workstation. V2.1 improves the processing of Microsoft Graph SDK sub-modules
#  

# Define the set of modules installed and updated from the PowerShell Gallery that we want to maintain - edit this set of modules to include the modules 
# you want to process.
[int]$InstalledModules = 0; [int]$UpdatedModules = 0; [int]$RemovedModules = 0
$O365Modules = @("MicrosoftTeams", "Microsoft.Graph", "Microsoft.Graph.Beta", "ExchangeOnlineManagement", "Microsoft.Online.Sharepoint.PowerShell", "ORCA", "Az.Accounts", "Az.Automation", "AIPService", "Az.Keyvault", "Pnp.PowerShell", "MSCommerce", "Microsoft365DSC", "MSAL.PS", "Pnp.PowerShell", "WhiteboardAdmin", "ImportExcel")
$O365Modules = $O365Modules | Sort-Object
Write-Host ("Starting up and preparing to process these modules: {0}" -f ($O365Modules -join ", ")) -foregroundcolor Yellow
[int]$UpdatedModules = 0; [int]$RemovedModules = 0; [int]$InstalledModules = 0

# We're installing from the PowerShell Gallery so make sure that it's trusted
Set-PSRepository -Name PsGallery -InstallationPolicy Trusted

# Check and update all modules to make sure that we're at the latest version
ForEach ($Module in $O365Modules) {
   Write-Host "Checking and updating module" $Module
   $CurrentModule = Find-Module -Name $Module
   If ($CurrentModule) {
     $CurrentVersion = $CurrentModule.Version
      If ($CurrentVersion -isnot [string]) {
        $CurrentVersion = $CurrentVersion.Major.toString() + "." + $CurrentVersion.Minor.toString() + "." + $CurrentVersion.Build.toString()
      }
     [datetime]$CurrentModuleDate = $CurrentModule.PublishedDate
     Write-Host ("Current version of the {0} module in the PowerShell Gallery is {1}" -f $Module, $CurrentVersion)
   }

   $PCModule = Get-InstalledModule -Name $Module -ErrorAction SilentlyContinue

   If (!($PCModule)) { 
   # No version of the module found. It's in our list, so we install it.
     Write-Host ("No module found on this PC for {0}" -f $Module)
     Write-Host ("Installing module {0}..." -f $Module)  -foregroundcolor Yellow
     Install-Module $Module -Scope AllUsers -Confirm:$False -AllowClobber -Force
     $InstalledModules++
   }

   If ($PCModule) {
      $PCVersion = $PCModule.Version
      If ($PCVersion -isnot [string]) {
         $PCVersion = $PCVersion.Major.toString() + "." + $PCVersion.Minor.toString() + "." + $PCVersion.Build.toString()
      }
      [datetime]$PCModuleDate = $PCModule.PublishedDate

      If ($PCModuleDate -eq $CurrentModuleDate) { 
         Write-Host ("Latest version of {0} is installed on this PC - no need to update" -f $Module) 
      } Else {
         Write-Host ("Updating {0} module to version {1}" -f $Module, $CurrentVersion) -foregroundcolor Yellow
         Remove-Module $Module -ErrorAction SilentlyContinue
         Update-Module $Module -Force -Confirm:$False -Scope AllUsers
         $UpdatedModules++
      } # End if 
    }
} # End ForEach Module

# Check and remove older versions of the modules from the PC
Write-Host "Beginning clean-up phase..."
[array]$SetofInstalledModules = Get-InstalledModule
[array]$GraphModules = $SetOfInstalledModules | Where-Object {$_.Name -Like "*Microsoft.Graph*"} | Select-Object -ExpandProperty Name
$ModulesToProcess = $O365Modules + $GraphModules | Sort-Object -Unique

ForEach ($Module in $ModulesToProcess) {
   Write-Host "Checking for older versions of" $Module
   [array]$AllVersions = Get-InstalledModule -Name $Module -AllVersions -ErrorAction SilentlyContinue
   If ($AllVersions) {
     $AllVersions = $AllVersions | Sort-Object PublishedDate -Descending 
     $MostRecentVersion = $AllVersions[0].Version
     If ($MostRecentVersion -isnot [string]) { # Handle PowerShell 5 - PowerShell 7 returns a string
        $MostRecentVersion = $MostRecentVersion.Major.toString() + "." + $MostRecentVersion.Minor.toString() + "." + $MostRecentVersion.Build.toString()
     }
     [datetime]$MostRecentVersionDate = $AllVersions[0].PublishedDate
     $PublishedDate = (Get-Date($MostRecentVersionDate) -format g)
     Write-Host ("Most recent version of {0} is {1} published on {2}" -f $Module, $MostRecentVersion, $PublishedDate)
     If ($AllVersions.Count -gt 1 ) { # More than a single version installed
      ForEach ($Version in $AllVersions) { #Check each version and remove old versions
        [datetime]$VersionDate = $Version.PublishedDate
        If ($VersionDate -lt $MostRecentVersionDate)  { # Old version - remove
           Write-Host ("Uninstalling version {0} of module {1}" -f $Version.Version, $Module) -foregroundcolor Red 
           Uninstall-Module -Name $Module -RequiredVersion $Version.Version -Force
           $RemovedModules++
         } #End if version check
      } # End ForEach versions 
    } Else {
        Write-Host ("No earlier versions of {0} module to remove" -f $Module)
    } # End check for more than one version
  } #End If
} #End ForEach

Write-Host ("Installed modules: {0} Updated modules: {1}  Removed old versions of modules: {2}" -f $InstalledModules, $UpdatedModules, $RemovedModules)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
