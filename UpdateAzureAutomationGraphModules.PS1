# UpdateAzureAutomationGraphModules.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/Update%20AzureAutomationGraphModules.PS1
# A script to update the set of Graph modules for an Azure Automation account
# V3.0 16-Dec-2022 - Updated to include Exchange Online management module
# Requires the Az.Automation PowerShell module

Write-Host "Connecting to Azure Automation"
# If your account uses MFA, as it should... you need to authenticate by passing the tenantid and subscriptionid 
# see https://learn.microsoft.com/en-us/powershell/module/az.accounts/connect-azaccount?view=azps-11.3.0 for more info
# e.g. 
# $SubscriptionId = "35429342-a1a5-4427-9e2d-551840f2ad25"
# $TenantId = "a662313f-14fc-43a2-9a7a-d2e27f4f3478"
# Connect-AzAccount -TenantId $TenantId -SubscriptionId $SubscriptionId
$Status = Connect-AzAccount
If (!($Status)) { 
  Write-Host "Account not authenticated - exiting" ; break 
}

# Find Latest version from PowerShell Gallery
$DesiredVersion = (Find-Module -Name Microsoft.Graph | Select-Object -ExpandProperty Version)
If ($DesiredVersion -isnot [string]) { # Handle PowerShell 5 - PowerShell 7 returns a string
   $DesiredVersion = $DesiredVersion.Major.toString() + "." + $DesiredVersion.Minor.toString() + "." + $DesiredVersion.Build.toString()
}
Write-Host ("Checking for version {0} of the Microsoft.Graph PowerShell module" -f $DesiredVersion)
# Process Exchange Online also...
$DesiredExoVersion = (Find-Module -Name ExchangeOnlineManagement | Select-Object -ExpandProperty Version)
If ($DesiredExoVersion -isnot [string]) { # Handle PowerShell 5 - PowerShell 7 returns a string
    $DesiredExoVersion = $DesiredExoVersion.Major.toString() + "." + $DesiredExoVersion.Minor.toString() + "." + $DesiredExoVersion.Build.toString()
}
Write-Host ("Checking for version {0} of the Exchange Online Management module" -f $DesiredExoVersion)

[Array]$AzAccounts = Get-AzAutomationAccount
If (!($AzAccounts)) { write-Host "No Automation accounts found - existing" ; break }
Write-Host ("{0} Azure Automation accounts will be processed" -f $AzAccounts.Count)

ForEach ($AzAccount in $AzAccounts) {
  $AzName = $AzAccount.AutomationAccountName
  $AzResourceGroup = $AzAccount.ResourceGroupName
  Write-Host ("Checking Microsoft Graph Modules in Account {0}" -f $AzName)

  [array]$GraphPSModules = Get-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup |  Where-Object {$_.Name -match "Microsoft.Graph"}
  If ($GraphPSModules.count -gt 0) {
    Write-Host ""
    Write-Host "Current Status"
    Write-Host "--------------"
    $GraphPSModules | Format-Table Name, Version, LastModifiedTime }
  
  $UpgradeNeeded = $True
  $ModulesToUpdate = $GraphPSModules | Where-Object {$_.Version -ne $DesiredVersion}
  $ModulesToUpdate = $ModulesToUpdate | Sort-Object Name
  If ($ModulesToUpdate.Count -eq 0) {
     Write-Host ("No modules need to be updated for account {0}" -f $AzName)
     Write-Host ""
     $UpgradeNeeded = $False
  } Else {
    Write-Host ""
    Write-Host ("Modules that need to be updated to {0}" -f $DesiredVersion)
    Write-Host ""
    $ModulesToUpdate | Format-Table Name, Version, LastModifiedTime
    Write-Host "Removing old modules..."
    ForEach ($Module in $ModulesToUpdate) {
       $ModuleName = $Module.Name
       Write-Host ("Uninstalling module {0} from Az Account {1}" -f $ModuleName, $AzName)
       Remove-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -Confirm:$False -Force }
   }

# Check if Modules to be updated contain Microsoft.Graph.Authentication. It should be done first to avoid dependency issues
 If ($ModulesToUpdate.Name -contains "Microsoft.Graph.Authentication" -and $UpgradeNeeded -eq $True) { 
   Write-Host ""
   Write-Host "Updating Microsoft Graph Authentication module first"
   $ModuleName = "Microsoft.Graph.Authentication"
   $Uri = "https://www.powershellgallery.com/api/v2/package/$ModuleName/$DesiredVersion"
   $Status = New-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -ContentLinkUri $Uri 
   Start-Sleep -Seconds 180 
   # Remove authentication from the set of modules for update
   [array]$ModulesToUpdate = $ModulesToUpdate | Where-Object {$_.Name -ne "Microsoft.Graph.Authentication"}
 }

# Only process remaining modules if there are any to update
If ($ModulesToUpdate.Count -gt 0 -and $UpgradeNeeded -eq $True) {
  Write-Host "Adding new version of modules..."
  ForEach ($Module in $ModulesToUpdate) { 
    [string]$ModuleName = $Module.Name
    $Uri = "https://www.powershellgallery.com/api/v2/package/$ModuleName/$DesiredVersion"
    Write-Host ("Updating module {0} from {1}" -f $ModuleName, $Uri)
    $Status = (New-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -ContentLinkUri $Uri)
  } #End ForEach
  Write-Host "Waiting for module import processing to complete..."
  # Wait for to let everything finish
  [int]$x = 0
  Do  {
    Start-Sleep -Seconds 60
    # Check that all the modules we're interested in are fully provisioned with updated code
    [array]$GraphPSModules = Get-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup | `
       Where-Object {$_.Name -match "Microsoft.Graph" -and $_.ProvisioningState -eq "Succeeded"}
    [array]$ModulesToUpdate = $GraphPSModules | Where-Object {$_.Version -ne $DesiredVersion}
    If ($ModulesToUpdate.Count -eq 0) {
      $x = 1
    } Else {
      Write-Host "Still working..." 
    }
  } While ($x = 0)

  Write-Host ""
  Write-Host ("Microsoft Graph modules are now upgraded to version {0} for AZ account {1}" -f $DesiredVersion, $AzName)
  Write-Host ""
  $GraphPSModules | Format-Table Name, Version, LastModifiedTime
 } # End If Modules

 # Check for updates to the Exchange Online Management module
 Write-Host ("Checking Exchange Online Management module in Account {0}" -f $AzName)
 [array]$ExoPSModule = Get-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup | Where-Object {$_.Name -match "ExchangeOnlineManagement" }
  If ($ExoPSModule) {
    Write-Host ""
    Write-Host "Current Status"
    Write-Host "--------------"
    $ExoPSModule | Format-Table Name, Version, LastModifiedTime }
  
  $UpgradeNeeded = $True
  [array]$ModulesToUpdate = $ExoPSModule | Where-Object {$_.Version -ne $DesiredExoVersion}
  If (!($ModulesToUpdate)) {
     Write-Host ("The Exchange Online Management module does not need to be updated for account {0}" -f $AzName)
     Write-Host ""
     $UpgradeNeeded = $False
  } Else {
    [string]$ModuleName = "ExchangeOnlineManagement"
    Write-Host ""
    Write-Host ("Updating the Exchange Online management module to version {0}" -f $DesiredExoVersion)
    Write-Host "Removing old module..."
    Write-Host ("Uninstalling module {0} from Az Account {1}" -f $ModuleName, $AzName)
    Remove-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -Confirm:$False -Force 
    $Uri = "https://www.powershellgallery.com/api/v2/package/$ModuleName/$DesiredExoVersion"
    Write-Host ("Updating module {0} from {1}" -f $ModuleName, $Uri)
    $Status = (New-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -ContentLinkUri $Uri)
   }

} #End ForEach Az Account

Write-Host "All done. The modules in your Azure Automation accounts are now up to date"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
