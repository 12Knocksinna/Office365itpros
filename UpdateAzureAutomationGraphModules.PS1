# UpdateAzureAutomationGraphModules.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateAzureAutomationGraphModules.PS1
# A script to update the set of Graph modules for an Azure Automation account
# Requires the Az.Automation PowerShell module

Write-Host "Connecting to Azure Automation"
$Status = Connect-AzAccount
If (!($Status)) { Write-Host "Account not authenticated - exiting" ; break }

$DesiredVersion = "1.14'
[Array]$AzAccounts = Get-AzAutomationAccount
If (!($AzAccounts)) { write-Host "No Automation accounts found - existing" ; break }

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
  $ModulesToUpdate = $GraphPSModules | ? {$_.Version -ne $DesiredVersion}
  $ModulesToUpdate = $ModulesToUpdate | Sort-Object Name
  If ($ModulesToUpdate.Count -eq 0) {
     Write-Host ("No modules need to be updated for account {0}" -f $AzName)
     Write-Host ""
     $UpgradeNeeded = $False
  }
  Else {
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
   Sleep 180 
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
    New-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -ContentLinkUri $Uri
  } #End ForEach
  Write-Host "Waiting for module import processing to complete..."
  # Wait for to let everything finish
  $i = 0
  Do  {
   Sleep 60
   # Check that all the modules we're interested in are fully provisioned with updated code
   [array]$GraphPSModules = Get-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup | `
       Where-Object {$_.Name -match "Microsoft.Graph" -and $_.ProvisioningState -eq "Succeeded"}
   [array]$ModulesToUpdate = $GraphPSModules | ? {$_.Version -ne $DesiredVersion}
   If ($ModulesToUpdate.Count -eq 0) {$x = 1}
   Else {Write-Host "Still working..." }
  } While ($x = 0)

  Write-Host ""
  Write-Host ("All Done - Your Microsoft Graph modules are now upgraded to version {0} for AZ account {1}" -f $DesiredVersion, $AzName)
  Write-Host ""
  $GraphPSModules | Format-Table Name, Version, LastModifiedTime
 } # End If Modules
} #End ForEach Az Account


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
