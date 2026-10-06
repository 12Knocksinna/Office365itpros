# Update-AzureAutomationModules.PS1
# This script updates the Azure Automation modules in an Azure Automation account to the latest versions available in the PowerShell Gallery. 
# It connects to the specified Azure Automation account, checks for updates to the modules, and applies any available updates.

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Update-AzureAutomationModules.PS1
# V1.0 25-Jul-2025
# V1.1 6-Oct-2026   Add output if error occurred when updating module

# Define the identifier of the Azure subscription to use 
$SubscriptionId = "35429342-a1a5-4427-9e2d-551840f2ad25" 
# Define the resource group used by the Azure Automation account
$ResourceGroupName = "ExoAutomation"
# Define the name of the Azure Automation account
$AutomationAccountName = "M365Automation"
# Define the name of the target custom runtime environment to update
$RuntimeEnvironment = "PS74-Graph"
# Define the tenant identifier
$TenantId = 'a662313f-14fc-43a2-9a7a-d2e27f4f3478' 

# Define the list of PowerShell modules to check for updates. This list includes commonly used modules for Microsoft 365 and Azure management.
[array]$Modules = "Az.Accounts", "Az.Automation", "ExchangeOnlineManagement", "Microsoft.Online.Sharepoint.PowerShell", `
    "MicrosoftTeams", "Pnp.PowerShell", "ImportExcel", "Microsoft.Graph.Authentication", "Microsoft.Graph.Groups", "Microsoft.Graph.Users",`
     "Microsoft.Graph.Planner", "Microsoft.Graph.Users.Actions", "Microsoft.Graph.Security", "Microsoft.Graph.Teams", "Microsoft.Graph.Sites", `
     "Microsoft.Online.SharePoint.PowerShell", "Microsoft.PowerShell.PSResourceGet", "MSIdentityTools", "Microsoft.Graph.Reports", `
     "Microsoft.Graph.Mail", "Microsoft.Graph.Applications", "Microsoft.Graph.Calendar", "Microsoft.Graph.Files"

Write-Output "Connecting to Azure"
Connect-AzAccount -Subscription $SubscriptionId -Tenant $TenantId -WarningAction SilentlyContinue

Write-Output ""
Write-Output "Preparing to check PowerShell modules installed in the target runtime environment..."
Write-Output ""

ForEach ($Module in $Modules) {
    Write-Output "Checking for updates to module: $Module"
    $CurrentModule = Find-PsResource -Name $Module
    $Version = $CurrentModule.Version.toString()
    $NupkgURI = "https://www.powershellgallery.com/api/v2/package/$Module/$Version"

$RequestBody = @"
    {
        "properties": {
            "contentLink": { "uri": "$NupkgURI" }
        }
    }
"@
    $Path = ("/subscriptions/{0}/resourceGroups/{1}/providers/Microsoft.Automation/automationAccounts/{2}/runtimeEnvironments/{3}/packages/{4}/?api-version=2023-05-15-preview" -f $SubscriptionId, $ResourceGroupName, $AutomationAccountName, $RuntimeEnvironment, $Module)
  
    $Result = Invoke-AzRestMethod -Method Put -Path $Path -Payload $RequestBody
    # check if everything was OK 
    $Output = $Result.content | ConvertFrom-Json
    $Output.Message
}

# Check installation
Do {
    Write-Output "Checking individual PowerShell modules (packages)..."
    $Data = Invoke-AzRestMethod -Method Get -Path $Path

    $Object = ($Data.Content | convertFrom-Json).Value
    $Array = $Object | Select-Object name, @{Name="Version";Expression={$_.properties.version}},@{Name="provisioningState";Expression={$_.properties.provisioningState}}
    $Packages = $Array | Where-Object {$_.provisioningState -ne "Succeeded" -and $_.provisioningState -ne "Failed"}
    Start-Sleep -Seconds 15
}
While ($Null -ne $Packages)

Write-Output "No packages remain to be installed into the target runtime environment"

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
