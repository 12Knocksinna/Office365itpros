# Update-AzureAutomationModules.PS1
# This script updates the Azure Automation modules in an Azure Automation account to the latest versions available in the PowerShell Gallery. 
# It connects to the specified Azure Automation account, checks for updates to the modules, and applies any available updates.

$SubscriptionId = "35429342-a1a5-4427-9e2d-551840f2ad25" # Add your Azure subscription ID here
$ResourceGroupName = "ExoAutomation" # Add the resource group used by the Azure Automation account
$AutomationAccountName = "M365Automation" # Add your Azure Automation account name here
$RuntimeEnvironment = "PS74-Graph" # Specify the runtime environment (e.g., PowerShell7, PowerShell5.1)

[array]$RuntimeModules = "Az.Accounts", "Az.Automation", "ExchangeOnlineManagement", "Microsoft.Online.Sharepoint.PowerShell", `
    "MicrosoftTeams", "Pnp.PowerShell", "ImportExcel", "Microsoft.Graph.Authentication", "Microsoft.Graph.Groups", "Microsoft.Graph.Users",`
     "Microsoft.Graph.Planner", "Microsoft.Graph.Users.Actions", "Microsoft.Graph.Security", "Microsoft.Graph.Teams", "Microsoft.Graph.Sites", `
     "Microsoft.Online.SharePoint.PowerShell", "Microsoft.PowerShell.PSResourceGet", "MSIdentityTools", "Microsoft.Graph.Reports", `
     "Microsoft.Graph.Mail", "Microsoft.Graph.Applications", "Microsoft.Graph.Calendar", "Microsoft.Graph.Files"

Write-Output "Connecting to Azure"
Connect-AzAccount -Subscription $SubscriptionId 

ForEach ($Module in $RuntimeModules) {
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
}

# Check installation
Do {
    Write-Host "Checking Packages..."
    $Data = Invoke-AzRestMethod -Method Get -Path $Path

    $Object = ($Data.Content | convertFrom-Json).Value
    $Array = $Object | Select-Object name, @{Name="Version";Expression={$_.properties.version}},@{Name="provisioningState";Expression={$_.properties.provisioningState}}
    $Packages = $Array | Where-Object {$_.provisioningState -ne "Succeeded" -and $_.provisioningState -ne "Failed"}
    Start-Sleep -Seconds 15
}
While ($Null -ne $Packages)

Write-Output "No packages remain to be installed into the runtime environment"

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Microsoft 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.