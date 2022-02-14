# GetPowerAutomateFlows.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/GetPowerAutomateFlows.PS1
# Reports the set of connectors used by Flows in a tenant. The idea being that you might want to know what connectors are in use when
# deciding how to configure a Power Automate DLP policy
# Requires a connection to AzureAD or AzureADPreview
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "AzureAD")) {Write-Host "Please connect to the Azure AD module and then restart the script"; break}
# OK, we seem to be fully connected and ready to go...
 
Write-Host "Finding flows in the tenant"
[array]$Flows = Get-AdminFlow
If (!($Flows)) { Write-Host "No flows found - exiting"; break }
$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Flow in $Flows){
    Write-Host "Processing" $Flow.DisplayName
    try{
        $User = Get-AzureADUser -ObjectId $Flow.CreatedBy.ObjectId
        $DisplayName = $User.DisplayName
        $UPN = $User.UserPrincipalName
    }
    catch{
        $DisplayName = "Unknown user"
        $UPN = $Null
    }
 
    # Retrieve additional details for the Connector Overview
    $FlowDetails = Get-AdminFlow -FlowName $Flow.FlowName -EnvironmentName $Flow.EnvironmentName
    $Environment = Get-AdminPowerAppEnvironment $Flow.EnvironmentName
 
    $ConnectorData = $FlowDetails.Internal.Properties.ConnectionReferences
    $ConnectorNames = [System.Collections.Generic.List[Object]]::new()


    ForEach ($C in $ConnectorData.PSObject.Properties) { $ConnectorNames.Add($C.Value.DisplayName) }
    $ConnectorNames = $ConnectorNames -Join ", "
 
    $FlowDetail = [PSCustomObject][Ordered]@{ 
        Flow        = $Flow.DisplayName
        Creator     = $DisplayName
        UPN         = $UPN
        Connectors  = $ConnectorNames
        FlowId      = $Flow.FlowName
        State       = $Flow.Enabled
        CreatedTime = Get-Date($Flow.CreatedTime) -format g
        Environment = $Environment.DisplayName
    }
    $Report.Add($FlowDetail)
}
 
$Report | Out-GridView
