# UpdateSensitivityLabelsForGroups.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateSensitivityLabelsForGroups.PS1
# A script to remap the sensitivity labels assigned to groups so that only labels with container settings are used.
# First Check for connection to compliance endpoint
$TenantLabels = @{}
Try {
    $Labels = Get-Label }
Catch {
    Write-Host "Your PowerShell session must be connected to the Compliance endpoint to fetch label data" ; break}
# Now Populate hash table with label data  
$Labels.ForEach( {
       $TenantLabels.Add([String]$_.ImmutableId, $_.DisplayName) })

# Now you need a connection to the Exchange Online management module

$GroupsWithLabels = Get-UnifiedGroup -ResultSize Unlimited | ? {$_.SensitivityLabel -ne $Null}
If ($GroupsWithLabels) {
   $Report = [System.Collections.Generic.List[Object]]::new() 
   CLS
   Write-Host "Processing" $GroupsWithLabels.Count "groups" 
   ForEach ($Group in $GroupsWithLabels) {
     Switch ($Group.SensitivityLabel.Guid) {
      "2fe7f66d-096a-469e-835f-595532b63560" { $NewLabel = "e42fd42e-7240-4df0-9d8f-d14658bcf7ce" } # Public = General Access
      "27451a5b-5823-4853-bcd4-2204d03ab477" { $NewLabel = "d6cfd185-f31c-4508-ae40-229ff18a9919" } # Internal = Limited Access
      "d179cfc9-43d4-41b6-9ddb-3e1aaf3224c8" { $NewLabel = "d6cfd185-f31c-4508-ae40-229ff18a9919" } # Employee Confidental = Limited Access
      "f3b23fed-2839-4270-9b35-1d634c84b2e9" { $NewLabel = "d6cfd185-f31c-4508-ae40-229ff18a9919" } # Market Sensitive = Limited Access
      "f5b1ba01-59f5-4ba0-b73b-f60e348cdc6e" { $NewLabel = "d6cfd185-f31c-4508-ae40-229ff18a9919" } # Financial Data = Limited Access
      "1b070e6f-4b3c-4534-95c4-08335a5ca610" { $NewLabel = "c99e52c6-f5ff-4050-9313-ca6a3a35710f" } # Confidental = Confidential Access
      "81955691-b8e8-4a81-b7b4-ab32b130bff5" { $NewLabel = "c99e52c6-f5ff-4050-9313-ca6a3a35710f" } # Secret = Confidential Access
      "9ec4cb17-1374-4016-a356-25a7de5e411d" { $NewLabel = "c99e52c6-f5ff-4050-9313-ca6a3a35710f" } # Ultra-Confidentoal = Confidential Access
      "c9001382-2af9-4e06-808b-2080c1a9861f" { $NewLabel = "c99e52c6-f5ff-4050-9313-ca6a3a35710f" } # Sensitive Stuff = Confidential Access
      "e42fd42e-7240-4df0-9d8f-d14658bcf7ce" { $NewLabel = $Null }                                  # Group already assigned General Access
      "c29e68f9-bc4f-413b-a741-6db8e38ad1c6" { $NewLabel = $Null }                                  # Group already assigned Guest Access
      "d6cfd185-f31c-4508-ae40-229ff18a9919" { $NewLabel = $Null }                                  # Group already assigned Limited Access
      "c99e52c6-f5ff-4050-9313-ca6a3a35710f" { $NewLabel = $Null }                                  # Group already assigned Confidential Access
      "default"                              { $NewLabel = "c29e68f9-bc4f-413b-a741-6db8e38ad1c6" } # Anything else = Guest Access
     } #End Switch
    If ($NewLabel -ne $Null) { # We can update with a new sensitivity label
       Write-Host "Updating group:" $Group.DisplayName "Old label:" ($TenantLabels[$Group.SensitivityLabel.Guid]) "New label:" ($TenantLabels[$NewLabel])
       Set-UnifiedGroup -Identity $Group.ExternalDirectoryObjectId -SensitivityLabel $NewLabel 
       $ReportLine  = [PSCustomObject] @{          
          Group    = $Group.DisplayName
          OldLabel = ($TenantLabels[$Group.SensitivityLabel.Guid])
          OldGuid  = $Group.SensitivityLabel.Guid
          NewLabel = ($TenantLabels[$NewLabel])
          NewGuid  = $NewLabel
          Status   = "Updated" }
       $Report.Add($ReportLine) 
       }
    ElseIf ($NewLabel -eq $Null) { # Signal that group already has a container management label 
       Write-Host "Group:" $Group.DisplayName "is already assigned the" ($TenantLabels[$Group.SensitivityLabel.Guid]) "container management label" -foregroundcolor Red 
       $ReportLine  = [PSCustomObject] @{          
         Group    = $Group.DisplayName
         OldLabel = ($TenantLabels[$Group.SensitivityLabel.Guid])
         OldGuid  = $Group.SensitivityLabel.Guid
         NewLabel = "Not changed"
         NewGuid  = "N/A"
         Status   = "Not Updated" }
       $Report.Add($ReportLine) }
  } #End ForEach
} #End if

$Report | Export-CSV -NoTypeInformation c:\temp\GroupLabelUpdates.csv
Write-Host "All done: results written to c:\temp\GroupLabelUpdates.csv"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
