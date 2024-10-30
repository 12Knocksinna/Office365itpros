

$SkuDataPath = "C:\temp\SkuDataComplete.csv"
$ServicePlanPath = "C:\temp\ServicePlanDataComplete.csv"

If ((Test-Path $skuDataPath) -eq $False) {
    Write-Host ("Can't find the product data file ({0}). Exiting..." -f $skuDataPath) ; break 
}
If ((Test-Path $servicePlanPath) -eq $False) {
    Write-Host ("Can't find the serivice plan data file ({0}). Exiting..." -f $servicePlanPath) ; break 
}
   
$ImportSkus = Import-CSV $skuDataPath
$ImportServicePlans = Import-CSV $servicePlanPath
$SkuHashTable = @{}
ForEach ($Line in $ImportSkus) { $SkuHashTable.Add([string]$Line.SkuId, [string]$Line.DisplayName) }
$ServicePlanHashTable = @{}
ForEach ($Line2 in $ImportServicePlans) { $ServicePlanHashTable.Add([string]$Line2.ServicePlanId, [string]$Line2.ServicePlanDisplayName) }

Get-MgUser -Userid Lotte.Vetler@Office365itpros.com -Property Id, displayname, assignedLicenses, licenseAssignmentAtates | `
  Select-Object -ExpandProperty licenseassignmentstates | `
  Where-Object {$_.Error -eq 'None' -and $_.State -eq 'Active'} | `
  Format-Table AssignedByGroup, @{expression={$SkuHashTable[$_.SkuId]};label="Product"}, LastUpdatedDateTime -AutoSize


$groups = Get-MgGroup -All -Property LicenseProcessingState, DisplayName, Id, AssignedLicenses | Select-Object  displayname, Id, LicenseProcessingState, AssignedLicenses | Select-Object DisplayName, Id, AssignedLicenses -ExpandProperty LicenseProcessingState | Select-Object DisplayName, State, Id, AssignedLicenses | Where-Object {$_.State -eq "ProcessingComplete"}
$groupInfoArray = @()
# Filter the groups to only include those that have licenses assigned
$groups = $groups | Where-Object {$_.AssignedLicenses -ne $null}
# For each group, get the group name, license types, total user count, licensed user count, and license error count
foreach ($group in $groups) {
    $groupInfo = New-Object PSObject
    $groupInfo | Add-Member -MemberType NoteProperty -Name "Group Name" -Value $group.DisplayName
    $groupInfo | Add-Member -MemberType NoteProperty -Name "Group ID" -Value $group.Id
    $groupInfo | Add-Member -MemberType NoteProperty -Name "License Types" -Value ($group.AssignedLicenses | Select-Object -ExpandProperty SkuId)
    $groupInfo | Add-Member -MemberType NoteProperty -Name "Total User Count" -Value (Get-MgGroupMember -GroupId $group.Id -All | Measure-Object).Count
    $groupInfo | Add-Member -MemberType NoteProperty -Name "Licensed User Count" -Value (Get-MgGroupMember -GroupId $group.Id -All | Where-Object {$_.      LicenseProcessingState -eq "ProcessingComplete"} | Measure-Object).Count
    $groupInfo | Add-Member -MemberType NoteProperty -Name "License Error Count" -Value (Get-MgGroupMember -GroupId $group.Id -All | Where-Object {$_.LicenseProcessingState -eq "ProcessingFailed"} | Measure-Object).Count
    $groupInfoArray += $groupInfo
}

# Format the output and print it to the console
$groupInfoArray | Format-Table -AutoSize