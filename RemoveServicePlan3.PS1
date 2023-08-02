# RemoveServicePlan3.PS1
# Remove an individual service plan from a SKU assigned to Microsoft 365 accounts
# Version of https://github.com/12Knocksinna/Office365itpros/blob/master/RemoveServicePlan.PS1 using Microsoft Graph PowerShell SDK cmdlets instead of MSOL cmdlets 
# to remove service plans from licenses.
# https://github.com/12Knocksinna/Office365itpros/blob/master/RemoveServicePlan3.PS1

Function Get-Response ([string]$Prompt,[int]$NumberPossibleAnswers) {
# Helper function to prompt a question and get a response
   $OKtoProceed = $False
   While ($OKToProceed -eq $False) {
     [int]$Answer = Read-Host $Prompt
     If ($Answer -gt 0 -and $Answer -le $NumberPossibleAnswers) {
      $OKtoProceed = $True
      Return ($Answer) }
     ElseIf ($Answer -eq 0) { #break out of loop
       $OKtoProceed = $True
       Return ($Answer)}
   } #End while
}

Connect-MgGraph -Scopes Directory.ReadWrite.All

$CSVOutputFile = "c:\temp\ServicePlanRemovals.csv"

# Find the set of SKUs used in the tenant
[array]$Skus = (Get-MgSubscribedSku)
Write-Host " "
Write-Host "Which Office 365 product do you want to remove a service plan from?"; [int]$i=0
   ForEach ($Sku in $Skus) {
      $i++
      Write-Host $i ":" $Sku.SkuPartNumber }
   [Int]$Answer = Get-Response -Prompt  "Enter the number of the product to edit" -NumberPossibleAnswers $i
   If (($Answer -gt 0) -and ($Answer -le $i)) {
       $i = ($Answer-1)
       [string]$SelectedSku = $Skus[$i].SkuPartNumber
       [string]$SelectedSkuId = $Skus[$i].SkuId
       Write-Host "OK. Selected product is" $SelectedSku
       $ServicePlans = $Skus[$i].ServicePlans | Select-Object ServicePlanName, ServicePlanId | Sort-Object ServicePlanName
    } #end if
    Elseif ($Answer -eq 0) { #Abort
       Write-Host "Script stopping..." ; break }

# Select Service plan to remove
Write-Host " "
Write-Host "Which Service plan do you want to remove from" $SelectedSku; [int]$i=0
   ForEach ($ServicePlan in $ServicePlans) {
      $i++
      Write-Host $i ":" $ServicePlan.ServicePlanName }
   [Int]$Answer = Get-Response -Prompt "Enter the number of the service plan to remove" -NumberPossibleAnswers $i
   If (($Answer -gt 0) -and ($Answer -le $i)) {
      [int]$i = ($Answer-1)
      [string]$ServicePlanId = $ServicePlans[$i].ServicePlanId
      [string]$ServicePlanName = $ServicePlans[$i].ServicePlanName
      Write-Host " "
      Write-Host ("Proceeding to remove service plan {0} from the {1} license for target users." -f $ServicePlanName, $SelectedSku)
    } #end If
       Elseif ($Answer -eq 0) { #Abort
       Write-Host "Script stopping..." ; break }

# 
Write-Host "Searching for licensed user accounts"
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All `
  -Property Id, userPrincipalName, displayName, assignedLicenses | Sort-Object DisplayName

Write-Host ("Total of {0} licensed user accounts found" -f $Users.count) -Foregroundcolor red

# Main loop through mailboxes to remove selected service plan from a SKU if the SKU is assigned to the account.
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
     Write-Host "Checking licenses for" $User.DisplayName
     If ($SelectedSkuId -in $User.AssignedLicenses.SkuId) { # User has selected SKU
     # Fetch enabled service plans
     [array]$AllLicenses = Get-MgUserLicenseDetail -UserId $User.Id | Select-Object -ExpandProperty ServicePlans | Sort-Object ServicePlanId -Unique
     [array]$Licenses = $AllLicenses | Where-Object {$_.ProvisioningStatus -eq 'Success'}
     [array]$DisabledLicenses = $AllLicenses | Where-Object {$_.ProvisioningStatus -eq 'Disabled'}
     # Figure out if any service plans are already disabled and add to the set to update
     [array]$DisabledSPs = $ServicePlanId
     If ($DisabledLicenses) {
        If ($DisabledLicenses.Count -eq 1) {
          $DisabledSPs += $DisabledLicenses.ServicePlanId 
      } Else {
          ForEach ($SP in $DisabledLicenses) {
            $DisabledSPs += $SP.ServicePlanId } 
        }
     } # End if
     
     If ($ServicePlanId -in $Licenses.ServicePlanId) {
        Write-Host ("Removing service plan {0} from SKU {1} for account {2}" -f $ServicePlanName, $SelectedSKUId, $User.DisplayName) -foregroundcolor Red
        $LicenseOptions = @{SkuId = $SelectedSkuId ; DisabledPlans = $DisabledSPs } 
        Try {
           $Status = Set-MgUserLicense -UserId $User.Id -AddLicenses $LicenseOptions -RemoveLicenses @() }
        Catch {
           Write-Output "Something bad happened:" $Status
        }
        $LicenseUpdateMsg = ("Service plan {0} removed from account {1} on {2} from {3}" -f  $ServicePlanName, $User.DisplayName, (Get-Date), $SelectedSKU)
         Write-Host ("Service plan {0} removed from SKU {1} for {2}" -f $ServicePlanName, $SelectedSku, $User.DisplayName)
        $ReportLine = [PSCustomObject][Ordered]@{    
           DisplayName     = $User.DisplayName    
           UPN             = $User.UserPrincipalName
           Info            = $LicenseUpdateMsg
           SKU             = $SelectedSKUId
           "Service Plan"  = $ServicePlanName
           "ServicePlanId" = $ServicePlanId }
        $Report.Add($ReportLine)
        } # End if Service Plan
   } # End if License
} #End Foreach User

Write-Host ("Total Licenses Removed: {0}. Output CSV file available in {1}" -f $Report.Count, $CSVOutputFile) 
# Output the report
$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation $CSVOutputFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
