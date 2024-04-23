# RemoveServicePlan3.PS1
# Remove an individual service plan from a SKU assigned to Microsoft 365 accounts
# Version of https://github.com/12Knocksinna/Office365itpros/blob/master/RemoveServicePlan.PS1 using Microsoft Graph PowerShell SDK cmdlets instead of MSOL cmdlets 
# to remove service plans from licenses.
# https://github.com/12Knocksinna/Office365itpros/blob/master/RemoveServicePlan3.PS1
# V1.1 8-Apr-2024 1. Made sure that you can only select a service plan that an admin can remove
#                 2. Made the call to Get-MgUser much more efficient by filtering on the SKU assigned to the user
#                 3. Added a check to see if the service plan is already disabled before trying to remove it
#                 4. Made sure that we only report a removed service plan if it's actually removed  
#
# See https://office365itpros.com/2024/04/23/remove-service-plan-powershell-2/ for more information

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

# Directory.ReadWrite.All is used to fetch subscription information for the tenant, read user details, and update user licenses
Connect-MgGraph -Scopes Directory.ReadWrite.All -NoWelcome

$CSVOutputFile = "c:\temp\ServicePlanRemovals.csv"

# Find the set of SKUs used in the tenant
Write-Host "Checking subscriptions known to the tenant" -Foregroundcolor Yellow
[array]$Skus = (Get-MgSubscribedSku)
Clear-Host
Write-Host " "
Write-Host "Which subscription do you want to remove a service plan from?"; [int]$i=0
ForEach ($Sku in $Skus) {
   $i++
   Write-Host $i ":" $Sku.SkuPartNumber 
}
[Int]$Answer = Get-Response -Prompt  "Enter the number of the subscription to edit" -NumberPossibleAnswers $i
If (($Answer -gt 0) -and ($Answer -le $i)) {
   $i = ($Answer-1)
   [string]$SelectedSku = $Skus[$i].SkuPartNumber
   [string]$SelectedSkuId = $Skus[$i].SkuId
   Write-Host "OK. Selected subscription is" $SelectedSku
   # Find the set of service plans in the selected SKU, excluding those that we can't remove because they are assigned by the tenant
   [array]$ServicePlans = $Skus[$i].ServicePlans | Where-Object {$_.AppliesTo -eq 'User'} | Select-Object ServicePlanName, ServicePlanId | Sort-Object ServicePlanName
} Elseif ($Answer -eq 0) { #Abort
   Write-Host "Script stopping..." ; break 
}
# Check if there are any service plans that can be removed from the SKU
If ($ServicePlans.Count -eq 0) {
   Write-Host "No service plans available to remove from" $SelectedSku
   Break
} 

# Select Service plan to remove
Write-Host " "
Write-Host "Which Service plan do you want to remove from" $SelectedSku; [int]$i=0
ForEach ($ServicePlan in $ServicePlans) {
   $i++
   Write-Host $i ":" $ServicePlan.ServicePlanName 
}
[Int]$Answer = Get-Response -Prompt "Enter the number of the service plan to remove" -NumberPossibleAnswers $i
If (($Answer -gt 0) -and ($Answer -le $i)) {
   [int]$i = ($Answer-1)
   [string]$ServicePlanId = $ServicePlans[$i].ServicePlanId
   [string]$ServicePlanName = $ServicePlans[$i].ServicePlanName
   Write-Host " "
   Write-Host ("Proceeding to remove service plan {0} from the {1} license for target users." -f $ServicePlanName, $SelectedSku)
} Elseif ($Answer -eq 0) { #Abort
   Write-Host "Script stopping..."  
   break 
}

# Find the set of users assigned the selected SKU
[guid]$TargetSku = $SelectedSkuId
Write-Host ("Searching for accounts assigned the {0} license" -f $SelectedSku) -foregroundcolor yellow
[array]$Users = Get-MgUser -ConsistencyLevel Eventual -CountVariable Licenses -All `
  -Sort 'displayName' -Property Id, displayName, userPrincipalName, assignedLicenses `
  -Filter "assignedLicenses/any(s:s/skuId eq $TargetSku)" | Sort-Object DisplayName

If ($Users.Count -eq 0) {
   Write-Host ("No user accounts found with the {0} license" -f $SelectedSku) -foregroundcolor red
   Break
} Else {
   Write-Host ("Total of {0} licensed user accounts found" -f $Users.count) -Foregroundcolor red
}

# Main loop through mailboxes to remove selected service plan from a SKU if the SKU is assigned to the account.
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
   $DisabledSPs = $null
   $OKtoProceed = $true
   Write-Host ("Checking service plans for {0}" -f $User.DisplayName)
   # Fetch the set of service plans for the selected SKU
   [array]$AllLicenses = Get-MgUserLicenseDetail -UserId $User.Id | Where-Object {$_.SkuId -eq $SelectedSkuId} | `
      Select-Object -ExpandProperty ServicePlans | Sort-Object ServicePlanId -Unique
   # Find if any service plans are already disabled
   [array]$DisabledLicenses = $AllLicenses | Where-Object {$_.ProvisioningStatus -eq 'Disabled'}
   # Figure out if any service plans are already disabled and add to the set to update

   If ($ServicePlanId -in $DisabledLicenses.ServicePlanId) {
      Write-Host ("Service plan {0} is already disabled for account {1}" -f $ServicePlanName, $User.DisplayName) -foregroundcolor Yellow
      $OKtoProceed = $false
   }

   [array]$DisabledSPs = $ServicePlanId
   If ($DisabledLicenses) {
      If ($DisabledLicenses.Count -eq 1) {
          $DisabledSPs += $DisabledLicenses.ServicePlanId 
      } Else {
         ForEach ($SP in $DisabledLicenses) {
            $DisabledSPs += $SP.ServicePlanId } 
      }
   } # End if
     
   # Go ahead and remove the service plan from the account if it hasn't already been removed
   If ($ServicePlanId -in $AllLicenses.ServicePlanId -and $OKtoProceed -eq $true) {
      Write-Host ("Removing service plan {0} from SKU {1} for account {2}" -f $ServicePlanName, $SelectedSKUId, $User.DisplayName) -foregroundcolor Red
      $LicenseOptions = @{SkuId = $SelectedSkuId ; DisabledPlans = $DisabledSPs } 
      Try {
         $Status = Set-MgUserLicense -UserId $User.Id -AddLicenses $LicenseOptions -RemoveLicenses @() -ErrorAction Stop
         $SPRemoved = $true
      } Catch {
         $SPRemoved = $false
         Write-Output "Something bad happened:" $Status
      }
      If ($SPRemoved) {
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
      } 
   }
} #End Foreach User

Write-Host ("Total service plans removed from user accounts: {0}. Output CSV file available in {1}" -f $Report.Count, $CSVOutputFile) 
# Output the report
$Report | Out-GridView -Title "Service Plan Removals from User Accounts"
$Report | Export-CSV -NoTypeInformation $CSVOutputFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
