# RemoveServicePlan.PS1
# A script to show how to remove a single service plan from the licenses assigned to a Microsoft 365 account.
# https://github.com/12Knocksinna/Office365itpros/blob/master/RemoveServicePlan.PS1
# Note: The AzureAD and MSOL modules were deprecated by Microsoft on 30-March-2024

# Check loaded modules
$ModulesLoaded = Get-Module | Select-Object Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
If (!($ModulesLoaded -match "MsOnline")) {Write-Host "Please connect to the Microsoft Online Services module and then restart the script"; break}
If (!($ModulesLoaded -like "*AzureAD*")) {Write-Host "Please connect to the Azure Active Directory module and then restart the script"; break}
# We seem to be fully connected to the necessary modules so we can proceed

Function Get-Response ([string]$Prompt,[int]$NumberPossibleAnswers) {
# Help function to prompt a question and get a response
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
$CSVOutputFile = "c:\temp\ServicePlanRemovals.csv"

Write-Host "Retrieving service domain for the tenant"
[string]$TenantName = $Null
$DomainInfo = Get-AzureADTenantDetail
[array]$Domains = Get-AcceptedDomain | Select-Object DomainName
[array]$ServiceDomains = $Domains | Where-Object {$_.DomainName -like "*onmicrosoft.com*"}
Switch ($ServiceDomains.Count) {
   0  { Write-Host "Can't locate the service domain" }
   1  { $TenantName = $ServiceDomains.DomainName.SubString(0,$ServiceDomains.DomainName.IndexOf(".")) + ":" }
   2  { Write-Host "Multiple service domains found"
        [int]$i = 0
        ForEach ($ServiceDomain in $ServiceDomains) {
           $i++
           Write-Host $i ":" $ServiceDomain.DomainName
        }
        [Int]$Answer = Get-Response -Prompt "Enter the number of the service domain to use" -NumberPossibleAnswers $i
        If ($Answer -gt 0) { 
            $i = ($Answer-1) }
        Elseif ($Answer -eq 0) { #Abort
            Write-Host "Script stopping..." ; break }
        $TenantName = $ServiceDomains[$i].DomainName.SubString(0,$ServiceDomains[$i].DomainName.IndexOf(".")) + ":"
        Write-Host ("Selected service domain for the {0} tenant is {1}" -f ($DomainInfo.DisplayName), $ServiceDomains[$i].DomainName)
    } # End multiple service domains
} # End switch

# And exit if we don't have a good service domain to use
If (!($TenantName)) { Write-Host "Exiting..." ; break}

# Find the set of SKUs used in the tenant
[array]$Skus = (Get-AzureADSubscribedSku)
Write-Host " "
Write-Host "Which Office 365 product do you want to remove a service plan from?"; [int]$i=0
   ForEach ($Sku in $Skus) {
      $i++
      Write-Host $i ":" $Sku.SkuPartNumber }
   [Int]$Answer = Get-Response -Prompt  "Enter the number of the product to edit" -NumberPossibleAnswers $i
   If (($Answer -gt 0) -and ($Answer -le $i)) {
       $i = ($Answer-1)
       [string]$SelectedSkuId = $Skus[$i].SkuPartNumber
       Write-Host "OK. Selected product is" $SelectedSkuId
       $ServicePlans = $Skus[$i].ServicePlans | Select-Object ServicePlanName, ServicePlanId | Sort-Object ServicePlanName
    }  Elseif ($Answer -eq 0) { #Abort
       Write-Host "Script stopping..." ; break 
  }

# Select Service plan to remove
Write-Host " "
Write-Host "Which Service plan do you want to remove from" $SkuId; [int]$i=0
   ForEach ($ServicePlan in $ServicePlans) {
      $i++
      Write-Host $i ":" $ServicePlan.ServicePlanName }
   [Int]$Answer = Get-Response -Prompt "Enter the number of the service plan to remove" -NumberPossibleAnswers $i
   If (($Answer -gt 0) -and ($Answer -le $i)) {
      [int]$i = ($Answer-1)
      [string]$ServicePlanId = $ServicePlans[$i].ServicePlanId
      [string]$ServicePlanName = $ServicePlans[$i].ServicePlanName
      Write-Host " "
      Write-Host ("Proceeding to remove service plan {0} from the {1} license for target users." -f $ServicePlanName, $SelectedSKUId)
    } #end If
       Elseif ($Answer -eq 0) { #Abort
       Write-Host "Script stopping..." ; break }

# We need  to know what target accounts to remove the service plan from. In this case, we use Get-ExoMailbox to find a bunch of user mailboxes, mostly because we can use a server-side
# filter. You can use whatever other technique to find target accounts (like Get-AzureADUser). The important thing is to feed the object identifier for the account to Get-MsolUser to 
# retrieve license information
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Select-Object DisplayName, UserPrincipalName, Alias, ExternalDirectoryObjectId
[int]$LicensesRemoved = 0
Write-Host ("Total of {0} matching mailboxes found" -f $mbx.count) -Foregroundcolor red

# Main loop through mailboxes to remove selected service plan from a SKU if the SKU is assigned to the account.
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
        Write-Host "Checking licenses for" $M.DisplayName
        $MsoUser = (Get-MsolUser -ObjectId $M.ExternalDirectoryObjectId)
        $i = 0   
        Foreach ($SKU in $MsoUser.Licenses.Accountsku.SKUPartNumber) {
        If ($SKU -eq $SelectedSkuId)
          {
            # Set up license options to remove the selected service plan ($ServicePlanName) from the SKU ($SelectedSkuId)
            $FullLicenseName = $TenantName + $SelectedSkuId # Combination of tenant name and SKU name
            $RemoveServicePlan = (New-MsolLicenseOptions -AccountSkuId $FullLicenseName -DisabledPlans $ServicePlanName )
            Write-Host ("Removing service plan {0} from SKU {1} for account {2}" -f $ServicePlanName, $SelectedSKUId, $M.DisplayName) -foregroundcolor Red
            Set-MsolUserLicense -ObjectId $M.ExternalDirectoryObjectId -LicenseOptions $RemoveServicePlan
            $LicenseUpdateMsg = $ServicePlanName + " service plan removed from account " + $M.UserPrincipalName + " on " + (Get-Date) + " from " + $FullLicenseName
            Set-Mailbox -Identity $M.Alias -ExtensionCustomAttribute2 $LicenseUpdateMsg
            Write-Host ("Service plan {0} removed from SKU {1} for {2}" -f $ServicePlanName, $SelectedSkuID, $M.DisplayName)
            $LicensesRemoved++
            $ReportLine = [PSCustomObject][Ordered]@{    
               DisplayName     = $M.DisplayName    
               UPN             = $M.UserPrincipalName
               Info            = $LicenseUpdateMsg
               SKU             = $SelectedSKUId
               "Service Plan"  = $ServicePlanName
               "ServicePlanId" = $ServicePlanId }
            $Report.Add($ReportLine)
          } # End if
        } # End ForEach
}
Write-Host ("Total Licenses Removed: {0}. Output CSV file available in {1}" -f $LicensesRemoved, $CSVOutputFile) 
# Output the report
$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation $CSVOutputFile


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
