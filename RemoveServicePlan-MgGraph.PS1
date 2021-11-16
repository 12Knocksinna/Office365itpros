# RemoveServicePlan-MgGraph.PS1
# Remove an individual service plan from a SKU assigned to Microsoft 365 accounts using cmdlets from the Microsoft Graph PowerShell SDK
# https://github.com/12Knocksinna/Office365itpros/blob/master/RemoveServicePlan-MgGraph.PS1

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
# Check loaded modules
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}
If (!($ModulesLoaded -like "*Microsoft.Graph*")) {Write-Host "Please connect to the Microsoft Graph PowerShell SDK and then restart the script"; break}
# We seem to be fully connected to the necessary modules so we can proceed

# Make sure we're using the beta endpoint to make sure we can get license information
$Profile = (Get-MgProfile).Name
If ($Profile -ne "beta") { Select-MgProfile Beta }
# Check Scopes
$Scopes = Get-MgContext | Select-Object -Expandproperty Scopes
If ("User.ReadWrite.All" -and "Directory.ReadWrite.All" -notin $Scopes) { Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All" }

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
       $ServicePlans = $Skus[$i].ServicePlans | Select ServicePlanName, ServicePlanId | Sort ServicePlanName
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

# We need to know what target accounts to remove the service plan from. In this case, we use Get-ExoMailbox to find a bunch of user mailboxes, mostly because we can use a server-side
# filter. You can use whatever other technique to find target accounts (like Get-MgUser). The important thing is to have an object identifier for each account to 
# retrieve license information
[array]$Mbx = (Get-ExoMailbox -RecipientTypeDetails UserMailbox -Filter {Office -eq "Dublin"} -ResultSize Unlimited | Select DisplayName, UserPrincipalName, Alias, ExternalDirectoryObjectId)
[int]$LicensesRemoved = 0
Write-Host ("Total of {0} matching mailboxes found" -f $mbx.count) -Foregroundcolor red

# Main loop through mailboxes to remove selected service plan from a SKU if the SKU is assigned to the account.
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
        Write-Host "Checking licenses for" $M.DisplayName
        $User = (Get-MgUser -UserId $M.ExternalDirectoryObjectId)
        $i = 0   
        Foreach ($License in $User.AssignedLicenses) {
          If ($License.SkuId -eq $SelectedSkuId)
          { # We match the service plan to remove
            Write-Host ("Removing service plan {0} from SKU {1} for account {2}" -f $ServicePlanName, $SelectedSKUId, $M.DisplayName) -foregroundcolor Red
            $ExistingDisabledPlans = $Null
            ForEach ($S in $User.AssignedLicenses) { # Check for existing disabled licenses
                If ($S.SkuId -eq $SelectedSkuId) {
                 $ExistingDisabledPlans = $S.DisabledPlans }
            }
            $ExistingDisabledPlans += $ServicePlanId # Add the plan we want to remove to the set of disabled plans
            $LicenseToRemove = New-Object -TypeName Microsoft.Graph.PowerShell.Models.MicrosoftGraphAssignedLicense
            $LicenseToRemove.SkuId = $SelectedSkuId
            $LicenseToRemove.DisabledPlans = $ExistingDisabledPlans
            $Status = Set-MgUserLicense -UserId $User.id -AddLicenses $LicenseToRemove -RemoveLicenses @()  
            $LicenseUpdateMsg = $ServicePlanName + " service plan removed from account " + $M.UserPrincipalName + " on " + (Get-Date) + " from " + $SelectedSku
            Set-Mailbox -Identity $M.Alias -ExtensionCustomAttribute2 $LicenseUpdateMsg
            Write-Host ("Service plan {0} removed from SKU {1} for {2}" -f $ServicePlanName, $SelectedSku, $M.DisplayName)
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
        } # End ForEach license
} #End Foreach mailbox

Write-Host ("Total Licenses Removed: {0}. Output CSV file available in {1}" -f $LicensesRemoved, $CSVOutputFile) 
# Output the report
$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation $CSVOutputFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
