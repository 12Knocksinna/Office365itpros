# Assign-LicensesViaCSV.PS1
# Script to assign licenses for a chosen SKU to a set of users imported from a CSV
# https://github.com/12Knocksinna/Office365itpros/blob/master/Assign-LicensesViaCSV.PS1

Function Get-Response ([string]$Prompt,[int]$NumberPossibleAnswers) {
# Helper function to prompt a question and get a response
    $OKtoProceed = $False       
    While ($OKToProceed -eq $False) {
     [int]$Answer = Read-Host $Prompt
     If ($Answer -gt 0 -and $Answer -le $NumberPossibleAnswers) {
          $OKtoProceed = $True
          Return ($Answer) 
     } ElseIf ($Answer -eq 0) { #break out of loop
          $OKtoProceed = $True
          Return ($Answer)}
    } #End while
}
    
# Connect to the Graph with permission to update the directory (with licenses)
Connect-MgGraph -Scopes Directory.ReadWrite.All
$InputFile = "c:\temp\Users.csv"
$LicensesAvailable = $True
    
# Find the set of SKUs used in the tenant
[array]$Skus = (Get-MgSubscribedSku)
$SkuList = [System.Collections.Generic.List[Object]]::new()  
ForEach ($Sku in $Skus) {
    $SkuAvailable = ($Sku.PrepaidUnits.Enabled - $Sku.ConsumedUnits)
    $ReportLine = [PSCustomObject]@{
        SkuId         = $Sku.SkuId
        SkuPartNumber = $Sku.SkuPartNumber
        Consumed      = $Sku.ConsumedUnits
        Paid          = $Sku.PrepaidUnits.Enabled
        Available     = $SkuAvailable }
    $SkuList.Add($ReportLine)
}
      
# Remove SKUs with no available licenses
$SkuList = $SkuList | Where-Object {$_.Available -gt 0}
If ($SkuList.count -eq 0) {
    $LicensesAvailable = $False
    Write-Host "No SKUs have avaiilable licenses"
}

If ($LicensesAvailable -eq $True) {
[int]$i = 0    
Write-Host " "
Write-host "Product SKUs with available licenses" -foregroundcolor Red
Write-Host "------------------------------------" -foregroundcolor Red
Write-Host ""
Write-Host "Select the Microsoft 365 product SKU to assign licenses to users; enter 0 to exit"; [int]$i=0
ForEach ($Sku in $SkuList) {
    $i++
    $Line = ("{0}: {1} (available units {2})" -f $i, $Sku.SkuPartNumber, $Sku.Available)
    Write-Host $Line 
  }

[Int]$Answer = Get-Response -Prompt  "Enter the number of the product SKU to assign" -NumberPossibleAnswers $i
If (($Answer -gt 0) -and ($Answer -le $i)) {
    $i = ($Answer-1)
    [string]$SelectedSku = $SkuList[$i].SkuPartNumber
    [string]$SelectedSkuId = $SkuList[$i].SkuId
    Write-Host "OK. The selected product SKU to assign to user accounts is:" $SelectedSku
    } Elseif ($Answer -eq 0) { 
    #Abort
        Write-Host "Script stopping..." ; break 
}
} # End listing of available SKUs

Write-Host ""
Write-Host "Looking for accounts to process..."
Write-Host ""
# Import user accounts to assign licenses to - all that's important here is to establish an array of
# user accounts to process. Instead of reading information from a CSV file, the data could come from 
# running the Get-MgUser cmdlet with a filter to fetch certain accounts
If ($LicensesAvailable -eq $True) {
    [array]$Users = Import-CSV $InputFile
    If (!($Users)) {
        Write-Host "Unable to find any users to process - exiting"; break
    } Else {  # Check that there are sufficient licenses available to assign to the number of accounts to process
        If ($Users.Count -gt $SkuList[$i].Available) {
        Write-Host ("{0} users are to receive licenses but there are only {1} licenses available - exiting." -f $Users.Count,  $SkuListing[$i].Available) 
        $LicensesAvailable = $False
        }
    }    
}

If ($LicensesAvailable -eq $True) {
    $AssignmentReport = [System.Collections.Generic.List[Object]]::new()  
    # Check each user to see if the account exists and if the SKU is already assigned
    Write-Host "Checking user accounts and assigning licenses..."
    $i = 0
    ForEach ($User in $Users) {
        $ErrorMsg = $Null; $i++
        Write-Host ("Processing account {0} {1}/{2}" -f $User.displayName, $i, $Users.count)
        $UserData = Get-MgUser -UserId $User.UPN.Trim() -Property id, assignedLicenses -ErrorAction SilentlyContinue
        If (!($UserData)) {
        # Can't find user account, so flag an error
           $ErrorMsg = ("Error: User account {0} does not exist in Entra ID" -f $User.UPN)
           Write-Host $ErrorMsg
           $ReportLine = [PSCustomObject]@{
              User      = $User.UPN
              Status    = $ErrorMsg
              TimeStamp = (Get-Date -format "dd-MMM-yyyy hh:mm:ss") }
           $AssignmentReport.Add($ReportLine)   
        } Else { 
        # Account is available, so check license and assign it if the account doesn't already have it
            $LicenseData = $UserData | Select-Object -ExpandProperty AssignedLicenses
            If ($SelectedSkuId -in $LicenseData.SkuId) {
                $ErrorMsg = ("Error: License {0} is already assigned to user account {1}" -f $SelectedSku, $User.UPN)
                Write-Host $ErrorMsg
                $ReportLine = [PSCustomObject]@{
                    User      = $User.UPN
                    Status    = $ErrorMsg
                    TimeStamp = (Get-Date -format "dd-MMM-yyyy hh:mm:ss") }
                $AssignmentReport.Add($ReportLine)   
            } Else {
            # Assign the license
                $License = Set-MgUserLicense -UserId $User.UPN -Addlicenses @{SkuId = $SelectedSkuId} `
                  -RemoveLicenses @() -ErrorAction SilentlyContinue
                If ($License) {
                    $StatusMsg = ("Success: Sku {0} successfully assigned to {1}" -f $SelectedSku, $User.UPN)
                    Write-Host $StatusMsg 
                    $ReportLine = [PSCustomObject]@{
                        User      = $User.UPN
                        Status    = $StatusMsg
                        TimeStamp = (Get-Date -format "dd-MMM-yyyy hh:mm:ss") }
                    $AssignmentReport.Add($ReportLine)   
                } Else {
                    $ErrorMsg = ("Error: Some problem stopped us assigning Sku {0} to {1}" -f $SelectedSku, $User.UPN)
                    Write-Host $ErrorMsg -ForegroundColor Yellow  
                    $ReportLine = [PSCustomObject]@{
                        User      = $User.UPN
                        Status    = $ErrorMsg
                        TimeStamp = (Get-Date -format "dd-MMM-yyyy hh:mm:ss") }
                     $AssignmentReport.Add($ReportLine)   
                }
            }   
        }
    }
} # End Processing accounts to assign licenses
[array]$Successes = $AssignmentReport | Where-Object {$_.Status -like "*Success:*"}
[array]$Failures = $AssignmentReport | Where-Object {$_.Status -like "*Error:*"}
$PCentSuccesses = ($Successes.Count/$Users.Count).toString("P")
$PcentFailures = ($Failures.count/$Users.Count).toString("P")
Write-Host ""
Write-Host "All done"
Write-Host ("Results of assigning the {0} product SKU" -f $SelectedSku)
Write-Host ""
Write-Host ("Successful license assignments:  {0} {1}" -f $Successes.Count, $PCentSuccesses )
Write-Host ("Failures in license assignments: {0} {1}" -f $Failures.Count, $PcentFailures )
$AssignmentReport | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
