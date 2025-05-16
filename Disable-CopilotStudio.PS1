# Disable-CopilotStudio.ps1
# a script to disable Copilot Studio in Microsoft 365 by removing the Copilot Studio service plan from users with Microsoft 365 Copilot licenses
# V1.0 14-May-2025

# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Disable-CopilotStudio.PS1

Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop

[guid]$CopilotSKUId = "639dec6b-bb19-468b-871c-c5c441c4b0cb"
[guid]$CopilotStudioPlanId = "fe6c28b3-d468-44ea-bbd0-a10a5167435c"
[array]$ExcludedUsers = "eff4cd58-1bb8-4899-94de-795f656b4a18"

# Find if any groups assign Copilot licenses and extract the identifiers for their members
[array]$Groups = Get-MgGroup -Filter "assignedLicenses/any(s:s/skuId eq $CopilotSkuId)" -All -Property Id, DisplayName, AssignedLicenses
ForEach ($Group in $Groups) {
    Write-Host "Copilot license found in group" $Group.DisplayName
    [array]$Members = Get-MgGroupMember -GroupId $Group.Id -All -PageSize 500 | Select-Object -ExpandProperty Id
    $ExcludedUsers += $Members
}
 
# Remove any duplicates
$ExcludedUsers = $ExcludedUsers | Sort-Object -Unique

Write-Host "Scanning for user accounts with Microsoft 365 Copilot licenses..."
[array]$Users = Get-MgUser -Filter "assignedLicenses/any(s:s/skuId eq $CopilotSkuId)" -All -Sort 'displayName' -Property Id, displayName, signInActivity, userPrincipalName -PageSize 500
If ($Users) {
    Write-Host "Processing user accounts to remove access to Copilot Studio..."
    ForEach ($User in $Users) {
   
        If ($User.Id -notin $ExcludedUsers) {
            Write-Host "Checking Copilot Studio license status for" $User.DisplayName
            [array]$UserPlans = Get-MgUserLicenseDetail -UserId $User.Id | Select-Object -ExpandProperty ServicePlans
            $Status = ($UserPlans | Where-Object {$_.ServicePlanId -eq $CopilotStudioPlanId} | Select-Object -ExpandProperty ProvisioningStatus )
            If ($Status -eq "Success") {
                Try {
                    Write-Host "Removing Copilot Studio from" $User.DisplayName -ForegroundColor Green
                    Set-MgUserLicense -UserId $User.Id -AddLicenses @{SkuId = $CopilotSkuID; DisabledPlans = $CopilotStudioPlanId } `
                        -RemoveLicenses @() -ErrorAction Stop | Out-Null
                } Catch {
                    Write-Host "Couldn't remove Copilot Studio from" $User.DisplayName -ForegroundColor Red
                }
            }
        } Else {
            Write-Host "Skipping excluded user" $User.DisplayName -ForegroundColor Yellow
        }
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
