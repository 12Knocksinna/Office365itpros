# Test-ComplianceHolds.PS1
# Test the Invoke-HoldRemovalAction cmdlet 
# https://learn.microsoft.com/en-us/powershell/module/exchange/invoke-holdremovalaction?view=exchange-ps&WT.mc_id=M365-MVP-9501

# Connect to Exchange Online and then to the Compliance endpoint
Connect-ExchangeOnline
Connect-IPPSSession

# Fetch user mailboxes
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -Properties LitigationHoldEnabled | `
        Sort-Object DisplayName
If (!($Mbx)) {
    Write-Host "Unable to find mailboxes - exiting"; break
}
$MbxReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
    Write-Host ("Analyzing holds on mailbox {0}" -f $M.displayName)
    $HoldData = Invoke-HoldRemovalAction -Action GetHolds -ExchangeLocation $M.ExternalDirectoryObjectId
    $Reportline = [PsCustomObject]@{
        User    = $M.DisplayName
        UPN     = $M.UserPrincipalName
        Holds   = $HoldData
    }
    $MbxReport.Add($ReportLine)

}




# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.