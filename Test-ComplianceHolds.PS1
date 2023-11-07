# Test-ComplianceHolds.PS1
# Test the Invoke-HoldRemovalAction cmdlet 
# Documentation: https://learn.microsoft.com/en-us/powershell/module/exchange/invoke-holdremovalaction?view=exchange-ps&WT.mc_id=M365-MVP-9501
# This script: https://github.com/12Knocksinna/Office365itpros/blob/master/Test-ComplianceHolds.PS1
# 1-Nov-2023

# Connect to Exchange Online and then to the Compliance endpoint
Connect-ExchangeOnline -SkipLoadingCmdletHelp
Connect-IPPSSession -CommandName Invoke-HoldRemovalAction

$CSVFileEXO = "c:\temp\EXOMailboxHolds.CSV"
$CSVFileSPO = "c:\temp\SPOSiteHolds.CSV"

Write-Host "Looking for mailboxes to process..."
# Fetch user mailboxes
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -Properties LitigationHoldEnabled | `
        Sort-Object DisplayName
If (!($Mbx)) {
    Write-Host "Unable to find mailboxes - exiting"; break
}

# Loop through mailboxes and figure out what kind of hold information is reported
[int]$i = 0
$MbxReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
    $i++
    $HoldOutput = $Null
    Write-Host ("Analyzing holds on mailbox {0} ({1}/{2})" -f $M.displayName, $i, $Mbx.count)
    [array]$HoldData = Invoke-HoldRemovalAction -Action GetHolds -ExchangeLocation $M.ExternalDirectoryObjectId
    # Throw away any blank entries
    $HoldData = $HoldData | Where-Object {$_.length -gt 0}
    If (!($HoldData)) {
        $HoldInformation = "No holds on mailbox"
    } Else {
        [array]$HoldInformation = $Null
        ForEach ($Hold in $HoldData) {
            If ($Hold -eq 'DelayHold') {
                Write-Output "Delay hold set"
                $HoldOutput += "Delay Hold set"
            }
            # Litigation hold
            If ($Hold -eq '98E9BABD09A04bcf8455A58C2AA74182Unlimit') {
                $HoldInformation += "Litigation Hold set"
            } 
            # Hold imposed by eDiscovery case
            If ($Hold.SubString(0,4) -eq 'UniH') {
                $eDiscoveryHoldId = $Hold.SubString(4, $Hold.Length-4)
                $eDiscoveryHoldName = $Null
                $eDiscoveryHoldName = (Get-CaseHoldPolicy -Identity $eDiscoveryHoldId -ErrorAction SilentlyContinue).Name 
                If ($eDiscoveryHoldName) {
                    $HoldInformation += ("eDiscoveryHold: {0}" -f $eDiscoveryHoldName)
                } Else {
                    Write-Output ("Unable to identify hold identifier {0}" -f $Hold) 
                    $HoldInformation += ("Unidentified eDiscoveryId: {0}" -f $Hold)
                }
            }
            # Anything else
            } If (($Hold -ne 'DelayHold') -and ($Hold -ne '98E9BABD09A04bcf8455A58C2AA74182Unlimit') `
                -and ($Hold.SubString(0,4) -ne 'UniH')) {
                Write-Host "Legacy hold set"
                $HoldInformation += $Hold
            }
    }
    If ($HoldInformation) {
        $HoldOutput = $HoldInformation -Join ", "
    }
    $Reportline = [PsCustomObject]@{
        User    = $M.DisplayName
        UPN     = $M.UserPrincipalName
        Holds   = $HoldOutput
    }
    $MbxReport.Add($ReportLine)
}
$MbxReport | Export-CSV -NoTypeInformation $CSVFileEXO
Write-Host ("Exchange mailbox legacy hold information available in {0}" -f $CSVFileEXO)

# SharePoint
Write-Host "Finished dealing with Exchange Online - Now moving to SharePoint Online"
Import-Module Microsoft.Online.Sharepoint.PowerShell
Connect-SPOService -url https://office365itpros-admin.sharepoint.com  -Credential (Get-Credential)

[array]$Sites = Get-SpoSite -Limit All -Template "Group#0"
$SPOReport = [System.Collections.Generic.List[Object]]::new()
[Int]$i = 0
ForEach ($Site in $Sites) {
  $i++
  Write-Host ("Processing site {0} {1}/{2}" -f $Site.Title, $i, $Sites.Count)
  [array]$HoldData = Invoke-HoldRemovalAction -Action GetHolds -SharePointLocation $Site.Url
  If ($HoldData) {
     Write-Host ("Found holds on site {0} {1}" -f $Site.Title, $Site.Url) -ForegroundColor Red
     ForEach ($Hold in $HoldData) {
       $HoldInfo = Get-RetentionCompliancePolicy -Identity $Hold
       If ($HoldInfo.IsAdaptivePolicy -eq $True) {
            $AdaptiveFlag = "Adaptive Policy"
       } Else {
            $AdaptiveFlag = "Static Policy"
       }
       Write-Host ("Microsoft 365 Retention Policy set is {0} ({1})" -f $HoldInfo.Name, $AdaptiveFlag) -ForegroundColor Yellow
       $Reportline = [PsCustomObject]@{
            Site            = $Site.Url
            Title           = $Site.Title
            HoldId          = $Hold
            Hold            = $HoldInfo.name
            'Policy Scope'  = $AdaptiveFlag
       }
       $SPOReport.Add($ReportLine)
    }
  }
}

$SPOReport | Export-CSV -NoTypeInformation $CSVFileSPO
Write-Host ("SharePoint Online site hold information available in {0}" -f $CSVFileSPO)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
