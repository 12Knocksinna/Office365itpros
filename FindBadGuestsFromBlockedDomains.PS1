# FindBadGuestsFromBlockedDomains.PS1
# Script showing how to report any existing guests in Microsoft 365 Groups which come from domains blocked by the Azure B2B Collaboration policy
# V2.0 12 June 2023
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindBadGuestsFromBlockedDomains.PS1

# Check that we're connected to Exchange Online
If ($Null -eq (Get-ConnectionInformation)) {
    Connect-ExchangeOnline
}

Connect-MgGraph -Scopes Directory.Read.All, Policy.Read.All

# Define output file for report
$OutputFile = "C:\Temp\BadGuestAccounts.csv"

# Get Blocked domains from the Azure AD B2B Collaboration settings
$Uri = "https://graph.microsoft.com/beta/legacy/policies/"
[array]$Policies = Invoke-MgGraphRequest -Uri $Uri -Method Get
$B2BPolicy = $Policies.Value | Where-Object {$_.displayName -eq "B2BManagementPolicy"}
$Uri = ("https://graph.microsoft.com/beta/legacy/policies/{0}/definition" -f $B2BPolicy.id)
$Data = Invoke-MgGraphRequest -Uri $Uri -Method Get
$Policy = $Data.Value | ConvertFrom-Json
[array]$BlockedDomains = $Policy.B2BManagementPolicy.InvitationsAllowedAndBlockedDomainsPolicy.BlockedDomains

If ($BlockedDomains -eq $Null) {
   Write-Host "Looks like you have not configured any blocked domains" ; break}
Else {
   Write-Host ("Checking Microsoft 365 Groups for guests from blocked domains {0}" -f ($BlockedDomains -join ", "))  }

[array]$Groups = Get-UnifiedGroup -ResultSize Unlimited -Filter {GroupExternalMemberCount -gt 0}
If (!($Groups)) { Write-Host "No Groups with external guests found"; break }
$Report = [System.Collections.Generic.List[Object]]::new() 
CLS; $GroupNumber = 0
ForEach ($Group in $Groups) {
   $GroupNumber++
   $ProgressBar = "Processing Group " + $Group.DisplayName + " (" + $GroupNumber + " of " + $Groups.Count + ")" 
   Write-Progress -Activity "Checking Microsoft 365 Groups for guest members from blocked domains" -Status $ProgressBar -PercentComplete ($GroupNumber/$Groups.Count*100)
   $Members = Get-UnifiedGroupLinks -Identity $Group.ExternalDirectoryObjectId -LinkType Member | ? {$_.RecipientTypeDetails -eq "GuestMailUser" } | Select ExternalEmailAddress, DisplayName, ExternalDirectoryObjectId
   ForEach ($Guest in $Members) {
      $Domain = $Guest.ExternalEmailAddress.Split("@")[1]
      $GuestEmail = $Guest.ExternalEmailAddress.Split(":")[1]
      If ($BlockedDomains -contains $Domain) {
         Write-Host ("Found guest user {0} ({1}) in group {2}" -f $Guest.DisplayName, $GuestEmail, $Group.DisplayName) -foregroundcolor Red
         $GuestData = Get-MgUser -UserId $Guest.ExternalDirectoryObjectId
         $AccountAge = ($GuestData.CreatedDateTime | New-TimeSpan).Days
         $ReportLine = [PSCustomObject]@{
               Guest           = $Guest.DisplayName
               "Guest Email"   = $GuestEmail
               Group           = $Group.DisplayName
               Created         = $CreationDate
               "Age in Days"   = $AccountAge
               "ObjectId"      = $Guest.ExternalDirectoryObjectId }
            $Report.Add($ReportLine) 
    } # End if   
   } #End ForEach Guest
} #End Foreach Group
# Get Unique set of guests and groups
$BadGuests = $Report.Guest | Sort -unique
$BadGroups = $Report.Group | Sort -unique
$Report | Export-Csv -NotypeInformation $OutputFile
# Output Details
CLS
Write-Host "Report of Guest Accounts from Blocked Domains"
Write-Host "---------------------------------------------"
Write-Host ""
Write-Host ("{0} Groups processed. {1} groups have guests from blocked domains, with a total of {2} guest accounts." -f $Groups.Count, $BadGroups.Count, $BadGuests.Count)
Write-Host "-----------------------------------------------------------------------------------------------------------"
Write-Host " "
Write-Host ("Guest accounts from blocked domains: {0}" -f ($BadGuests -join ", "))
Write-Host "-----------------------------------------------------------------------------------------------------------"
Write-Host " "
Write-Host "Problem Groups:"
Write-Host "---------------"
$BadGroups
Write-Host " "
Write-Host "Report data is available in" $OutputFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
