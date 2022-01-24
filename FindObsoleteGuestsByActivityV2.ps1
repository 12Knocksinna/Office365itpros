# Version 2 of the script to perform an activity-based analysis of AAD Guest User Accounts and report/highlight
# accounts that aren't being used.  Modules used are Azure AD (V2) and Exchange Online
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindObsoleteGuestsByActivityV2.ps1
# 
# Start by finding all Guest Accounts
$Guests = (Get-AzureADUser -Filter "UserType eq 'Guest'" -All $True| Select Displayname, UserPrincipalName, Mail, ObjectId)
Write-Host $Guests.Count "guest accounts found. Checking their activity..."
$StartDate = Get-Date(Get-Date).AddDays(-90) #For audit log
$StartDate2 = Get-Date(Get-Date).AddDays(-10) #For message trace
$EndDate = Get-Date; $Active = 0; $EmailActive = 0; $Inactive = 0; $AuditRec = 0; $GNo = 0
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report
CLS; $GNo = 0
ForEach ($G in $Guests) {
    $GNo++
    $ProgressBar = "Processing guest " + $G.DisplayName + " (" + $GNo + " of " + $Guests.Count + ")" 
    Write-Progress -Activity "Checking Azure Active Directory Guest Accounts for activity" -Status $ProgressBar -PercentComplete ($GNo/$Guests.Count*100)
    $LastAuditRecord = $Null; $GroupNames = $Null; $LastAuditAction = $Null; $i = 0; $ReviewFlag = $False
    # Search for audit records for this user
    $Recs = (Search-UnifiedAuditLog -UserIds $G.Mail, $G.UserPrincipalName -Operations UserLoggedIn, SecureLinkUsed, TeamsSessionStarted -StartDate $StartDate -EndDate $EndDate -ResultSize 1)
    If ($Recs) { # We found some audit records
       $LastAuditRecord = $Recs[0].CreationDate; $LastAuditAction = $Recs[0].Operations; $AuditRec++}
    Else {
       $LastAuditRecord = "None found"; $LastAuditAction = "N/A" }
    # Check email tracking logs because guests might receive email through membership of Outlook Groups. Email address must be valid for the check to work
    If ($G.Mail -ne $Null) {$EmailRecs = (Get-MessageTrace -StartDate $StartDate2 -EndDate $EndDate -Recipient $G.Mail)}           
    If ($EmailRecs.Count -gt 0) {$EmailActive++}
    # Figure out the domain the guest is from so that we can report this information
    $Domain = ($G.UserPrincipalName.Split("#EXT#")[0]).Split("_")[1]
    # Find what Microsoft 365 Groups the guest belongs to
    $GroupNames = $Null
    $Dn = (Get-Recipient -Identity $G.UserPrincipalName).DistinguishedName
    If ($Dn -like "*'*")  {
       $DNNew = "'" + "$($dn.Replace("'","''''"))" + "'"
       $Cmd = "Get-Recipient -Filter 'Members -eq '$DNnew'' -RecipientTypeDetails GroupMailbox | Select DisplayName, ExternalDirectoryObjectId"
       $GuestGroups = Invoke-Expression $Cmd }
     Else {
       $GuestGroups = (Get-Recipient -Filter "Members -eq '$Dn'" -RecipientTypeDetails GroupMailbox | Select DisplayName, ExternalDirectoryObjectId) }
     If ($GuestGroups -ne $Null) {
       $GroupNames = $GuestGroups.DisplayName -join ", " }

     # Figure out age of guest account in days using the creation date in the extension properties of the guest account
     $CreationDate = (Get-AzureADUserExtension -ObjectId $G.ObjectId).get_item("createdDateTime") 
     $AccountAge = ($CreationDate | New-TimeSpan).Days
      # Find if there's been any recent sign on activity
     $UserLastLogonDate = $Null
     $UserObjectId = $G.ObjectId
     $UserLastLogonDate = (Get-AzureADAuditSignInLogs -Top 1  -Filter "userid eq '$UserObjectId' and status/errorCode eq 0").CreatedDateTime 
     If ($UserLastLogonDate -ne $Null) {
        $UserLastLogonDate = Get-Date ($UserLastLogonDate) -format g }
     Else {
        $UserLastLogonDate = "No recent sign in records found" }
     # Flag the account for potential deletion if it is more than a year old and isn't a member of any Office 365 Groups.
     If (($AccountAge -gt 365) -and ($GroupNames -eq $Null))  {$ReviewFlag = $True} 
     # Write out report line     
     $ReportLine = [PSCustomObject]@{ 
          Guest             = $G.Mail
          Name              = $G.DisplayName
          Domain            = $Domain
          Inactive          = $ReviewFlag
          Created           = $CreationDate 
          AgeInDays         = $AccountAge
          EmailCount        = $EmailRecs.Count
          "Last sign-in"       =  $UserLastLogonDate
          "Last Audit record"  = $LastAuditRecord
          "Last Audit action"  = $LastAuditAction
          "Members of"      = $GroupNames } 
       $Report.Add($ReportLine) 
} 
$Report | Sort Name | Export-CSV -NoTypeInformation c:\temp\GuestActivity.csv   
Cls     
$Active = $AuditRec + $EmailActive  
# Figure out the domains guests come from
$Domains = $Report.Domain | Sort
$DomainsCount = @{}
$Domains | ForEach {$DomainsCount[$_]++}
$DomainsCount = $DomainsCount.GetEnumerator() | Sort -Property Value -Descending
$DomainNames = $Domains | Sort -Unique

$PercentInactive = (($Guests.Count - $Active)/$Guests.Count).toString("P")
Write-Host ""
Write-Host "Statistics"
Write-Host "----------"
Write-Host "Guest Accounts          " $Guests.Count
Write-Host "Active Guests           " $Active
Write-Host "Audit Record found      " $AuditRec
Write-Host "Active on Email         " $EmailActive
Write-Host "InActive Guests         " ($Guests.Count - $Active)
Write-Host "Percent inactive guests " $PercentInactive
Write-Host ("Domain with most guests  {0} ({1})" -f $DomainsCount[0].Name, $DomainsCount[0].Value)
Write-Host "Guests from domains     " ($DomainNames -join ", ")
Write-Host " "
Write-Host "The output file containing detailed results is in c:\temp\GuestActivity.csv" 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
