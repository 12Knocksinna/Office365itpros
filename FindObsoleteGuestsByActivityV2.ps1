# Version 2 of the script to perform an activity-based analysis of AAD Guest User Accounts and report/highlight
# accounts that aren't being used.  Modules used are Azure AD (V2) and Exchange Online
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindObsoleteGuestsByActivityV2.ps1
# 
# Start by finding all Guest Accounts
$Guests = (Get-AzureADUser -Filter "UserType eq 'Guest'" -All $True| Select Displayname, UserPrincipalName, Mail, ObjectId)
Write-Host $Guests.Count "guest accounts found. Checking their activity..."
$StartDate = Get-Date(Get-Date).AddDays(-90) -Format g #For audit log
$StartDate2 = Get-Date(Get-Date).AddDays(-10) -Format g #For message trace
$EndDate = Get-Date -Format g; $Active = 0; $EmailActive = 0; $Inactive = 0; $AuditRec = 0; $GNo = 0
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
CLS
ForEach ($G in $Guests) {
    $GNo++
    $ProgressBar = "Processing guest " + $G.DisplayName + " (" + $GNo + " of " + $Guests.Count + ")" 
    Write-Progress -Activity "Checking Azure Active Directory Guest Accounts" -Status $ProgressBar -PercentComplete ($GNo/$Guests.Count*100)
    $LastAuditRecord = $Null; $GroupNames = $Null; $LastAuditAction = $Null; $i = 0; $ReviewFlag = $False
    # Search for audit records for this user
    $Recs = (Search-UnifiedAuditLog -UserIds $G.Mail, $G.UserPrincipalName -Operations UserLoggedIn, SecureLinkUsed, TeamsSessionStarted -StartDate $StartDate -EndDate $EndDate -ResultSize 1)
    If ($Recs.CreationDate -ne $Null) { # We found some audit records
       $LastAuditRecord = $Recs[0].CreationDate; $LastAuditAction = $Recs[0].Operations; $AuditRec++}
    # Check email tracking logs because guests might receive email through membership of Outlook Groups. Email address must be valid for the check to work
    If ($G.Mail -ne $Null) {$EmailRecs = (Get-MessageTrace -StartDate $StartDate2 -EndDate $EndDate -Recipient $G.Mail)}           
    If ($EmailRecs.Count -gt 0) {$EmailActive++}
    # Find what Office 365 Groups the guest belongs to
    $DN = (Get-Recipient -Identity $G.UserPrincipalName).DistinguishedName
    $GuestGroups = (Get-Recipient -Filter "Members -eq '$Dn'" -RecipientTypeDetails GroupMailbox | Select DisplayName, ExternalDirectoryObjectId)
    If ($GuestGroups -ne $Null) {
         ForEach ($Group in $GuestGroups) { 
           If ($i -eq 0) { $GroupNames = $Group.DisplayName; $i++ }
         Else 
           {$GroupNames = $GroupNames + "; " + $Group.DisplayName }}}  
     # Figure out age of guest account in days using the creation date in the extension properties of the guest account
     $CreationDate = (Get-AzureADUserExtension -ObjectId $G.ObjectId).get_item("createdDateTime") 
     $AccountAge = ($CreationDate | New-TimeSpan).Days
     # Flag the account for potential deletion if it is more than a year old and isn't a member of any Office 365 Groups.
     If (($AccountAge -gt 365) -and ($GroupNames -eq $Null))  {$ReviewFlag = $True} 
     # Write out report line     
     $ReportLine = [PSCustomObject][Ordered]@{ 
          Guest            = $G.Mail
          Name             = $G.DisplayName
          ReviewForDelete  = $ReviewFlag
          Created          = $CreationDate 
          AgeInDays        = $AccountAge
          EmailCount       = $EmailRecs.Count
          LastConnectOn    = $LastAuditRecord
          LastConnect      = $LastAuditAction
          O365Groups       = $GroupNames } 
       $Report.Add($ReportLine) 
} 
$Report | Sort Name | Export-CSV -NoTypeInformation c:\temp\GuestActivity.csv   
Cls; Write-Host "All Done... The output file is in c:\temp\GuestActivity.csv"      
$Active = $AuditRec + $EmailActive  
Write-Host ""
Write-Host "Statistics"
Write-Host "----------"
Write-Host "Guest Accounts          " $Guests.Count
Write-Host "Active Guests           " $Active
Write-Host "Audit Record foun       " $AuditRec
Write-Host "Active on Email         " $EmailActive
Write-Host "InActive Guests         " ($Guests.Count - $Active)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
