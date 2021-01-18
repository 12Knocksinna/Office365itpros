# Script to find guest accounts that are inactive - used in https://petri.com/guest-account-obsolete-activity
# https://github.com/12Knocksinna/Office365itpros/blob/master/FindObsoleteGuestsByActivity.ps1
$Guests = (Get-AzureADUser -Filter "UserType eq 'Guest'" -All $True| Select Displayname, UserPrincipalName, Mail, RefreshTokensValidFromDateTime)
Write-Host $Guests.Count "guest accounts found. Checking their recent activity..."
$StartDate = (Get-Date).AddDays(-90) #For audit log
$StartDate2 = (Get-Date).AddDays(-10) #For message trace
$EndDate = (Get-Date); $Active = 0; $EmailActive = 0; $Inactive = 0; $AuditRec = 0 
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
ForEach ($G in $Guests) {
    Write-Host "Checking" $G.DisplayName  
    $LastAuditAction = $Null; $LastAuditRecord = $Null
    # Search for audit records for this user
    $Recs = (Search-UnifiedAuditLog -UserIds $G.Mail, $G.UserPrincipalName -Operations UserLoggedIn, SecureLinkUsed, TeamsSessionStarted -StartDate $StartDate -EndDate $EndDate -ResultSize 1)
    If ($Recs.CreationDate -ne $Null) { # We found some audit logs
       $LastAuditRecord = $Recs[0].CreationDate; $LastAuditAction = $Recs[0].Operations; $AuditRec++
       Write-Host "Last audit record for" $G.DisplayName "on" $LastAuditRecord "for" $LastAuditAction -Foregroundcolor Green }
    Else { Write-Host "No audit records found in the last 90 days for" $G.DisplayName "; account created on" $G.RefreshTokensValidFromDateTime -Foregroundcolor Red } 
    # Check message trace data because guests might receive email through membership of Outlook Groups. Email address must be valid for the check to work
    If ($G.Mail -ne $Null) {
       $EmailRecs = (Get-MessageTrace –StartDate $StartDate2 –EndDate $EndDate -Recipient $G.Mail)            
       If ($EmailRecs.Count -gt 0) {
           Write-Host "Email traffic found for" $G.DisplayName "at" $EmailRecs[0].Received -foregroundcolor Yellow
           $EmailActive++ }}
     # Write out report line     
     $ReportLine = [PSCustomObject][Ordered]@{ 
          Guest            = $G.Mail
          Name             = $G.DisplayName
          Created          = $G.RefreshTokensValidFromDateTime 
          EmailCount       = $EmailRecs.Count
          LastConnectOn    = $LastAuditRecord
          LastConnect      = $LastAuditAction} 
       $Report.Add($ReportLine)  }          
 
$Active = $AuditRec + $EmailActive
$Report | Export-CSV -NoTypeInformation c:\temp\GuestActivity.csv      
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
