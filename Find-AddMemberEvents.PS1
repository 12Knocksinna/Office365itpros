# A script to show how to detect users with a specifc type of displayname can be detected if someone adds them to 
# a team

Function Add-MessageRecipients {
 # Function to build an addressee list to send email   
 [cmdletbinding()]
     Param(
     [array]$ListOfAddresses )
      ForEach ($SMTPAddress in $ListOfAddresses) {
           @{
              emailAddress = @{address = $SMTPAddress}
           }    
        }
    }       

# Check if we can run an Exchange Online cmdlet. If we can, go on, else connect to Exchange Online
If ($Null -eq (Get-ConnectionInformation)) {
    Connect-ExchangeOnline
}
# Connect to the Graph
# Scopes: Mail.Send.Shared required to send email from a shared mailbox
#         Directory.Read.All required to read user information from Entra ID
Connect-MgGraph -Scopes Mail.Send.Shared, Directory.Read.All

# Find users who have Project in their display name
Write-Host "Finding user details to check"
[array]$Users = Get-MgUser -Search "displayName:Project" -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel Eventual
If (!($Users)) { 
    Throw "No users found"
}

Write-Host ("Checking audit records for {0} users" -f $Users.count)
# Build hash table of users that we want to check
$UserLookup = @{}
ForEach ($User in $Users) {
   $UserLookup.Add($User.UserPrincipalName, $User.DisplayName)
}

$StartDate = (Get-Date).AddDays(-7)
$EndDate = (Get-Date).AddDays(1)
[array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Formatted -ResultSize 5000 `
   -RecordType MicrosoftTeams -Operations MemberAdded
If (!($Records)) {
    Throw "No records found"
}
$Report = [System.Collections.Generic.List[Object]]::new()  

ForEach ($Rec in $Records) {
    $Role = $Null
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    # Check the members noted as added to a group
    ForEach ($Member in $AuditData.Members) {
        If ($UserLookup[$Member.Upn]) {
           Write-Host ("User {0} added to team {1}" -f $Member.DisplayName, $AuditData.TeamName) 
           Switch ($Member.Role) {
            "1"  { $Role = "Member" }
            "2"  { $Role = "Owner"}
            "3"  { $Role = "Guest" }
           }
           $ReportLine = [PSCustomObject]@{
             Date = $AuditData.CreationTime
             User = $Member.Upn   
             Name = $Member.DisplayName
             Team = $AuditData.TeamName
             Role = $Role
             AddedBy = $AuditData.UserId
           }
          $Report.Add($ReportLine)
        }
    }
}

$Report | Out-GridView

$EmailRecipient = "Lotte.Vetler@Office365itpros.com"
Write-Host ("Sending results with {0} monitored events to {1}" -f $Report.count, $EmailRecipient )
# Send a message from the shared mailbox
$MsgFrom = "Customer.Services@Office365itpros.com"
# Add your recipient address here
$ToRecipientList   = @( $EmailRecipient )
[array]$MsgToRecipients = Add-MessageRecipients -ListOfAddresses $ToRecipientList
$MsgSubject = "Monitored User Additions to Team membership"
$HtmlHead = "<h2>Monitored User Additions to Teams</h2>"
$HtmlBody = $Report | ConvertTo-Html -Fragment 
$HtmlMsg = "</body></html><p>" + $HtmlHead + $Htmlbody + "<p>"
# Construct the message body
$MsgBody = @{
  Content = "$($HtmlMsg)"
  ContentType = 'html'  
 }

$Message =  @{subject           = $MsgSubject}
$Message += @{toRecipients      = $MsgToRecipients}  
$Message += @{body              = $MsgBody}
$Params   = @{'message'         = $Message}
$Params  += @{'saveToSentItems' = $True}
$Params  += @{'isDeliveryReceiptRequested' = $True}

# And send the message using the parameters that we've filled in
Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params