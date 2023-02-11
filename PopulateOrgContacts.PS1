# PopulateOrgContacts.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/PopulateOrgContacts.PS1
# Based on idea in https://github.com/smcavinue/AdminSeanMc/blob/master/Graph%20Scripts/graph-PopulatecontactsfromCSV/graph-PopulateContactsFromCSV.ps1
# A script to write organization contacts to user mailboxes
# Needs an app registered in Azure AD with consent given for the Application Contacts.ReadWrite permission
# 
# Check we have the right module loaded
$Modules = Get-Module
If ("ExchangeOnlineManagement" -notin  $Modules.Name) {Write-Host "Please connect to Exchange Online Management  before continuing...";break}

# Find the set of organization contacts - marked as such by having OrgContact in CustomAttribute4 of their object properties. This attribute is chosen at random. You can use
# whatever other way you want to find the set of contacts, including having a CSV file.
[array]$OrgContacts = Get-ExoRecipient -Filter {CustomAttribute4 -eq "OrgContact"} -Properties CustomAttribute4, ExternalEmailAddress -RecipientTypeDetails MailContact 
If (!($OrgContacts)) {Write-Host "No organization contacts found - exiting" ; break }
Write-Host ("Found {0} organization contacts - continuing..." -f $OrgContacts.count)

# Look for target mailboxes. In this example, we get the mailboxed created in the last month
[datetime]$LastMonth = (Get-Date).AddDays(-30)
[array]$Mailboxes = Get-ExoMailbox -Filter "WhenMailboxCreated -gt '$LastMonth'" -RecipientTypeDetails UserMailbox | Select ExternalDirectoryObjectId, DisplayName, UserPrincipalName
If (!($Mailboxes)) { Write-Host "No mailboxes found to process - exiting" ; break }
Write-Host ("Found {0} mailboxes to process - continuing..." -f $Mailboxes.count)
$Now = Get-Date -format D

# Get Graph access token - change these values for the app you use.
$AppSecret = "~C2U~tDO74Tj-w2glcIWqa_BtCZ46KmVA."
$AppId = "c1e486db-d511-4778-bc36-57213fa3427f"
$Tenantid = "b762313f-14fc-43a2-9a7a-d2e27f4f3478"

# Construct URI and body needed for authentication
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"
}

$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
# Unpack Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
$Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $Token" }

$OrgNotes = "Organization contact created " + $Now

ForEach ($Mbx in $Mailboxes) {
   # Populate email addresses for existing contacts
   $Uri = "https://graph.microsoft.com/v1.0/users/" + $Mbx.ExternalDirectoryObjectId + "/contacts"
   [array]$ContactsInMbx =  Invoke-RestMethod -Headers $Headers -Uri $Uri -UseBasicParsing -Method "GET" 
   # Build hash table of contacts that exist in the mailbox
   If ($ContactsInMbx.Value.Count -gt 0) { $ExistingContacts = $ContactsInMbx.Value | Select-Object -ExpandProperty emailaddresses | Sort-Object -Unique }
   $CheckTable = @{}
   ForEach ($C in $ExistingContacts) { $CheckTable.Add($C.Address.toString(), $C.Name.toString()) }
   Write-Host "Processing mailbox" $Mbx.DisplayName
   $ApiUri = "https://graph.microsoft.com/v1.0/users/$mailbox/contacts" 
   ForEach ($Contact in $OrgContacts) {
    Write-Host "Processing contact" $Contact.DisplayName
    $Assistant   = $Null
    $Phone       = $Contact.Phone
    $HomePage    = $Null
    $Company     = $Contact.Company
    $Department  = $Contact.Department
    $DisplayName = $Contact.DisplayName
    $Title       = $Contact.Title
    $First       = $Contact.FirstName
    $Last        = $Contact.LastName
    $Middle      = $Null
    $Nickname    = $Contact.Alias
    $Notes       = $OrgNotes 
    Switch ($Contact.RecipientTypeDetails) { # Populate the contact details depending on the type of mail recipient object
     "MailContact" { # Mail contacts
         $Email       = $Contact.ExternalEmailAddress.Split(":")[1]   
         $Profession  = "Contact"       
     }
    "MailUniversalDistributionGroup"  { # Distribution lists
         $Email       = $Contact.PrimarySmtpAddress
         $Profession  = "Distribution list"
     }
    "GroupMailbox"  { # Microsoft 365 Groups
         $Email       = $Contact.PrimarySmtpAddress
         $Profession  = "Microsoft 365 Group"
     }
    "UserMailbox" { # User mailboxes
         $Email       = $Contact.PrimarySmtpAddress
         $Profession  = "User Mailbox"
     }
    "SharedMailbox" { # User mailboxes
         $Email       = $Contact.PrimarySmtpAddress
         $Profession  = "Shared Mailbox"
     }
   } # End Switch

# Check if the contact is already there. If not, we go ahead and add the contact
    If ($CheckTable[$Email]) {
      Write-Host ("Contact record for {0} is already present in the mailbox" -f $Email) }
   Else {
      Write-Host "Proceeding..."
      # Build the contact object
      $ContactObject = @"
      {
        "assistantName":    "$($Assistant)",
        "businessHomePage": "$($HomePage)",
        "businessPhones": [
            "$($Phone)"
          ],
        "companyName":      "$($Company)",
        "department":       "$($Department)",
        "displayName":      "$($DisplayName)",
        "emailAddresses": [
            {
                "address":  "$($Email)",
                "name":     "$($Displayname)"
            }
        ],
        "givenName":       "$($First)",
        "jobTitle":        "$($Title)",
        "middleName":      "$($Middle)",
        "nickName":        "$($Nickname)",
        "profession":      "$($Profession)",
        "personalNotes":   "$($OrgNotes)",
        "surname":         "$($Last)",
        "title":           "$($Saluation)"
       }
"@
   
    # And add the new contact
    Try {
        $NewContact = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -ContentType 'application/json' -Body $ContactObject -Uri $Uri -Method Post)
    }
    catch {
        throw "Error creating contact $($contact.emailaddress) for $mailbox $($_.Exception.Message)"
        break
    }
  } #End Else
 } #End ForEach OrgContacts
} #End ForEach Mailboxes


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
