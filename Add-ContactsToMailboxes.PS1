# Add-ContactsToMailboxes.PS1
# Demonstration of how to add contacts to user mailboxes with the Microsoft Graph PowerShell SDK
# V1.0 14-Aug-2024
# GitHub link: 
$CertThumbPrint = "F79286DB88C21491110109A0222348FACF694CBD"
$AppId = "0c75253c-3c01-43ed-adcf-1415c68ff455"
$TenantId = "22e90715-3da6-4a78-9ec6-b3282389492b"

# We need to use application permissions to add contacts to mailboxes
Connect-MgGraph -NoWelcome -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $CertThumbPrint 

# Define the service plan IDs for Exchange Online (Plan 1) and Exchange Online (Plan 2)
$ExoServicePlan1 = "9aaf7827-d63c-4b61-89c3-182f06f82e5c"
$ExoServicePlan2 = "efb87545-963c-4e0d-99df-69c6916d9eb0" 
# Define the check date to find new mailboxes
$CheckDate = [datetime]::UtcNow.AddDays(-7).ToString("s") + "Z"

# Find the organizational contacts to add
$Site = Get-MgSite -Search 'Experts Team'
If (!$Site) {
    Write-Host "Site not found"
    Break
}   
$List = Get-MgSiteList -SiteId $Site.Id -Filter "displayName eq 'Organizational Contacts'"
If (!$List) {
    Write-Host "List not found"
    Break
}   

[array]$ListItems = Get-MgSiteListItem -ListId $List.Id -SiteId $Site.Id -PageSize 999 -All `
    -ExpandProperty "fields(`$select=id,title,fullname,firstname,workphone,email,workcity,description)" 
$ItemData = [System.Collections.Generic.List[Object]]::new()
ForEach ($Item in $ListItems.fields) { 
    
    $ReportLine = [PSCustomObject] @{ 
        Id          = $Item.Id
        FullName    = $Item.AdditionalProperties.FullName
        FirstName   = $Item.AdditionalProperties.FirstName
        LastName    = $Item.AdditionalProperties.Title
        PhoneNumber = $Item.AdditionalProperties.WorkPhone
        Email       = $Item.AdditionalProperties.Email
        City        = $Item.AdditionalProperties.WorkCity
        Description = $Item.AdditionalProperties.Description
    }
    $ItemData.Add($ReportLine)
}
Write-Host ("Found {0} organizational contacts to process" -f $ItemData.Count)

# Find users assigned a license that includes the Exchange Online (Plan 1) or Exchange Online (Plan 2) service plans who were created
# since the check date. The check also looks for users with the Exchange Online service plan enabled.
[array]$Users = Get-MgUser -Filter "assignedPlans/any(c:c/servicePlanId eq $ExoServicePlan1 and capabilityStatus eq 'Enabled') `
  or assignedPlans/any(c:c/servicePlanId eq $ExoServicePlan2 and capabilityStatus eq 'Enabled') and (CreatedDateTime ge $CheckDate)" `
    -ConsistencyLevel eventual -CountVariable Test -All -PageSize 999 -Sort ('displayname') `
    -Property Id, displayName, userprincipalName, assignedLicenses, assignedPlans, department, country, CreatedDateTime

If ($Users.count -eq 0) {
    Write-Host "No new mailboxes found to update with organizational contacts"
    Break 
} Else {
    Write-Host ("Found {0} new mailboxes to update with organizational contacts" -f $Users.Count)
}

[int]$ContactAdded = 0
ForEach ($User in $Users) {
    Write-Host ("Processing user {0}" -f $User.displayName)
    # Get the existing contacts so that we can avoid adding duplicates
    [array]$Contacts = Get-MgUserContact -UserId $User.Id -All -PageSize 999 | Select-Object -ExpandProperty $EmailAddresses
    $ContactsHash = @{}
    ForEach ($Contact in $Contacts) {
        # Before adding, check that the contact has an email address and that it's not already in the hash
        $EmailAddress = $Contact.emailAddresses[0].address
        If ($EmailAddress -and $null -eq $ContactsHash[$EmailAddress]) {
            $ContactsHash.Add($EmailAddress,$Contact.displayName)
        }
    }

    ForEach ($Item in $ItemData) {

        $NewContactEmail = @{
            address = $Item.Email
            name = $Item.FullName
        }
        $NewContactEmails = @($NewContactEmail)

        # Build body for new contact
        $NewContact = @{}
        $NewContact.Add("FileAs", $Item.FullName)
        $NewContact.Add("Surname", $Item.LastName)
        $NewContact.Add("GivenName", $Item.FirstName)
        $NewContact.Add("OfficeLocation", $Item.City)
        $NewContact.Add("PersonalNotes", $Item.Description)
        $NewContact.Add("emailAddresses", $NewContactEmails)
        $NewContact.Add("businessPhones", @($Item.PhoneNumber))
        $NewContact.Add("DisplayName", $Item.FullName)

        # Check if the contact already exists and add it if it's not there
        If ($null -eq $ContactsHash[$Item.Email]) {
            Write-Host ("Adding contact {0} to {1}" -f $Item.FullName, $User.displayName)
            Try {
                $Contact = New-MgUserContact -UserId $User.Id -BodyParameter $NewContact
            } Catch {
                Write-Host ("Failed to add contact {0} to {1}" -f $Item.FullName, $User.displayName)
            }
            If ($Contact) {
                $ContactAdded++
            }   
        }
    }
}

Write-Host ("Added {0} contacts to {1} mailboxes" -f $ContactAdded, $Users.Count)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.