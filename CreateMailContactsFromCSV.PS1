# CreateMailContactsFromCSV.PS1
# Create new mail contact objects from a CSV file
# 31-Dec-2018

$InputContacts = import-csv c:\temp\inputcontacts.csv
Write-Host $InputContacts.Count "contacts found"
ForEach ($Contact in $InputContacts) {
    $Alias = $Contact.First + "." + $Contact.Last
    # Real simple code to make sure that we have an alias to use
    If ($Alias -eq $Null) { $Alias = $N.Name.Split("")[0] + "." + $N.Name.Split("")[1] }
    # Check that the email address isn't known within the tenant - if it's not, we can add the contact
    If ((Get-Recipient -Identity $Contact.EmailAddress -ErrorAction SilentlyContinue) -eq $Null) {
       # Recipient is not known, so we can add them
       Write-Host "Adding contact" $Contact.EmailAddress
       New-MailContact -Name $Contact.Name -ExternalEmailAddress $Contact.EmailAddress -Alias $Alias -FirstName $Contact.First -LastName $Contact.Last
       # Update country and phone numbers
       Set-Contact -Identity $Alias -MobilePhone $Contact.MobilePhone -Phone $Contact.WorkPhone -CountryOrRegion $Contact.Country -Company $Contact.Company }
}
    