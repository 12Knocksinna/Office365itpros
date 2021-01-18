# UpdateOutlookSignature.PS1 A script to Update the Outlook signature on a PC
# Written as an example, not as a final production-ready script
# V1.0 17-Feb-2020  Tony Redmond
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateOutlookSignature.PS1
# ==============================================================================

# Define the name of the signature
$SignatureName = 'Office365ITPros' 
$LocalSignaturePath = (Get-Item env:appdata).Value + '\Microsoft\Signatures'
$HtmlPath = $LocalSignaturePath + '\' + $SignatureName + '.htm'

# Find the User Principal Name for the account as stored in the system registry
$UserAccount = Get-ItemProperty -Path HKCU:\Software\Microsoft\Office\Outlook\Settings -Name Accounts | Select -ExpandProperty Accounts
$UserId = (ConvertFrom-Json $UserAccount).UserUpn[0]

# Retrieve the properties of the user from Azure Active Directory
$UserProperties = Get-AzureADUser -ObjectId $UserId

# Find Outlook Profiles in registry
$CommonSettings = $False
$Profiles = (Get-ChildItem HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles).PSChildName
# This script can only deal with a single (default profile); more code needed to handle multiple profiles
If ($Profiles -eq $Null -or $Profiles.Count -ne 1) {
   Write-Host "Warning - Applying signature to all Outlook profiles" 
   $OutlookProfilePath = "HKCU:\Software\Microsoft\\Office\16.0\Common\MailSettings"
   $CommonSettings = $True}
Else { # Path to default profile is elsewhere in the registry
   $OutLookProfilePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\" + $Profiles.Trim() + "\9375CFF0413111d3B88A00104B2A6676\00000001" }

# If we have an Outlook profile, check that we can match the User Principal name with the account name that's stored in the registry for Outlook ProPlus 
# But only do this for now if we're not updating all profiles 
If ($CommonSettings -eq False) {
    $OutlookProfile = Get-ItemProperty -Path $OutLookProfilePath
    If ($OutlookProfile."Account Name" -ne $UserId) { # We don't have a profile match
       Write-Host "Can't match signature and Office 365 user principal name" ; exit }}

Switch ($UserProperties.City)  { # Make sure that users from the same location have the same address
     "Foxrock"      {$City = "Dublin"; $StreetAddress = "Foxrock"; $PostalCode = "D18A52R2 Ireland" }
     "Frankfurt"    {$City = "Frankfurt am Main"; $StreetAddress = "Freidrich-Ebert-Anlage 35-37"; $PostalCode = "D-60327 Germany"} 
     "San Franciso" {$City = "San Francisco"; $StreetAddress = "14 Warren Street"; $PostalCode = "93404 United States of America"}
      Default       {$City = "Dublin"; $StreetAddress = "Foxrock"; $PostalCode = "D18A52R2 Ireland" }
}
# Make sure we have a company name
If ($Null -eq $UserProperties.CompanyName) { $CompanyName = "Office 365 for IT Pros"} Else { $CompanyName = $UserProperties.CompanyName }
If ($Null -eq $UserProperties.Title) { $JobTitle = "Valued Employee" } Else { $JobTitle = $UserProperties.Title }

# Construct a signature file in HTML format using the information fetched from Azure Active Directory
$CompanyLogo      = "https://i1.wp.com/office365itpros.com/wp-content/uploads/2020/02/2020EditionVerySmall.jpg"
$HeadingLine      = "<HTML><HEAD><TITLE>Signature</TITLE><BODY><BR><table style=`"FONT-SIZE: 8pt; COLOR: gray; FONT-FAMILY: `'Segoe UI`' `"> <tr>"
$ImageLine        = "<td ><img src='" + $CompanyLogo + "' border='0'></td>"
$PersonLine       = "<td padding='0'><B>" + $UserProperties.DisplayName + " </B> " + $JobTitle + "<BR>"
$CompanyLine      = "<b>" + $CompanyName + "</b> " + $StreetAddress + ", " + $City + ", " + $PostalCode + "<BR>" + $UserProperties.TelephoneNumber + "/" + $UserProperties.Mobile + " Email: " + $UserProperties.Mail + "<br><br>"
# Facebook and Twitter icons
$IconsLine        = '<tr><td style="font-size: 10pt; font-family: Arial, sans-serif; padding-bottom: 0px; padding-top: 5px; padding-left: 10px; vertical-align: bottom;" valign="bottom"><span><a href="https://www.facebook.com/Office365itpros/" target="_blank" rel="noopener"><img border="0" width="23" alt="facebook icon" style="border:0; height:23px; width:23px" src="https://i0.wp.com/office365itpros.com/wp-content/uploads/2020/02/Facebook.png"></a> </span><span><a href="https://twitter.com/12Knocksinna" target="_blank" rel="noopener"><img border="0" width="23" alt="twitter icon" style="border:0; height:23px; width:23px" src="https://i1.wp.com/office365itpros.com/wp-content/uploads/2020/02/Twitter.png"></a></span></td></tr>'
$EndLine          = "</td></tr></table><BR><BR></BODY></HTML>"

# Put everything together and output the HTML file
$SignatureHTML = $HeadingLine + $ImageLine + $PersonLine + $CompanyLine + $Iconsline + $EndLine | Out-File $HtmlPath

# Update the registry settings where Outlook picks up its signature information
If (Test-Path $TargetForSignatures) {
   Get-Item -Path $OutlookProfilePath | New-Itemproperty -Name "New Signature" -value $SignatureName -Propertytype string -Force 
   Get-Item -Path $OutlookProfilePath | New-Itemproperty -Name "Reply-Forward Signature" -value $SignatureName -Propertytype string -Force }

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
