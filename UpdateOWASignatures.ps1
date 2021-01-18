# UpdateOWASignatures.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateOWASignatures.ps1
# Update the OWA signature data for all mailboxes to make sure that they comply with company rules
CLS
# Define some variables that we'll use in the HTML output
# You need to use your own image files for company logo etc. Define them here.
$HeadingLine      = "<HTML><HEAD><TITLE>Signature</TITLE><BODY><BR><table style=`"FONT-SIZE: 8pt; COLOR: gray; FONT-FAMILY: `'Segoe UI`' `"> <tr>"
$CompanyLogo      = "https://i1.wp.com/office365itpros.com/wp-content/uploads/2020/02/2020EditionVerySmall.jpg"
$CompanyLogoLine  = "<td ><img src='" + $CompanyLogo + "' border='0'></td>"
$TwitterLogo      = "https://i1.wp.com/office365itpros.com/wp-content/uploads/2020/02/Twitter.png" 
$TwitterLink      = "https://twitter.com/12Knocksinna"
$FacebookLogo     = "https://i0.wp.com/office365itpros.com/wp-content/uploads/2020/02/Facebook.png"
$FacebookLink     = "https://www.facebook.com/Office365itpros/"
# Facebook and Twitter icons inserted into HTML
$IconsLine        = '<tr><td style="font-size: 10pt; font-family: Arial, sans-serif; padding-bottom: 0px; padding-top: 5px; padding-left: 10px; vertical-align: bottom;" valign="bottom"><span><a href="' + $($FacebookLink) + '" target="_blank" rel="noopener"><img border="0" width="23" alt="facebook icon" style="border:0; height:23px; width:23px" src="' + $($FacebookLogo) + '"></a> </span><span><a href="' + $($TwitterLink) + '" target="_blank" rel="noopener"><img border="0" width="23" alt="twitter icon" style="border:0; height:23px; width:23px" src="' + $($TwitterLogo) + '"></a></span></td></tr>'
$EndLine          = "</td></tr></table><BR><BR></BODY></HTML>"

Write-Host "Fetching mailbox information..."
$Mbx = Get-User -RecipientTypeDetails UserMailbox -ResultSize Unlimited 
$ProgDelta = 100/($Mbx.count); $CheckCount = 0
CLS
Write-Host "Processing mailboxes"
ForEach ($M in $Mbx) {
  $MbxNumber++
  $MbxStatus = $M.DisplayName + " ["+ $MbxNumber +"/" + $Mbx.Count + "]"
  Write-Progress -Activity "Processing mailbox" -Status $MbxStatus -PercentComplete $CheckCount
  $CheckCount += $ProgDelta
  # Populate user properties
  # Make sure that we have a valid postal address instead of taking an address from the user's AAD record (which could be personal)
  Switch ($M.City)  { # Define this information for your own company
     "Foxrock"      {$City = "Dublin"; $StreetAddress = "Foxrock"; $PostalCode = "D18A52R2 Ireland" }
     "Frankfurt"    {$City = "Frankfurt am Main"; $StreetAddress = "Freidrich-Ebert-Anlage 35-37"; $PostalCode = "D-60327 Germany"} 
     "San Franciso" {$City = "San Francisco"; $StreetAddress = "14 Warren Street"; $PostalCode = "93404 United States of America"}
     Default        {$City = "Dublin"; $StreetAddress = "Foxrock"; $PostalCode = "D18A52R2 Ireland" }
   }
   # Make sure we have a company name
   If ($Null -eq $M.Company) { $CompanyName = "Office 365 for IT Pros"} Else { $CompanyName = $M.Company }
   If ($Null -eq $M.Title) { $JobTitle = "Valued Employee" } Else { $JobTitle = $M.Title }
   # Now build the HTML Info for the signature
   $PersonLine       = "<td padding='0'><B>" + $M.DisplayName + " </B> " + $JobTitle + "<BR>"
   # Create a Mailto: link for the user's email address
   $EmailLink        = '<a href=mailto:"' + $($M.WindowsEmailAddress) + '">' + $($M.WindowsEmailAddress) + '</a>'
   $CompanyLine      = "<b>" + $CompanyName + "</b> " + $StreetAddress + ", " + $City + ", " + $PostalCode + "<BR>" + $M.Phone + "/" + $M.MobilePhone + " Email: " +  $EmailLink + "<br><br>"     
   $SignatureHTML = $HeadingLine + $CompanyLogoLine + $PersonLine + $CompanyLine + $IconsLine + $EndLine
   Set-MailboxMessageConfiguration -Identity $M.UserPrincipalName -SignatureHTML $SignatureHTML -AutoAddSignature $True -AutoAddSignatureOnReply $False 
}
Write-Host "All done. Signatures updated for" $Mbx.count "mailboxes"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
