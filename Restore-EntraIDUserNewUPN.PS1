# Restore-EntraIDUserNewUPN.PS1 - Restore a deleted Entra ID user account with a new UPN
#
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Restore-EntraIDUserNewUPN.PS1

# Directory.ReadWrite.All will work, but User.DeleteRestore.All is a lower more specific permission
Connect-MgGraph -Scopes User.DeleteRestore.All

$DeletedUserUPN = Read-Host "What is the UPN of the deleted user account ?"
$UserDisplayName = $DeletedUserUPN.split('@')[0].Replace("."," ")

$NewUserPrincipalName = Read-Host "What is the new UPN for the restored account ?"
$CheckAccount = Get-MgUser -UserId $NewUserPrincipalName -ErrorAction SilentlyContinue
If ($CheckAccount) { 
    Write-Host "The new UPN is already in use. Please choose another" 
    Break
}   

$Headers = @{}
$Headers.Add("consistencylevel","eventual")
$NewUPNDetails = @{}
$NewUPNDetails.Add("newUserPrincipalName",$NewUserPrincipalName)
$NewUPNDetails.Add("autoReconcileProxyConflict",$true)  

# Attempt to find the deleted object - for whatever reason startsWith doesn't work
$DeletedObject = Get-MgDirectoryDeletedItemAsUser -Filter "endsWith(UserPrincipalName,'$DeletedUserUPN')" -Headers $Headers -CountVariable DeletedCount

If ($DeletedObject) { 
    Write-Host "Restoring Object with new User principal Name"
    $Status = Restore-MgBetaDirectoryDeletedItem -DirectoryObjectId $DeletedObject.Id -BodyParameter $NewUPNDetails
    If ($Status) { 
        Write-Host ("Account restored for {0} with UPN using {1}" -f $UserDisplayName, $NewUserPrincipalName)
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.

