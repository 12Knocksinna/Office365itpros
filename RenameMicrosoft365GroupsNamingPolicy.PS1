# RenameMicrosoft365GroupsNamingPolicy.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/RenameMicrosoft365GroupsNamingPolicy.PS1
# A script to rename Microsoft 365 Groups that don't comply with the tenant naming policy

# Check that we are connected to Exchange Online, Azure Active Directory, and Teams
Write-Host "Checking that prerequisite PowerShell modules are loaded..."
Try { $OrgName = (Get-OrganizationConfig).Name }
   Catch  {
      Write-Host "Your PowerShell session is not connected to Exchange Online."
      Write-Host "Please connect to Exchange Online using an administrative account and retry."
      Break }

$AzureADCheck = Get-Module -Name AzureADPreview     
If ($AzureADCheck -eq $Null) {
     Write-Host "Your PowerShell session is not connected to Azure Active Directory (Preview)."
     Write-Host "Please connect to Azure Active Directory using an administrative account and retry."; Break } 

# Get organization naming policy setting
$Policy = Get-AzureADDirectorySetting | ? {$_.DisplayName -eq "Group.Unified"}
$NamingPolicy = $Policy["PrefixSuffixNamingRequirement"]

If (!($NamingPolicy)) 
   { Write-Host "No naming policy defined..."
     EXIT }
   Else
   { Write-Host "Groups naming policy is" $NamingPolicy }

# Check if a prefix or a suffix is used to name groups
$Suffix = $False; $Prefix = $False
If ($NamingPolicy.Substring(0,11) -eq "[GroupName]") { # Suffix is used, not a prefix
    $Suffix = $NamingPolicy.Split("]")[1]
    $GroupsNameMatch = "*" + $Suffix + "*" }
Else { #Assume a prefix
    $Prefix = $NamingPolicy.SubString(0,($NamingPolicy).IndexOf("[GroupName"))
    $GroupsNameMatch = "*" + $Prefix + "*" }
# Find Microsoft 365 Groups that don't match the naming policy
$Groups = (Get-UnifiedGroup | ? {$_.DisplayName -NotLike $GroupsNameMatch})

If ($Groups.Count -gt 0)
   { $Prompt = "You have " + $Groups.Count + " groups to update. Proceed? [Y/N]"
     $Answer = Read-host -Prompt $Prompt 
     If ($Answer.ToUpper() -ne "Y")
        {
         Write-Host "Exiting..." 
         EXIT }
     }
# Update Groups
Write-Host "Updating Groups with new display names..."

ForEach ($G in $Groups) {
   If ($Prefix -ne $False) { # Update with a new prefix
        $NewDisplayName = $Prefix + $G.DisplayName
        Write-Host $G.DisplayName "updated to" $NewDisplayName
        Set-UnifiedGroup -Identity $G.DistinguishedName -DisplayName $NewDisplayName 
   } #End If
   Elseif ($Suffix -ne $False) {
        $NewDisplayName = $G.DisplayName + $Suffix
        Write-Host $G.DisplayName "updated to" $NewDisplayName
        Set-UnifiedGroup -Identity $G.DistinguishedName -DisplayName $NewDisplayName 
   } #End Elseif
} #End ForEach

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
