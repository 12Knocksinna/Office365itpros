# ReportDLMembershipsCounts.PS1
# Report the membership and counts for distribution lists in Exchange Online
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportDLMembershipsCounts.PS1

Function Get-RecursiveAzureAdGroupMemberUsers{
[cmdletbinding()]
# Modified from the code on https://saemundsson.se/?p=734
param(
   [parameter(Mandatory=$True,ValueFromPipeline=$true)]
   $AzureGroup
)
   [array]$Members = Get-AzureADGroupMember -ObjectId $AzureGroup.ObjectId -All $true
   [array]$GroupMembers = $Members | ? {$_.ObjectType -eq 'Group'} | Select ObjectId, DisplayName, Mail, @{Name="Type"; e={$_.ObjectType}}
        
   [array]$UserMembers = $Members | ? {$_.ObjectType -ne 'Group'} | Select ObjectId, DisplayName, Mail, @{Name="Type"; e= {$_.UserType}}
   If($GroupMembers) {
       [array]$ExtractedMembers = $Members | Where-Object{$_.ObjectType -eq 'Group'} | ForEach-Object{ Get-RecursiveAzureAdGroupMemberUsers -AzureGroup $_}
        $UserMembers += $ExtractedMembers }

   # Figure out what mix of user (including mail contacts and other recipients) and group members we have
   [array]$ReturnMembers = $Null
   If ($UserMembers) { $ReturnMembers = $UserMembers}
   If ($GroupMembers) { $ReturnMembers += $GroupMembers }

   # Remove duplicates if found in multiple (nested) DLs
   $ReturnMembers = $ReturnMembers | Sort ObjectId -Unique
   Return $ReturnMembers
}

# Check we have the right modules loaded
$Modules = Get-Module
If ("ExchangeOnlineManagement" -notin  $Modules.Name) {Write-Host "Please connect to Exchange Online Management  before continuing...";break}
If ("AzureADPreview" -notin  $Modules.Name) {Write-Host "Please connect to Azure AD  before continuing...";break}

# Find all distribution lists
Write-Host "Finding Exchange Online Distribution Lists..."
# Find distribution lists, excluding room lists
$DLs = Get-DistributionGroup -ResultSize Unlimited -Filter {RecipientTypeDetails -ne "Roomlist"} | Select DisplayName, ExternalDirectoryObjectId, ManagedBy
If (!($DLs)) { Write-Host "No distribution lists found... sorry! "; break }
 Else { Write-Host ("{0} distribution lists found" -f $DLs.count) }

CLS; $DLNumber = 0
$Report = [System.Collections.Generic.List[Object]]::new()
$DLCSVOutput = "c:\temp\DLMemberCounts.CSV"

ForEach ($DL in $DLs) { 
  $DLNumber++
  $ProgressBar = "Processing distribution list " + $DL.DisplayName + " (" + $DLNumber + " of " + $DLs.Count + ")" 
  Write-Progress -Activity "Analzying membership of distribution list " -Status $ProgressBar -PercentComplete ($DLNumber/$DLs.Count*100)
  [array]$Members =  Get-AzureADGroup -ObjectId $DL.ExternalDirectoryObjectId | Get-RecursiveAzureAdGroupMemberUsers | Sort DisplayName
  If (!($Members)) { # For whatever reason, no members in this group, so we zeroize everything to prepare for the next group
        $CountOfMembers = 0
        [array]$TenantMembers = @() 
        [array]$GroupMembers = @() 
        [array]$GuestMembers = @() 
        [array]$OtherMembers = @() 
        $MemberNames = $Null }  
    Else { 
       [int]$CountOfMembers = $Members.Count
       [array]$TenantMembers = ( $Members | ? {$_.Type -eq "Member" })
       [array]$GroupMembers = ( $Members | ? {$_.Type -eq "Group" })
       [array]$GuestMembers = ( $Members | ? {$_.Type -eq "Guest" })
       [array]$OtherMembers = ( $Members | ? {$_.Type -eq $Null })
       $MemberNames = $Members.DisplayName -join ", " | Out-String
    }
    [string]$OutputNames = $MemberNames 
     $ReportLine = [PSCustomObject][Ordered]@{  
       DLName             = $DL.DisplayName
       ManagedBy          = $DL.ManagedBy -join ", "
       "Members"          = $CountOfMembers
       "Tenant Users"     = $TenantMembers.Count
       "Groups"           = $GroupMembers.Count
       "Guest members"    = $GuestMembers.Count
       "Other Recipients" = $OtherMembers.Count
       "Member names"     = $OutputNames}
    $Report.Add($ReportLine) 
}

Write-Host ("All done. {0} distribution lists analyzed. Output is in {1}" -f $DLs.Count, $DLCSVOutput)
$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation $DLCSVOutput

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
