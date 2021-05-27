# ReportDLMembershipsCounts.PS1
# Report the membership and counts for distribution lists in Exchange Online
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportDLMembershipsCounts.PS1

Function Get-RecursiveAzureAdGroupMemberUsers{
[cmdletbinding()]
# Fron https://saemundsson.se/?p=734
param(
   [parameter(Mandatory=$True,ValueFromPipeline=$true)]
   $AzureGroup
)
    Begin{
        If(-not(Get-AzureADCurrentSessionInfo)){Connect-AzureAD}
    }
    Process {
        Write-Verbose -Message "Enumerating $($AzureGroup.DisplayName)"
        [array]$Members = Get-AzureADGroupMember -ObjectId $AzureGroup.ObjectId -All $true
        [array]$GroupMembers = $Members | ? {$_.ObjectType -eq 'Group'} | Select ObjectId, DisplayName, Mail, @{Name="Type"; e={$_.ObjectType}}
        
        [array]$UserMembers = $Members | ? {$_.ObjectType -ne 'Group'} | Select ObjectId, DisplayName, Mail, @{Name="Type"; e= {$_.UserType}}
        If($GroupMembers) {
           [array]$ExtractedMembers = $Members | Where-Object{$_.ObjectType -eq 'Group'} | ForEach-Object{ Get-RecursiveAzureAdGroupMemberUsers -AzureGroup $_}
           $UserMembers += $ExtractedMembers
        }
    }
    end {
        # Figure out what mix of user (including mail contacts and other recipients) and group members we have
        [array]$ReturnMembers = $Null
        If ($UserMembers) { $ReturnMembers = $UserMembers}
        If ($GroupMembers) { $ReturnMembers += $GroupMembers }

        # Remove duplicates if found in multiple DLs
        $ReturnMembers = $ReturnMembers | Sort ObjectId -Unique
        Return $ReturnMembers
    }
}

# Check we have the right modules loaded
$Modules = Get-Module
If ("ExchangeOnlineManagement" -notin  $Modules.Name) {Write-Host "Please connect to Exchange Online Management  before continuing...";break}
If ("AzureAD" -notin  $Modules.Name) {Write-Host "Please connect to Azure AD  before continuing...";break}

# Find all distribution lists
Write-Host "Finding Exchange Online Distribution Lists..."
# Find distribution lists, excluding room lists
$DLs = Get-DistributionGroup -ResultSize Unlimited -Filter {RecipientTypeDetails -ne "Roomlist"} | Select DisplayName, ExternalDirectoryObjectId
If (!($DLs)) { Write-Host "No distribution lists found... sorry! "; break }
 Else { Write-Host ("{0} distribution lists found" -f $DLs.count) }

$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($DL in $DLs) { 
  Write-Host "Processing" $DL.DisplayName ($Dl.ExternalDirectoryObjectId)
  [array]$Members =  Get-AzureADGroup -ObjectId $DL.ExternalDirectoryObjectId | Get-RecursiveAzureAdGroupMemberUsers
  If (!($Members)) { $CountOfMembers = 0 }
    Else { 
       [int]$CountOfMembers = $Members.Count
       [array]$TenantMembers = ( $Members | ? {$_.Type -eq "Member" })
       [array]$GroupMembers = ( $Members | ? {$_.Type -eq "Group" })
       [array]$GuestMembers = ( $Members | ? {$_.Type -eq "Guest" })
       [array]$OtherMembers = ( $Members | ? {$_.Type -eq $Null })
       $MemberNames = $Members.DisplayName -join ", "
    }

     $ReportLine = [PSCustomObject][Ordered]@{  
       DLName             = $DL.DisplayName
       "Members"          = $CountOfMembers
       "Tenant Users"     = $TenantMembers.Count
       "Groups"           = $GroupMembers.Count
       "Guest members"    = $GuestMembers.Count
       "Other Recipients" = $OtherMembers.Count
       "Member names"     = $MemberNames}

    $Report.Add($ReportLine) 

}
$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation c:\temp\DLMemberCounts.CSV
