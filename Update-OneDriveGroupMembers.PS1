# Update-OneDriveGroupMembers.PS1
# Sample script to illustrate the principal of using groups to drive the allocation of storage quotas to OneDrive accounts
# See https://office365itpros.com/2023/06/15/onedrive-storage-quota-manage/ for more information
# https://github.com/12Knocksinna/Office365itpros/blob/master/Update-OneDriveGroupMembers.PS1
#
# Check that we're connected to SharePoint Online
[array]$Modules = Get-Module
If ("Microsoft.Online.Sharepoint.PowerShell" -notin $Modules.Name) { 
   Write-Host "Please connect to the SharePoint Online management module and restart the script."; break 
}

# Connect to the Microsoft Graph
Connect-MgGraph -Scopes Directory.Read.All, Group.Read.All
# Read in allocated sizes for different group members
# The CSV file used should have columns for group, groupid, and allocation (in megabytes)
If (Test-Path "c:\temp\OneDriveAllocations.csv") {
   [array]$OneDriveAllocations = Import-CSV -Path c:\temp\OneDriveAllocations.csv 
} Else {
   Write-Host "Can't find group information to process - exiting" ; break 
}

# Figure out the service domain and what this means for the root of OneDrive account URLs
$ServiceDomain = (Get-MgOrganization).verifiedDomains | Where-Object {$_.IsInitial -eq $True} | Select-Object -ExpandProperty Name
$OneDriveDomain =("https://{0}-my.sharepoint.com/personal/" -f $ServiceDomain.Split(".")[0])

$Report = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Group in $OneDriveAllocations) {
   # Get group members and extract their user principal names
   [array]$GroupMembers = Get-MgGroupMember -GroupId $Group.GroupId
   $GroupMemberUPN = $GroupMembers.additionalProperties.userPrincipalName
   # Calculate the quota allocation for group members (in megabytes)
   [int]$NewAllocation = $Group.Allocation; $NewAllocation = $NewAllocation*1024
   ForEach ($Member in $GroupMemberUPN) {
     # Figure out the OneDrive site URL for the user's account
     $MemberDomain = $Member.Replace(".","_")
     $MemberDomain = $MemberDomain.Replace("@", "_")
     $OneDriveSiteURL = $OneDriveDomain + $MemberDomain
     $CurrentQuotaGB = ((Get-SPOSite -Identity $OneDriveSiteURL).StorageQuota/1024)
     If ($CurrentQuotaGB -lt $Group.Allocation) {
     # Current allocation is less than the quota assigned to the group, so increase it to the group allocation
        Write-Host ("Updating OneDrive storage allocation for account {0} to {1}" -f $Member, $Group.Allocation)
        Set-SPOSite -Identity $OneDriveSiteURL -StorageQuota $NewAllocation
        $ReportLine = [PSCustomObject]@{
          Guest       = $Member
          OneDrive    = $OneDriveSiteURL
          "Old Quota" = $CurrentQuotaGB
          "New Quota" = $Group.Allocation }
        $Report.Add($ReportLine) 
     } #End update quota
   } #End members of a group
} #End all groups

# Show what we did
$Report | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
