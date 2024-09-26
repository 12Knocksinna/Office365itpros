# UpdateSubscribersInGroupsUsedByTeams.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateSubscribersInGroupsUsedByTeams.PS1
# Update the subscriber list for Teams-enabled groups so that members receive calendar updates

CLS
Write-Host "Finding team-enabled Groups to process..."
$Groups = Get-UnifiedGroup -Filter {ResourceProvisioningOptions -eq "Team"} -ResultSize Unlimited
$Groups = $Groups | Where-Object {$_.AutoSubscribeNewMembers -eq $False -Or $_.AlwaysSubscribeMembersToCalendarEvents -eq $False}

$Report = [System.Collections.Generic.List[Object]]::new() # Create output file 
#initialize progress bar
$ProgDelta = 100/($Groups.count)
$CheckCount = 0 ; $GroupNumber = 0 ; CLS
ForEach ($Group in $Groups) {
   $GroupNumber++
   $CheckCount += $ProgDelta
   $GroupStatus = "Processing " + $Group.DisplayName + " ["+ $GroupNumber +"/" + $Groups.Count + "]"
   Write-Progress -Activity "Updating subscriber information for group" -Status $GroupStatus -PercentComplete $CheckCount
   # Update group so that new members are added to the subscriber list and will receive calendar events
   Set-UnifiedGroup -Identity $Group.ExternalDirectoryObjectId -AutoSubscribeNewMembers:$True -AlwaysSubscribeMembersToCalendarEvents
   # Get current members and the subscribers list
   $Members = Get-UnifiedGroupLinks -Identity $Group.ExternalDirectoryObjectId -LinkType Member
   $Subscribers = Get-UnifiedGroupLinks -Identity $Group.ExternalDirectoryObjectId -LinkType Subscribers
   # Check each member and if they're not in the subscriber list, add them
   ForEach ($Member in $Members) {
     If ($Member.ExternalDirectoryObjectId -notin $Subscribers.ExternalDirectoryObjectId) { # Not in the list
    #    Write-Host "Adding" $Member.PrimarySmtpAddress "as a subscriber"
         Add-UnifiedGroupLinks -Identity $Group.ExternalDirectoryObjectId -LinkType Subscribers -Links $Member.PrimarySmtpAddress 
         $ReportLine = [PSCustomObject] @{
            Group      = $Group.DisplayName
            Subscriber = $Member.PrimarySmtpAddress
            Name       = $Member.DisplayName}
         $Report.Add($ReportLine) }   
     } #End ForEach
} #End ForEach
$Report | Export-CSV -NoTypeInformation c:\temp\SubscriberGroupUpdates.csv
Write-Host "All done. Details of updates are in c:\temp\SubscriberGroupUpdates.csv"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
