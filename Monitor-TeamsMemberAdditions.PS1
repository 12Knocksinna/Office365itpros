# Monnitor-TeamMemberAdditions.PS1
# An example of using the unified audit log to monitor for additions of new members to Teams.
# Each new member added to some designated teams is checked to see if they come from a specific 
# department. If they do, they're removed. If not, details of the new members are posted to a 
# channel in the target team

# https://github.com/12Knocksinna/Office365itpros/blob/master/Monitor-TeamsMemberAdditions.PS1

# Need to connect to Exchange Online to get audit event information
If ($Null -eq (Get-ConnectionInformation)) {
    Connect-ExchangeOnline
}
# Connect to the Microsoft Graph
Connect-MgGraph -Scopes Directory.Read.All, Channel.Send.Message

# Define the set of Groups (Teams) to check
[array]$GroupstoCheck = "dc9e6f8b-6734-4180-af25-aa40fae79280", "107fe4dd-809c-4ec9-a3a1-ab88c96e0a5e"
# Define the set of departments that we don't want members from in the monitored teams
[array]$ExcludedDepartments = "Services", "Sales"

# Define identifiers for the target team and channel to post new member notifications to
$TargetTeamId = "107fe4dd-809c-4ec9-a3a1-ab88c96e0a5e"
$TargetTeamChannelId = "19:77ccce09d63741668605721946e45f61@thread.tacv2"

# Search for audit events created in the last three hours
$StartDate = (Get-Date).AddHours(-3)
$EndDate = (Get-Date).AddHours(1)
Write-Host "Searching for audit records..."
[array]$Records = Search-UnifiedAuditLog -Start $StartDate -End $EndDate -Operations "Add member to group" -Formatted -ResultSize 500
If (!($Records)) { Write-Host "No member additions to groups to check" ; break }

$MembersReport = [System.Collections.Generic.List[Object]]::new() 
Write-Host "Processing audit records..."
ForEach ($Rec in $Records) {
  $AuditData = $Rec.AuditData | ConvertFrom-Json
  $GroupId = $AuditData.ModifiedProperties | Where-Object {$_.Name -eq 'Group.ObjectID'} | Select-Object -ExpandProperty NewValue
  $GroupName = $AuditData.ModifiedProperties | Where-Object {$_.Name -eq 'Group.DisplayName'} | Select-Object -ExpandProperty NewValue
  $UserAdded = $AuditData.ObjectId
  $Actor = $Rec.UserIds
  If ($GroupId -in $GroupsToCheck) {
    $UserData = Get-MgUser -UserId $UserAdded -Property Id, displayName, department
    $ReportLine = [PSCustomObject]@{
      Team       = $GroupName
      User       = $UserAdded
      UserName   = $UserData.displayName
      UserId     = $UserData.Id
      Addedby    = $Actor
      Timestamp  = $Rec.CreationDate
      Department = $UserData.Department
      GroupId    = $GroupId
      Id         = ("{0}_{1}_{2}" -f $GroupName, $UserAdded, $Rec.CreationDate) }
    $MembersReport.Add($Reportline)
  }

}
# Sort to make sure that we only report unique records
$MembersReport = $MembersReport | Sort-Object Id -Unique

# Check each addition and either remove the new member if they come from a problem department
# or post a message to the target channel to welcome them to the team
ForEach ($R in $MembersReport) {
  If ($R.Department -in $ExcludedDepartments) {
     Write-Host ("User {0} with department {1} will be removed from team" -f $R.User, $R.Department) -ForegroundColor Red
     Remove-MgGroupMemberByRef -DirectoryObjectId $R.UserId -GroupId $R.GroupId 
  } Else {
    Write-Host ("Sending channel message about new team member {0}" -f $R.UserName) -ForegroundColor Yellow
    [string]$UserName = $R.UserName
    $HtmlContent = "<h1>New User Has Joined Our Team</h1>
               <h2>$UserName has joined this team</h2><p></p>
               <p>Please welcome <b>$UserName</b> to the team. They will bring great joy to all of us!</p>"
    New-MgTeamChannelMessage -TeamId $TargetTeamId `
     -ChannelId $TargetTeamChannelId `
     -Body @{Content = $HTMLContent; ContentType = "html"} `
     -Subject "New User Join Report" `
     -Importance "High"
  }
} 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.