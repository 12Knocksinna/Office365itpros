# Get-ServiceHealthInformation.PS1
# Example of using the Graph Service Communication API with Microsoft Graph PowerShell SDK cmdlets
# V1.0 25-Jan-2024
# https://github.com/12Knocksinna/Office365itpros/blob/master/Get-ServiceHealthInformation.PS1

Function Add-MessageRecipients {
    # Function to build an addressee list to send email   
    [cmdletbinding()]
        Param(
        [array]$ListOfAddresses )
        ForEach ($SMTPAddress in $ListOfAddresses) {
            @{ emailAddress = @{address = $SMTPAddress}}    
        }
} 

# Connect to the Graph with the required permissions
Connect-MgGraph -Scopes ServiceHealth.Read.All, ServiceMessage.Read.All, Mail.Send -NoWelcome
       
# Recipient for the email sent at the end of the script - define the addresses you want to use here. They can be single recipients,
# distribution lists, or Microsoft 365 groups. Each recipient address is defined as an element in an array
[array]$EmailRecipient = "Email.Admins@office365itpros.com", "Kim.Akers@office365itpros.com"
# When run interactively, email will be sent from the account running the script. This is commented out for use with Azure Automation
# If used with the Mail.Send permission in an Azure Automation runbook, the sender can be any mailbox in the organization
$MsgFrom = (Get-MgContext).Account
# $MsgFrom = "Azure.Management.Account@office365itpros.com"

Write-Host "Looking for open service health advisories..."
[string]$RunDate = Get-Date -format 'dd-MMM-yyyy HH:mm'
# Find the set of service health items that have advisory status (the ones shown in the M365 admin center)
[array]$ServiceHealthItems = Get-MgServiceAnnouncementIssue -All `
    -Filter "classification eq 'Advisory' and status eq 'serviceDegradation'" | `
    Sort-Object {$_.LastModifiedDateTime -as [datetime]} -Descending

If ($ServiceHealthItems) {
    Write-Host ("{0} Service health items found..." -f $ServiceHealthItems.count)
} Else {
    Write-Host "No service health items found - exiting"
    Break
}

# What's happening in the tenant
$ServiceHealthItems | Format-Table LastModifiedDateTime, Status, ImpactDescription

# Get overall service health
[array]$ImportantServices = "Exchange", "Teams", "SharePoint", "OrgLiveID", "Planner", "microsoftteams", "O365Client"
[array]$ImportantServiceStatus = Get-MgServiceAnnouncementHealthOverview | Where-Object {$_.Id -in $ImportantServices}
$ImportantServiceStatus | Sort-Object Service | Format-Table Service, Status -AutoSize

$Report = [System.Collections.Generic.List[Object]]::new() 
ForEach ($Issue in $ServiceHealthItems) {
    $IssueId = $Issue.Id
    # Get the posts from the issue
    $Posts = $Issue  | Select-Object -ExpandProperty Posts | Sort-Object {$_.CreatedDateTime -as [datetime]} -Descending
    $DaysIssueOld = (New-TimeSpan -Start $Issue.StartDateTime).Days  
    $HoursSinceUpdate = (New-TimeSpan -Start $Issue.LastModifiedDateTime).Hours
    $DataLine = [PSCustomObject] @{
        Id                  = $IssueId
        FeatureGroup        = $Issue.FeatureGroup
        Title               = $Issue.Title
        Description         = $Issue.ImpactDescription
        Status              = $Issue.Status
        'Last Note'         = $Posts[0].description.content
        'Start'             = $Issue.StartDateTime
        'Last Update'       = $Issue.LastModifiedDateTime
        'Days old'          = $DaysIssueOld
        'Hours since update'= $HoursSinceUpdate
    }
    $Report.Add($DataLine)
}
    
# Email the report
Write-Host ("All done - emailing details to {0}" -f ($EmailRecipient -join ", "))
$ToRecipientList   = @( $EmailRecipient )

[array]$MsgToRecipients = Add-MessageRecipients -ListOfAddresses $ToRecipientList
$MsgSubject = ("Current tenant Service Health advisories as at {0}" -f $RunDate)

$OrgName = (Get-MgOrganization).DisplayName
$HTMLHead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.StateOrange{background: #FFA500;}
	   td.StateRed{background: #FF0000;}
       td.StateYellow{background: #FFFF00;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Open Service Health Advisories</h1></p>
           <p><h2><b>For the " + $Orgname + " tenant</b></h2></p>
           <p><h3>Generated: " + $RunDate + "</h3></p><p>Details of the known service health advisories.</p></div>"
$Line = ("<p>There are currently <b>{0}</b> service advistories open.</p>" -f $ServiceHealthItems.count)
$HTMLHead = $HTMLHead + $Line

$HtmlBody = $Report | Select-Object Id, FeaturGroup, Title, Status, 'Last Note', 'Start', 'Last Update', 'Hours since update', 'Days old' | ConvertTo-Html -Fragment 

# This section highlights whether a conditional access policy is enabled or disabled in the summary.
# Idea from https://stackoverflow.com/questions/37662940/convertto-html-highlight-the-cells-with-special-values
# First, convert the CA Policies report to HTML and then import it into an XML structure
$HTMLTable = $Report | ConvertTo-Html -Fragment
[xml]$XML = $HTMLTable
# Create an attribute class to use, name it, and append to the XML table attributes
$TableClass = $XML.CreateAttribute("class")
$TableClass.Value = "State"
$XML.table.Attributes.Append($TableClass) | Out-Null
# Conditional formatting for the table rows. The number of available units is in table row 6, so we update td[5]
ForEach ($TableRow in $XML.table.SelectNodes("tr")) {
    # each TR becomes a member of class "tablerow"
    $TableRow.SetAttribute("class","tablerow")
    [int]$HoursSinceUpdate = 0; [int]$DaysSinceStart = 0
    If ($TableRow.td.count -eq 10) {
        # If a valid table row, extract the hours since the last update and the days since the incident start
        [int]$HoursSinceUpdate = $TableRow.td[9]
        [int]$DaysSinceStart = $TableRow.td[8]
        # Orange state for incident that's between 12 and 23 hours old
        If (($TableRow.td) -and ($HoursSinceUpdate -ge 12 -and $HoursSinceUpdate -lt 24))  {
            $TableRow.SelectNodes("td")[9].SetAttribute("class","StateOrange")
        }
        # Red state for incidents open longer than 24 hours
        If (($TableRow.td) -and ($HoursSinceUpdate -ge 24)) {
            $TableRow.SelectNodes("td")[9].SetAttribute("class","StateRed")
        }
        # Days old is greater than 7
        If (($TableRow.td) -and ($DaysSinceStart -gt 7)) {
            $TableRow.SelectNodes("td")[8].SetAttribute("class","StateYellow")
        }
    }
}
# Wrap the output table with a div tag
$HTMLBody = [string]::Format('<div class="tablediv">{0}</div>',$XML.OuterXml)
# Put the message content together
$HTMLMsg = "</body></html><p>" + $HTMLHead + $HTMLBody + "<p>"

# Construct the message body
$MsgBody = @{
  Content = "$($HTMLMsg)"
  ContentType = 'html'  
}

$Message =  @{subject           = $MsgSubject}
$Message += @{toRecipients      = $MsgToRecipients}  
$Message += @{body              = $MsgBody}
$Params   = @{'message'         = $Message}
$Params  += @{'saveToSentItems' = $True}
$Params  += @{'isDeliveryReceiptRequested' = $True}

# And send the message using the parameters that we've filled in
Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params
Write-Output ("Message containing information about current open service advisories sent to {0}!" -f ($EmailRecipient -join ", "))

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
