# TeamsCreationReportByEmail.PS1
# A script to locate Office 365 audit records for the creation of new Teams and report the fact via email.
# V2.0 22 Oct 2019
# Uses the Exchange Online PowerShell module...
$StartDate = (Get-Date).AddDays(-90); $EndDate = (Get-Date).AddDays(1)
#HTML header with styles
$htmlhead="<html>
     <style>
      BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    </style>"

#Header for the message
$HtmlBody = "<body>
     <h1>Teams Creation Report for teams created between $(Get-Date($StartDate) -format g) and $(Get-Date($EndDate) -format g)</h1>
     <p><strong>Generated:</strong> $(Get-Date -Format g)</p>  
     <h2><u>Details of Teams Created</u></h2>"
#Person to get the email
$EmailRecipient = "Tony.Redmond@Live.com"

If (-not $O365Cred) { #Make sure we have credentials
    $O365Cred = (Get-Credential)}
$MsgFrom = $O365Cred.UserName ; $SmtpServer = "smtp.office365.com" ; $SmtpPort = '587'

# Find records for team creation in the Office 365 audit log
Write-Host "Looking for Team Creation Audit Records..."
$Records = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "TeamCreated" -ResultSize 1000)
If ($Records.Count -eq 0) {
    Write-Host "No Team Creation records found." }
Else {
    Write-Host "Processing" $Records.Count "audit records..."
    $Report = @()
    ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      $O365Group = (Get-UnifiedGroup -Identity $AuditData.TeamName) # Need some Office 365 Group properties
      $ReportLine = [PSCustomObject][Ordered]@{
        TimeStamp      = Get-Date($AuditData.CreationTime) -format g
        User           = $AuditData.UserId
        Action         = $AuditData.Operation
        TeamName       = $AuditData.TeamName
        Privacy        = $O365Group.AccessType
        Classification = $O365Group.Classification
        MemberCount    = $O365Group.GroupMemberCount 
        GuestCount     = $O365Group.GroupExternalMemberCount
        ManagedBy      = $O365Group.ManagedBy}
     $Report += $ReportLine }
}

# Add details of each team
$Report | Sort TeamName -Unique | ForEach {
    $htmlHeaderTeam = "<h2>" + $_.TeamName + "</h2>"
    $htmlline1 = "<p>Created on <b>" + $_.TimeStamp + "</b> by: </b>" + $_.User + "</b></p>"
    $htmlline2 = "<p>Privacy: <b>" + $_.Privacy + "</b> Classification: <b>" + $_.Classification + "</b></p>"
    $htmlline3 = "<p>Member count: <b>" + $_.MemberCount + "</b> Guest members: <b>" + $_.GuestCount + "</b></p>"
    $htmlbody = $htmlbody + $htmlheaderTeam + $htmlline1 + $htmlline2 + $htmlline3 + "<p>"
}

# Finish up the HTML message body  
$HtmlMsg = "</body></html>" + $HtmlHead + $HtmlBody

# Construct the message parameters and send it off...
$MsgParam = @{
     To = $EmailRecipient
     From = $MsgFrom
     Subject = "Teams Creation Report"
     Body = $HtmlMsg
     SmtpServer = $SmtpServer
     Port = $SmtpPort
     Credential = $O365Cred }
Send-MailMessage @msgParam -UseSSL -BodyAsHTML ; Write-Host "Teams Creation Report sent by email to" $EmailRecipient
