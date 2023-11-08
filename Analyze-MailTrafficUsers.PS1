# Analyze-MailtrafficUsers.PS1
# An example of using the Exchange Online message trace log to analyze inbound and outbound traffic
# for a Microsoft 365 tenant
# https://github.com/12Knocksinna/Office365itpros/blob/master/Analyze-MailTrafficUsers.PS1

# Check if we can run an Exchange Online cmdlet. If we can, go on, else connect to Exchange Online
If ($Null -eq (Get-ConnectionInformation)) {
    Connect-ExchangeOnline
}

[int]$i = 1
$MoreMessages = $True
[array]$Messages = $Null
# Message trace date is kept for a maximum of 10 days
$StartDate = (Get-Date).AddDays(-10)
$EndDate = (Get-Date).AddDays(1)

Write-Host ("Message trace data will be analyzed between {0} and {1}" -f $StartDate, $EndDate)
While ($MoreMessages -eq $True) {
    Write-Host ("Fetching message trace data to analyze - Page {0}" -f $i)
    [array]$MessagePage = Get-MessageTrace -StartDate $StartDate -EndDate $EndDate -PageSize 1000 -Page $i -Status "Delivered"
    If ($MessagePage)  {
        $i++
        $Messages += $MessagePage
    } Else {
        $MoreMessages = $False
    }
}

# Remove Exchange Online public folder hierarchy synchronization messages
$Messages = $Messages | Where-Object {$_.Subject -NotLike "*HierarchySync*"}

# Get a list of vertified domains for the tenant so we can differentiate between internal
# and external email based on recipient address
[array]$Domains = Get-AcceptedDomain | Select-Object -ExpandProperty DomainName

# Fetch a list of user and shared mailboxes to process
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -ResultSize Unlimited | Sort-Object DisplayName
$MessageReport = [System.Collections.Generic.List[Object]]::new() 

ForEach ($User in $Mbx) {
    Write-Host ("Processing email for {0}" -f $User.DisplayName)
    # Get messages sent by the user
    [array]$UserMessages = $Messages | Where-Object {$_.SenderAddress -eq $User.PrimarySmtpAddress}
    If ($UserMessages) {
    # We’ve found some messages to process, so let’s do that
    [int]$ExternalEmail = 0; [int]$InternalEmail = 0; [array]$ExternalDomains = $Null
    ForEach ($M in $UserMessages) {
        $MsgRecipientDomain = $M.RecipientAddress.Split('@')[1]    
        If ($MsgRecipientDomain -in $Domains) {
            $InternalEmail++ 
        } Else {
            $ExternalEmail++
            $ExternalDomains += $MsgRecipientDomain
        }
      } # End Foreach message
      $ExternalDomains = $ExternalDomains | Sort-Object -Unique
      $PercentInternal = "N/A"; $PercentExternal = "N/A"
      If ($InternalEmail -gt 0) {
         $PercentInternal = ($InternalEmail/($UserMessages.count)).toString("P") }
      If ($ExternalEmail -gt 0) {
         $PercentExternal = ($ExternalEmail/($UserMessages.count)).toString("P") }

     Switch ($User.RecipientTypeDetails) {
        "UserMailbox" { $Type = "User"}
        "SharedMailbox" { $Type = "Shared"}
     }    
  
      $ReportLine = [PSCustomObject]@{
        User               = $User.UserPrincipalName
        Name               = ("{0} ({1})" -f $User.DisplayName, $Type)
        Internal           = $InternalEmail
        "% Internal"       = $PercentInternal
        External           = $ExternalEmail 
        "% External"       = $PercentExternal
        "External Domains" = $ExternalDomains -Join ", "
       }
      $MessageReport.Add($ReportLine)
    } # End if user (has some messages)
  } # End ForEach mailboxes
  
  # Generate a report
  $ReportFile = "c:\temp\UserMailTraffic.html"
  $CSVFile = "c:\temp\UserMailTraffic.csv"
  $HtmlHead="<html>
         <style>
         BODY{font-family: Arial; font-size: 8pt;}
         H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
         H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
         H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
         TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
         TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
         TD{border: 1px solid #969595; padding: 5px; }
         td.pass{background: #B7EB83;}
         td.warn{background: #FFF275;}
         td.fail{background: #FF2626; color: #ffffff;}
         td.info{background: #85D4FF;}
         </style>
         <body>
             <div align=center>
             <p><h1>Message Traffic User Analysis</h1></p>
             <p><h3>Generated: " + (Get-Date -format 'dd-MMM-yyyy hh:mm') + " for " + (Get-OrganizationConfig | Select-Object -ExpandProperty DisplayName) + "</h3></p></div>"
          
  $HtmlBody = $MessageReport | ConvertTo-Html -Fragment 
 "</body></html><p>" + $HtmlHead + $Htmlbody + "<p>" | Out-File $ReportFile  -Encoding UTF8

$MessageReport | Export-Csv -NoTypeInformation $CSVFile 
Write-Host ("All done - the HTML report is available in {0} and CSV in {1}" -f $ReportFile, $CSVFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.