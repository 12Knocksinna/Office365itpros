# FetchServiceMessagesGraph.ps1
# Fetch Service Messages from the Microsoft Graph
# https://github.com/12Knocksinna/Office365itpros/blob/master/FetchServiceMessagesGraph.ps1
# Updated 5-Feb-2024 to use the Microsoft Graph PowerShell cmdlets based on the service communications API

Connect-MgGraph -NoWelcome -Scopes ServiceMessage.Read.All

$CSVOutputFile = "c:\temp\MessageCenterPosts.csv"

# Fetch information from Graph
Write-Host "Fetching Microsoft 365 Message Center Notifications..."
[array]$MCPosts = Get-MgServiceAnnouncementMessage -Sort 'LastmodifiedDateTime desc' -All 

# And Report what we find
Write-Host "Generating a report..."
$Report = [System.Collections.Generic.List[Object]]::new() 
ForEach ($M in $MCPosts) {

   [array]$Services = $M.Services

   If ([string]::IsNullOrEmpty($M.ActionRequiredByDateTime)) { # No action required date
       $ActionRequiredDate = $null 
   }  Else {
      $ActionRequiredDate = Get-Date($M.ActionRequiredByDateTime) -format "dd-MMM-yyyy" 
   }
   # Get age of update
   $Age = New-TimeSpan($M.LastModifiedDateTime)
   # Trim the message text

   $Body = $M | Select-Object -ExpandProperty Body
   $HTML = New-Object -Com "HTMLFile"
   $HTML.write([ref]$body.content)
   $MCPostText = $HTML.body.innerText
   
   $ReportLine  = [PSCustomObject] @{          
      MessageId            = $M.Id
      Title                = $M.Title
      Workloads            = ($Services -join ",")
      Category             = $M.category
      'Start Time'         = Get-Date($M.StartDateTime) -format "dd-MMM-yyyy HH:mm"
      'End Time'           = Get-Date($M.EndDateTime) -format "dd-MMM-yyyy HH:mm"
      'Last Update'        = Get-Date($M.LastModifiedDateTime) -format "dd-MMM-yyyy HH:mm"
      'Action Required by' = $ActionRequiredDate
      MessageText           = $MCPostText
      Age                   = ("{0} days {1} hours" -f $Age.Days.ToString(), $Age.Hours.ToString())
      IsRead                = $M.ViewPoint.IsRead
      IsDismissed           = $M.ViewPoint.IsDismissed
      MinutesOld            = $Age.Minutes
   }
   $Report.Add($ReportLine) } 

$Report | Sort-Object {$_.'Last Update' -as [DateTime]} -Descending | `
   Select-Object MessageId, Title, Category, 'Last Update', 'Action Required By', Age | Out-GridView
$Report | Export-CSV -NoTypeInformation $CSVOutputFile

# Figure out delays
$DelayedPosts = $Report | Where-Object {$_.Title -like "*(Updated)*"}
[int]$TeamsDelays = 0; [int]$ExchangeDelays = 0; [int]$SharePointDelays = 0; [int]$OtherDelays = 0
[int]$StreamDelays = 0; [int]$PlannerDelays = 0; [int]$IntuneDelays = 0; [int]$OneDriveDelays = 0; [int]$WebAppsDelays = 0
[int]$VivaDelays = 0; [int]$M365AppsDelays = 0; [int]$M365SuiteDelays = 0; [int]$DefenderDelays = 0
ForEach ($Delay in $DelayedPosts) {
   Switch -wildcard ($Delay.Workloads) {
      "*Teams*" {
         $TeamsDelays++
      }
      "Exchange*" {
         $ExchangeDelays++
      }
      "SharePoint*" {
         $SharePointDelays++
      }
      "*Viva*" {
         $VivaDelays++
      }
      "Microsoft 365 apps" {
         $M365AppsDelays++
      }
      "*Stream*" {
         $StreamDelays++
      }
      "*Defender*" {
         $DefenderDelays++
      }
      "*Planner*" {
         $PlannerDelays++
      }
      "*OneDrive*" {
         $OneDriveDelays++
      }
      "Microsoft 365 suite" {
         $M365SuiteDelays++
      }
      "*Intune*" {
         $IntuneDelays++
      }
      "*web*" {
         $WebAppsDelays++
      }

      Default {
         $OtherDelays++
      }
   }
}

$PercentDelayed = ($DelayedPosts.count/$Report.count).toString('P')
Write-Host ("{0} message center posts analyzed - output is in {1}." -f $MCPosts.count, $CSVOutputFile)
Write-Host ""
Write-Host ("Number of delayed posts: {0} ({1})" -f $DelayedPosts.count, $PercentDelayed)
Write-Host ""
Write-Host "Delayed Posts by Workload"
Write-Host "-------------------------"
Write-Host ""

Write-Host ("Exchange Online           {0}" -f $ExchangeDelays)
Write-Host ("Intune                    {0}" -f $IntuneDelays)
Write-Host ("Microsoft 365 apps        {0}" -f $M365AppsDelays)
Write-Host ("Microsoft 365 suite       {0}" -f $M365SuiteDelays)
Write-Host ("Microsoft 365 for the web {0}" -f $WebAppsDelays)
Write-Host ("Microsoft Defender        {0}" -f $DefenderDelays)
Write-Host ("OneDrive for Business     {0}" -f $OneDriveDelays)
Write-Host ("Planner                   {0}" -f $PlannerDelays)
Write-Host ("SharePoint Online         {0}" -f $SharePointDelays)
Write-Host ("Stream                    {0}" -f $StreamDelays)
Write-Host ("Teams                     {0}" -f $TeamsDelays)
Write-Host ("Viva                      {0}" -f $VivaDelays)
Write-Host ("Other workloads           {0}" -f $OtherDelays)


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
