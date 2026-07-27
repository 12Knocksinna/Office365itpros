#Requires -Modules Microsoft.Graph.Groups, Microsoft.Graph.Planner

<#
     
    Report-ObsoletePlannerTasks.PS1
    Inventories Planner plans that might be affected by a Microsoft Purview Data Lifecycle
    Management retention/deletion policy (Roadmap ID 486828, MC1384412).
    Flags completed tasks older than a retention threshold, aged on either
    creation date or last-modified date.
    
    GitHub link:
    V1.0 25-Jul-2025

    .PARAMETER RetentionYears
        Retention period in years. Completed tasks older than
        (today - RetentionYears) are counted as "would be deleted".

    .PARAMETER DateBasis
        Which task date the retention threshold is measured against:
        'Created'      -> CreatedDateTime (default)
        'Completed' -> CompletedDateTime 
#>

[CmdletBinding()]
param(
    [double]$RetentionYears = 7,

    [ValidateSet('Created','Completed')]
    [string]$DateBasis = 'Created'
)


function Get-TaskAgeDate {
    param($Task, [string]$Basis)

    # Return the date to check against
    If ($Basis.tolower() -eq 'completed') {
        # Only available for completed tasks, otherwise returns null
        return $Task.CompletedDateTime
    }
    return $Task.CreatedDateTime
}

# Use app-only mode (connect with certificate) to be able to read all plans in the tenant. Delegated access only reads the plans that
# the signed-in user has access to. See https://learn.microsoft.com/en-us/graph/api/resources/planner-overview?view=graph-rest-1.0#permissions for more information.
Connect-MgGraph -Scopes 'GroupMember.Read.All','Tasks.Read.All' -NoWelcome

$Cutoff = (Get-Date).AddYears(-1 * $RetentionYears)
Write-Host ("Retention window: {0} year(s). Basis: {1}. Cutoff: {2}" -f $RetentionYears, $DateBasis, $Cutoff.ToString('dd-MMM-yyyy')) -ForegroundColor Cyan

# Get collection of Microsoft 365 Groups
[array]$Groups = Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified')" -Property Id,DisplayName,CreatedDateTime 
If (!$Groups) {
    Write-Warning "No Microsoft 365 Groups found in the tenant."
    return
}
# Graph API does not support sorting by DisplayName, so we sort the collection in PowerShell
$Groups = $Groups | Sort-Object DisplayName

$Report = [System.Collections.Generic.List[object]]::new()
ForEach ($Group in $Groups) {
    Write-Host ("Scanning group '{0}' for plans" -f $Group.DisplayName) -ForegroundColor Green
    try {
        [array]$Plans = Get-MgGroupPlannerPlan -GroupId $Group.Id -All -ErrorAction Stop
        If ($Plans) {
            Write-Host ("Found {0} plan(s) in group '{1}'" -f $Plans.Count, $Group.DisplayName) -ForegroundColor Yellow
        }
    } catch {
        Write-Warning "Could not read plans for group '$($Group.DisplayName)': $($_.Exception.Message)"
        continue
    }

    ForEach ($Plan in $Plans) {
        [array]$Tasks = Get-MgPlannerPlanTask -PlannerPlanId $Plan.Id -All -ErrorAction SilentlyContinue

        $Completed  = @($Tasks | Where-Object PercentComplete -eq 100)
        $InProgress = @($Tasks | Where-Object PercentComplete -eq 50)
        $NotStarted = @($Tasks | Where-Object PercentComplete -eq 0)

        # Completed tasks whose ageing date is older than the threshold
        $Expiring = @($Completed | Where-Object {
            $ageDate = Get-TaskAgeDate -Task $_ -Basis $DateBasis
            $ageDate -and $ageDate -lt $Cutoff
        })

        # Oldest ageing date among completed tasks (respects chosen basis)
        $OldestCompleted = $Completed | ForEach-Object { Get-TaskAgeDate -Task $_ -Basis $DateBasis } | Sort-Object | Select-Object -First 1
       
        If ($OldestCompleted) {
            $OldestCompleted = $OldestCompleted.ToString('dd-MMM-yyyy')
        }

        $ReportLine = [PSCustomObject][Ordered]@{ 
            GroupName              = $Group.DisplayName
            GroupId                = $Group.Id
            PlanTitle              = $Plan.Title
            PlanId                 = $Plan.Id
            TaskCount              = $Tasks.Count
            NotStartedCount        = $NotStarted.Count
            InProgressCount        = $InProgress.Count
            CompletedCount         = $Completed.Count
            ExpiringCompletedCount = $Expiring.Count
            DateBasis              = $DateBasis
            OldestCompletedDate    = $OldestCompleted
        }
        $Report.Add($ReportLine)
    }
}

$Report | Sort-Object ExpiringCompletedCount -Descending | Format-Table GroupName, PlanTitle, TaskCount, CompletedCount, ExpiringCompletedCount, DateBasis, OldestCompletedDate -AutoSize

# Generate report file
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $OutputXLSXFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Planner Task Analysis.xlsx"
    Remove-Item -Path $OutputXLSXFile -ErrorAction SilentlyContinue
    $Report | Export-Excel -Path $OutputXLSXFile -WorksheetName "Planner Task Analysis" -Title ("Planner Task Analysis {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "PlannerTasks" 
} Else {
    $OutputCSVFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Planner Task Analysis.csv"
    Remove-Item -Path $OutputCSVFile -ErrorAction SilentlyContinue
    $Report | Export-Csv -Path $OutputCSVFile -NoTypeInformation -Encoding Utf8
}
  
If ($ExcelGenerated) {
  Write-Host ("An Excel report is available in {0}" -f $OutputXLSXFile)
} Else {
  Write-Host ("A CSV report is available in {0}" -f $OutputCSVFile)
}

$totalExpiring = ($Report | Measure-Object ExpiringCompletedCount -Sum).Sum
Write-Host ("Total completed tasks older than {0} year(s) by {1} date: {2}" -f `  $RetentionYears, $DateBasis, $totalExpiring) -ForegroundColor Yellow

# Send email with the results - define your own email address here and the address to receive the report.
$MsgFrom = "PowerShell.Reports@office365itpros.com"
$DestinationEmailAddress = "User.Administration@office365itpros.com"
# Build the array of a single TO recipient detailed in a hash table - change this to the appropriate recipient for your tenant
$ToRecipient = @{}
$ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
[array]$MsgTo = $ToRecipient
# Define the message subject
$MsgSubject = "Important: Shared Mailboxes that potentially need Microsoft Defender for Office 365 Licenses"
# Create the HTML content
$HtmlMsg = ("</body></html><p>The following <b>{0} Shared Mailboxes need Microsoft Defender for Office 365 Licenses</b> based on external email traffic received over the last 90 days.</p>" -f $Report.Count)
$HtmlMsg += "<p>The report also checked shared mailboxes to detect online Teams meetings organized by the mailbox. If this situation exists, you should assign a license containing an OneDrive for Business service plan to the mailbox to ensure meeting recordings are retained. </p>"
$HtmlMsg += "<p>Please review the information at your convenience</p>"
# Construct the message body 	
$HtmlBody = $Report | ConvertTo-Html -As Table -PostContent "</body></html>" -Fragment
$HtmlMsg = $HtmlMsg + $HtmlBody
$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlMsg)")
$MsgBody.Add('ContentType','html')
# Build the parameters to submit the message
$Message = @{}
$Message.Add('subject', $MsgSubject)
$Message.Add('toRecipients', $MsgTo)
$Message.Add('body', $MsgBody)
$Message.Add("attachments", $MsgAttachments)

$EmailParameters = @{}
$EmailParameters.Add('message', $Message)
$EmailParameters.Add('saveToSentItems', $true)
$EmailParameters.Add('isDeliveryReceiptRequested', $true)
Write-Output ""
Write-Output "Sending email..."
Write-Output ""

# Send the message
Try {
    Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters -ErrorAction Stop
    Write-Output ("Shared Mailboxes potential licensing report emailed to {0}" -f $ToRecipient.emailAddress.address)
} Catch {
    Write-Output "Unable to send email"
    Write-Output $_.Exception.Message
}
# Optional: export for your compliance documentation
# $results | Export-Csv -Path .\PlannerRetentionScope.csv -NoTypeInformation -Encoding UTF8

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.