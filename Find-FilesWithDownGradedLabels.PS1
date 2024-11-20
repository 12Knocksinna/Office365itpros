# Find-FilesWithDownGradedLabels.PS1
# Find files in SharePoint Online that have been downgraded in terms of the assigned sensitivity label by users

# V1.0 18-Nov-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-FilesWithDownGradedLabels.PS1

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) { 
   Write-Host "Connecting to Exchange Online..."
   Connect-ExchangeOnline -ShowBanner:$False -SkipLoadingCmdletHelp
   Connect-IPPSSession -ShowBanner:$False
}
# Define the operations to look for
[array]$Operations = 'FileSensitivityLabelRemoved', 'FileSensitivityLabelChanged'
$StartDate = (Get-Date).AddDays(-180)
$EndDate = (Get-Date).AddDays(1)
# Averahe rate of downgrades per day that is acceptable
$DownGradeThreshold = 20

Write-Host "Looking for audit records for sensitivity label changes in SharePoint Online..."
[array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Formatted -Operations $Operations -ResultSize 5000 -SessionCommand ReturnLargeSet
If (!($Records)) {
    Write-Host "No audit records found"
    Break
}   Else {
    $Records = $Records | Sort-Object Identity -Unique | Sort-Object {$_.CreationDate -as [DateTime]}
}

# Get labels
Write-Host "Retrieving sensitivity labels..."
[array]$Labels = Get-Label | Where-Object {$_.ContentType -like "*File*"} | Select-Object DisplayName, ImmutableId, Priority
$LabelDisplayNameHash = @{}
$Labels.ForEach( {$LabelDisplayNameHash.Add([String]$_.ImmutableId, $_.DisplayName) } )
$LabelPriorityHash = @{}
$Labels.ForEach( {$LabelPriorityHash.Add([String]$_.ImmutableId, $_.Priority) } )


# Output report
$Report = [System.Collections.Generic.List[Object]]::new()
Write-Host "Analyzing audit records..."
ForEach ($Rec in $Records) {
    $LabelDisplayName = $null; $LabelAction = $null; $OldLabelDisplayName = $null; [int]$Priority = $null; [int]$OldPriority = $null
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    # Figure out what happened... Note thatg the $AuditData.Sensitivitylabeleventdata.LabelEventType contains codes that are not documented
    # to indicate the kind of change. 1 = Upgraded, 2 = Downgraded, 3 = Removed. We use the priority of the labels to determine the change
    Switch ($Rec.Operations) {
        'FileSensitivityLabelRemoved' {
            $LabelAction = "Removed"
            $OldLabelDisplayName = $LabelDisplayNameHash[$AuditData.SensitivityLabelEventData.OldSensitivityLabelId]
            If ($null -eq $OldLabelDisplayName) {
                $OldLabelDisplayName = "Unknown"
            }   
        }
        'FileSensitivityLabelChanged' {
            $LabelDisplayName = $LabelDisplayNameHash[$AuditData.SensitivityLabelEventData.SensitivityLabelId]
            $OldLabelDisplayName = $LabelDisplayNameHash[$AuditData.SensitivityLabelEventData.OldSensitivityLabelId]
            [int]$Priority = $LabelPriorityHash[$AuditData.SensitivityLabelEventData.SensitivityLabelId]
            [int]$OldPriority = $LabelPriorityHash[$AuditData.SensitivityLabelEventData.OldSensitivityLabelId]
            $LabelAction = $null
            If ($Priority -gt $OldPriority) {
                $LabelAction = "Upgraded"
            } ElseIf ($Priority -lt $OldPriority) {
                $LabelAction = "Downgraded"
            } ElseIf ($OldPriority -eq $Priority) {
                $LabelAction = "No change"
            }
        }
    }
    $ReportLine = [PSCustomObject]@{
        Timestamp       = Get-Date $Rec.CreationDate -format 'dd-MMM-yyyy HH:mm:ss'
        User            = $AuditData.UserId
        Operation       = $AuditData.Operation
        File            = $AuditData.SourceFileName
        Folder          = $AuditData.SourceRelativeUrl
        LabelAction     = $LabelAction
        OldLabel        = $OldLabelDisplayName
        OldPriority     = $OldPriority
        NewLabel        = $LabelDisplayName
        NewPriority     = $Priority
        URL             = $AuditData.ObjectId
        EventCode       = $AuditData.Sensitivitylabeleventdata.LabelEventType
    }
    $Report.Add($ReportLine)
}

$Report = $Report | Sort-Object {$_.TimeStamp -as [DateTime]} -Descending
$EarliestRecord = $Report[-1].TimeStamp
$LatestRecord = $Report[0].TimeStamp
$ReviewPeriodDays = (New-TimeSpan -Start $EarliestRecord -End $LatestRecord).Days
[array]$DowngradedFiles = $Report | Where-Object {$_.LabelAction -eq 'Downgraded'}
[array]$UpgradedFiles = $Report | Where-Object {$_.LabelAction -eq 'Upgraded'}
[array]$RemovedFiles = $Report | Where-Object {$_.LabelAction -eq 'Removed'}
[array]$RemovedAndDowngraded = $RemovedFiles + $DowngradedFiles
[array]$UsersWhoDowngrade = $DowngradedFiles | Where-Object {$_.User -ne 'SHAREPOINT\System'} | Group-Object User -NoElement | Sort-Object Count -Descending  
$AverageDowngradeRate = [Math]::Round($RemovedAndDowngraded.Count / $ReviewPeriodDays, 2)

Write-Host ""
Write-Host "The following accounts have downgraded or removed sensitivity labels applied to documents"
Write-Host "-----------------------------------------------------------------------------------------"
$UsersWhoDowngrade | Format-Table Name, Count -AutoSize
Write-Host ""

Write-Host ("Over the recorded period from {0} to {1}, {2} audit records found" -f $EarliestRecord, $LatestRecord, ($RemovedAndDowngraded.Count + $UpgradedFiles.count))
Write-Host ("Audit records for sensitivity label changes found {0} downgrades, {1} removals, and {2} upgradeds" -f $DowngradedFiles.Count, $RemovedFiles.count, $UpgradedFiles.Count)
If ($AverageDowngradeRate -lt $DowngradeThreshold) {
    Write-Host ("An average rate of {0} label downgrades and removals per day is not an issue" -f $AverageDowngradeRate)
} Else {
    Write-Host ("An average rate of {0} labnel downgrades and removals per day is of some concern" -f $AverageDowngradeRate)
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.