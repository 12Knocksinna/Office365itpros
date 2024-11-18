# Find-FilesWithDownGradedLabels.PS1
# Find files in SharePoint Online that have been downgraded in terms of the assigned sensitivity label by users

# V1.0 18-Nov-2024
# GitHub link: 

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
    }
    $Report.Add($ReportLine)
}




# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.