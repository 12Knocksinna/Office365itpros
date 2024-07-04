# Report-SharingAuditEvents.PS1
# Example of extracting details of sharing events from the audit log to see what's going on

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-SharingAuditEvents.PS1

$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "Loading Exchange Online Management module"
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

[array]$Operations = "SharingSet", "SecureLinkUsed", "SecureLinkCreated", "CompanyLinkCreated", "CompanyLinkUsed", "AnonymousLinkCreated"
[array]$Records = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date).AddDays(+1) `
   -Operations $Operations -ResultSize 2000 -Formatted -SessionCommand ReturnLargeSet
If (!($Records)) {
     Write-Host "No audit events for sharing found"
     Break
 }
$Records = $Records | Sort-Object Identity -Unique | Sort-Object { $_.CreationDate -as [datetime]} -Descending
$Organization = Get-OrganizationConfig | Select-Object -ExpandProperty DisplayName
$Report = [System.Collections.Generic.List[Object]]::new()
 
ForEach ($Rec in $Records) {
    $AuditData = ConvertFrom-Json $Rec.AuditData
    $EventData = $null; $SharingTarget = $null
    Switch ($AuditData.Operation) {
        "SecureLinkCreated" {   
            $File = $AuditData.SourceFileName
            $URL = $AuditData.ObjectId
            $EventData = $AuditData.EventData
            $LinkType = "Secure link"
        }
        "SharingSet" {
            If ($Rec.UserIds -eq "app@sharepoint") {
                $File = "Teams Meeting Recording"
            } Else {
                $File = $AuditData.SourceFileName
            }
            $URL = $AuditData.ObjectId
            $EventData = $AuditData.EventData
            $LinkType = "Secure link"
            $SharingTarget = $AuditData.TargetUserOrGroupName
        }
        "SecureLinkUsed" {
            $File = $AuditData.SourceFileName
            $URL = $AuditData.ObjectId
            $SharingTarget = $AuditData.UserId
            $LinkType = "Secure link"
        }
        "CompanyLinkCreated" {
            $File = $AuditData.SourceFileName
            $URL = $AuditData.ObjectId
            $EventData = $AuditData.EventData
            $LinkType = "Company link"
            $SharingTarget = $Organization
        }
        "CompanyLinkUsed" {
            $File = $AuditData.SourceFileName
            $URL = $AuditData.ObjectId
            $SharingTarget = $AuditData.UserId
            $LinkType = "Company link"
        }
        "AnonymousLinkCreated" {
            $File = $AuditData.SourceFileName
            $URL = $AuditData.ObjectId
            $EventData = $AuditData.EventData
            $LinkType = "Anonymous link"
        }
    }
 
    $ReportLine = [PSCustomObject]@{
        CreationDate    = (Get-Date $Rec.CreationDate -format "dd-MMM-yyyy HH:mm:ss")
        Operation   = $AuditData.Operation
        User        = $AuditData.UserId
        File        = $File
        URL         = $URL
        Target      = $SharingTarget
        EventData   = $EventData
        Linktype    = $LinkType

     }
    $Report.Add($ReportLine)
 }

 $Report | Out-GridView -Title "Audit events for sharing"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.