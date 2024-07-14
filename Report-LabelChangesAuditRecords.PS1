# Report-LabelChangesAuditRecords.PS1
# Example of how to use Microsoft 365 audit records to track changes made in sensitivity labels and policies
# and retention labels and policies.
# V1.0 14-Jul-2024
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-LabelChangesAuditRecords.PS1

$Modules = Get-Module | Select-Object Name
If ("ExchangeOnlineManagement -notin $Modules") {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
    Connect-IPPSSession
}
# Connect to the Graph to resolve service principals
Connect-MgGraph -Scopes Application.Read.All

# It seems like only one service principal (75367c9a-9a5b-41be-840f-ee9ee448c1f5, "Microsoft Exchange Online Protection")
# is ever found in the audit records found by this search. If this is the case, it would be possible to hard code the
# check for the service principal identifier and avoid the overhead of running the Get-MgServicePrincipal cmdlet

# Get sensitivity labels and build a hash table to lookup labels
$Labels = Get-Label
$LabelHash = @{}
ForEach ($L in $Labels) {
    $LabelHash.Add($L.ImmutableId, $L.Name)
}

Write-Host "Looking for audit records..."
# Find evidence that someone has been messing with sensitivity label and retention label policies
[array]$Operations = "Set-RetentionCompliancePolicy", "Set-Label", "Set-LabelPolicy", `
    "Set-RetentionComplianceRule", "Update label."
[array]$Records = Search-UnifiedAuditLog -StartDate (Get-date).AddDays(-90) -EndDate (Get-Date).AddDays(1) -ResultSize 1000 -Formatted -Operations $Operations -SessionCommand ReturnLargeSet
If (!($Records)) { 
    Write-Host "No audit records found"; break 
} Else {
    $Records = $Records | Sort-Object Identity -Unique | Sort-Object {$_.CreationDate -as [datetime]} -Descending
    Write-Host ("{0} audit records found" -f $Records.Count)
}
$Report = [System.Collections.Generic.List[Object]]::new()
Write-Host "Processing audit records..."
ForEach ($Rec in $Records) {
    $Data = $null; $LabelName = $null
    $AuditData = $Rec.AuditData | ConvertFrom-Json

    Switch ($Rec.Operations) {
        "Set-RetentionCompliancePolicy" { # Update retention policy
            $Data = $AuditData.Parameters
        }
        "Set-RetentionComplianceRule" { # Update rule for retention policy
            $Data = $AuditData.Parameters
        }
        "Set-Label" { # Update sensitivity label
            $Data = $AuditData.Parameters
            $LabelName = $LabelHash[$AuditData.ObjectId]
            If ($null -eq $LabelName) {
                $LabelName = "Sensitivity label not found in tenant"
            }
        }
        "Set-LabelPolicy" { # Update sensitivity label policy
            $Data = $AuditData.Parameters
        }   
        "Update label." { # Update retention label
            $LabelName = $AuditData.target.id[3]
            $Data = $AuditData.ObjectId
        }
    }

    If ($Rec.UserIds -like "*ServicePrincipal*") {
        $SPStart = $Rec.UserIds.IndexOf("_")+1
        $SPId = $Rec.UserIds.SubString(($SPStart), $Rec.UserIds.Length - ($SPStart))

        $User = (Get-MgServicePrincipal -ServicePrincipalId $SPId).displayName
    } Else {
        $User = $Rec.UserIds
    }

    $ReportLine = [PSCustomObject] @{ 
        User        = $User
        Operation   = $Rec.Operations
        Timestamp   = (Get-Date $Rec.CreationDate -format 'dd-MMM-yyyy HH:mm:ss')
        Data        = $Data
        Label       = $LabelName    
    }
    $Report.Add($ReportLine)
}

$Report | Out-Gridview -Title 'Audit Records for Label Changes'

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.