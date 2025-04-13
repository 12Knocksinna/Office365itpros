# Add-OwnerToApps.PS1
# Look for Entra ID registered apps without owners and add the last person to update the app or its creator as the owner. The information
# about the user who last managed apps comes from audit records.

# V1.0 12-Apr-2025
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Add-OwnerstoApps.PS1

Connect-MgGraph -Scopes Application.ReadWrite.All, User.ReadBasic.All, Directory.Read.All

$Modules = Get-Module | Select-Object -ExpandProperty Name
If ("ExchangeOnlineManagement" -notin $Modules) {
    Write-Host "connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop
}

Write-Host "Looking for audit records for application management..."
$Operations = "Update application.", "Add application."
[array]$Records = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-180) -EndDate (Get-Date) -RecordType AzureActiveDirectory `
    -Operations $Operations -ResultSize 5000 -SessionCommand ReturnLargeset

If (!$Records) {
    Write-Host "No audit records found"
    Break
} Else {
    Write-Host ("{0} records found" -f $Records.count)
}
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    $ReportLine = [PSCustomObject][Ordered]@{ 
        Timestamp       = $Rec.CreationDate
        Operation       = $Rec.Operations
        UPN             = $Rec.UserIds
        UserId          = $Auditdata.Actor[4].Id
        AppId           = $Auditdata.Target[4].Id
        Appname         = $Auditdata.Target[3].Id
        ObjectId        = $AuditData.ObjectId.Split("_")[1]  
    }
    $Report.Add($ReportLine)
}
# Make sure the output is sorted by date
$Report = $Report | Sort-Object {$_.CreationDate -as [datetime]} -Descending
Write-Host ("Audit records found for {0} applications" -f $UniqueApps.count)

# Group by AppId and select the most recent record for each application
[array]$UniqueApps = $Report | Group-Object -Property AppId | ForEach-Object {
    $_.Group | Sort-Object -Property Timestamp -Descending | Select-Object -First 1
}

# Find the set of app registrations for the tenant
[array]$Apps = Get-MgApplication -All -Property displayName, Id, AppId, Owners, CreatedDateTime, SigninAudience, PublisherDomain
Write-Host ("{0} registered applications found in Entra ID - now checking each app" -f $Apps.count)

# If you want just apps with no owners, use this command:
# [array]$Apps = Get-MgApplication -Filter "owners/`$count eq 0" -CountVariable CountVar -ConsistencyLevel eventual -All -Property displayName, Id, AppId, Owners, CreatedDateTime, SigninAudience, PublisherDomain

[int]$UpdatedApps = 0
$AppReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($App in $Apps) {
    $AppOwners = $null; $AppOwnersReport = $null
    [array]$AppOwners = Get-MgApplication -ApplicationId $App.Id -ExpandProperty Owners | Select-Object -ExpandProperty Owners 
    If (-not $AppOwners) {
        Write-Host ("No owners found for application {0} ({1})" -f $App.DisplayName, $App.AppId)
        $LastUpdateRecord = $UniqueApps | Where-Object { $_.AppId -eq $App.AppId }
        If ($LastUpdateRecord) {
            $UserId = $LastUpdateRecord.UserId
            $UPN = $LastUpdateRecord.UPN
            Write-Host ("Adding {0} as owner for application {1}" -f $UserId, $App.DisplayName)
            Try {
                $OwnerRef = ("https://graph.microsoft.com/v1.0/directoryObjects/{0}" -f $UserId)
                $OwnerId = @{}
                $OwnerId.Add("@odata.id", $OwnerRef)
                New-MgApplicationOwnerByRef -ApplicationId $App.Id -BodyParameter $OwnerId
                Write-Host ("Successfully added {0} as owner for application {1}" -f $UPN, $App.DisplayName) -ForegroundColor Yellow
                $AppOwnersReport = $UPN
                $UpdatedApps++
            } Catch {
                Write-Host ("Failed to add owner for application {0}: {1}" -f $App.DisplayName, $_.Exception.Message)
            }
        } Else {
            Write-Host ("No update record found for application {0}" -f $App.DisplayName)
        }
    } Else {
        $AppOwnersReport = $AppOwners.additionalProperties.userPrincipalName -join ", "
        Write-Host ("Owners exist for application {0} ({1})" -f $App.DisplayName, $App.AppId) -ForegroundColor Green
    }
    $AppReportLine = [PSCustomObject]@{
        AppId               = $App.AppId
        AppName             = $App.DisplayName
        Owners              = $AppOwnersReport
        Created             = $App.CreatedDateTime
        SigninAudience      = $App.SigninAudience
        PublisherDomain     = $App.PublisherDomain
    }
    $AppReport.Add($AppReportLine)
}

Write-Host ""
Write-Host ("Successfully updated {0} apps with owner details" -f $UpdatedApps) 
Write-Host ""
Write-Host "These applications are still ownerless" -ForegroundColor Red
Write-Host "----------------------------------------" -ForegroundColor Red
$AppReport | Sort-Object {$_.Created -as [datetime]} -Descending | Where-Object { $null -eq $_.Owners} | Format-Table AppName, Created, AppId -AutoSize 

$AppReport | Out-GridView -Title "Entra ID App Registration Owners Report" 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.