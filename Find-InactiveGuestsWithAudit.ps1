# Find-InactiveGuestsWithAudit.PS1
# A script to find inactive Entra ID guests and report what they've been doing

# V1.0 5-Oct-2025
# GitHub Link: 

Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All" -NoWelcome
Connect-ExchangeOnline -ShowBanner:$false

# Find information about sharing events so that we know  when someone has been invited to the tenant
$ShareDateStart = (Get-Date).AddDays(-365)
$EndDate = (Get-Date).AddDays(1)
# SharePoint Sharing Invitation
Write-Output "Searching for SharePoint sharing invitations..."
[array]$SharingRecords = Search-UnifiedAuditLog -StartDate $ShareDateStart -EndDate $EndDate -SessionCommand ReturnLargeSet -Formatted -ResultSize 5000 -Operations SharingInvitationCreated
$SharingRecords = $SharingRecords | Sort-Object Identity -Unique
$SharingData = [System.Collections.Generic.List[Object]]::new()
ForEach ($Record in $SharingRecords) {
    $AuditData = $Record.AuditData | ConvertFrom-Json
    If ($AuditData.TargetUserOrGroupType -eq "Guest") {
        $Guest = $AuditData.TargetUserOrGroupName
        $InvitationSource = $AuditData.UserId
        $Item = [PSCustomObject]@{
            Guest               = $Guest.toLower()
            InvitationSource    = $InvitationSource
            Timestamp           = $Record.CreationDate
            App                 = 'SharePoint Online'
        }
        $SharingData.Add($Item)
    }
}
# Guests added to Microsoft 365 Groups
Write-Output "Searching for Microsoft 365 Groups invitations..."
[array]$GroupData =  Search-UnifiedAuditLog -StartDate $ShareDateStart -EndDate $EndDate -SessionCommand ReturnLargeSet -Formatted `
    -ResultSize 5000 -Operations 'Add member to group.'
$GroupData = $GroupData | Sort-Object Identity -Unique
ForEach ($Record in $GroupData) {
    $AuditData = $Record.AuditData | ConvertFrom-Json
    If ($AuditData.ObjectId -Like "*#EXT#*" -and $AuditData.UserId -notlike "*ServicePrincipal*") {
        $Guest = $AuditData.ObjectId
        $InvitationSource = $AuditData.UserId
        $Item = [PSCustomObject]@{
            Guest               = $Guest.toLower()  
            InvitationSource    = $InvitationSource
            Timestamp           = $Record.CreationDate
            App                 = 'Groups'
        }
        $SharingData.Add($Item)
    }
}
# Users invited from the Entra ID portal
Write-Output "Searching for Entra ID invitations..."
[array]$InvitationData = Search-UnifiedAuditLog -StartDate $ShareDateStart -EndDate $EndDate -SessionCommand ReturnLargeSet -Formatted `
    -ResultSize 5000 -Operations 'Add user.'
$InvitationData = $InvitationData | Sort-Object Identity -Unique
ForEach ($Record in $InvitationData) {
    $AuditData = $Record.AuditData | ConvertFrom-Json
    If ($AuditData.ObjectId -Like "*#EXT#*" -and $AuditData.UserId -notlike "*ServicePrincipal*") {
        $Guest = $AuditData.ObjectId
        $InvitationSource = $AuditData.UserId
        $Item = [PSCustomObject]@{
            Guest               = $Guest.toLower()
            InvitationSource    = $InvitationSource
            Timestamp           = $Record.CreationDate
            App                 = 'Entra ID'
        }
        $SharingData.Add($Item)
    }
}

$SharingData  = $SharingData | Sort-Object {$_.TimeStamp -as [datetime]} -Descending

$StartDate = (Get-Date).AddDays(-30)

# Define the audit records used to figure out what guests have been doing
[array]$Operations = "FileAccessed","FileModified","FileUploaded","FileDeleted","FileDownloaded","MessageSent", "ReactedToMessage", "MessageRead", "MessageDeleted"

# Find all guests - a complex query is used to sort the retrieved results
[array]$Guests = Get-MgUser -Filter "usertype eq 'Guest'" -PageSize 500 -All  `
    -Property DisplayName,UserPrincipalName,SignInActivity,Mail,Sponsors,Id,CreatedDateTime -ExpandProperty Sponsors | Sort-Object displayName
If ($Guests.Count -eq 0) {
    Write-Host "No guest users found."
    break
} Else {
    Write-Host ("Found {0} guest users" -f $Guests.Count)
}

[int]$i = 0
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Guest in $Guests) {
    $i++
    $DaysSinceLastSignIn = $null; $DaysSinceLastSuccessfulSignIn = $null
    $GuestStatus = "Inactive"
    $TopActivities = $null
    Write-Host "Processing guest user $($Guest.DisplayName) <$($Guest.Mail)> ($($i)/$($Guests.Count))" -ForegroundColor Cyan
    [array]$Records = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -UserIds $Guest.UserPrincipalName `
        -SessionCommand ReturnLargeSet -Formatted -ResultSize 5000 -Operations $Operations
    If ($Records.Count -eq 0) {
        Write-Host "No audit records found for guest user" -ForegroundColor Red
        $LastActivity = $null
    } Else {
        # Remove duplicate records
        $Records = $Records | Sort-Object Identity -Unique
        Write-Host ("Found {0} audit records for guest user" -f $Records.Count) -ForegroundColor Yellow
        $LastActivity = $Records | Sort-Object {$_.CreationDate -as [datetime]} -Descending | Select-Object -First 1
        [array]$GuestActivity = $Records | Group-Object Operations -NoElement | Sort-Object Count -Desc | Select-Object -First 3
        $TopActivities = ($GuestActivity | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ", "
        $GuestStatus = "Active"
    }
     
    # Can we find when the guest was invited to the tenant?
    [array]$Invitation = $SharingData | Where-Object { $_.Guest -eq $Guest.UserPrincipalName.tolower() } | `
        Sort-Object {$_.TimeStamp -as [datetime]} -Descending | Select-Object -First 1
    If ($Invitation) {
        $InvitedTimeStamp = Get-Date $Invitation.Timestamp -Format 'dd-MMMM-yyyy HH:mm'
        $InvitedSource = $Invitation.InvitationSource
    } Else {
        $InvitedTimeStamp = $null
        $InvitedSource = $null
    }

    If (!([string]::IsNullOrWhiteSpace($Guest.signInActivity.lastSuccessfulSignInDateTime))) {
        [datetime]$LastSuccessfulSignIn = $Guest.signInActivity.lastSuccessfulSignInDateTime
        $DaysSinceLastSuccessfulSignIn = (New-TimeSpan $LastSuccessfulSignIn).Days 
    }
    If (!([string]::IsNullOrWhiteSpace($Guest.signInActivity.lastSignInDateTime))) {
        [datetime]$LastSignIn = $Guest.signInActivity.lastSignInDateTime
        $DaysSinceLastSignIn = (New-TimeSpan $LastSignIn).Days
    }
   
    $ReportItem = [PSCustomObject]@{
        DisplayName                     = $Guest.DisplayName
        UserPrincipalName               = $Guest.UserPrincipalName
        Email                           = $Guest.Mail
        Sponsors                        = ($Guest.Sponsors | ForEach-Object { Get-MgUser -UserId $_.Id | Select-Object -ExpandProperty DisplayName }) -join "; "
        'Creation Date'                 = Get-Date $Guest.CreatedDateTime -format 'dd-MMMM-yyyy HH:mm'
        'Days since creation'           = (New-TimeSpan $Guest.CreatedDateTime).Days
        'Date of Last Audit Activity'   = If ($LastActivity) { Get-Date $LastActivity.CreationDate -format 'dd-MMMM-yyyy HH:mm'} else { $null }
        'Last Audit Activity'           = If ($LastActivity) { $LastActivity.Operations } else { $null }
        'Number of Audit Activities'    = $Records.Count
        'Top 3 activities'              = $TopActivities
        'Last administrator action'     = $InvitedTimeStamp
        'Administrator'                 = $InvitedSource
        LastSignInDateTime              = $LastSignIn
        DaysSinceLastSignIn             = $DaysSinceLastSignIn
        LastSuccessfulSignInDateTime    = $LastSuccessfulSignIn
        DaysSinceLastSuccessfulSignIn   = $DaysSinceLastSuccessfulSignIn
        EmailDomain                     = ($Guest.Mail -split "@")[1]
        'Guest status'                  = $GuestStatus
    }
    $Report.Add($ReportItem)

}

[array]$InactiveGuests = $Report | Where-Object { $_.'Guest status' -eq "Inactive" } | Sort-Object DisplayName
Write-Host ""
Write-Host ("Found {0} inactive guests ({1})" -f $InactiveGuests.Count,($InactiveGuests.Count/$Report.Count).toString("P")) -ForegroundColor Green

Write-Host ""
Write-Host "Inactive guests come from these domains"
$InactiveGuests | Group-Object EmailDomain | Sort-Object Count -Descending | Format-Table Name,Count -AutoSize

# And generate an output file
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $true
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveGuests.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    } 
    $InactiveGuests | Export-Excel -Path $ExcelOutputFile -WorksheetName "Inactive Guests" -Title ("Inactive Guests Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "InactiveGuests" 
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveGuests.CSV"
    $InactiveGuests | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}

If ($ExcelGenerated) {
    Write-Host ("Excel worksheet output written to {0}" -f $ExcelOutputFile)
} Else {
    Write-Host ("CSV output file written to {0}" -f $CSVOutputFile)
}   


# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
