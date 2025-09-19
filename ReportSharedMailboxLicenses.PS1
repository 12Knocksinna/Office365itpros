# ReportSharedMailboxLicenses.PS1
# a script to analyze shared mailboxes and report if they need licenses
# Github link: https://github.com/12Knocksinna/Office365itpros/blob/master/ReportSharedMailboxLicenses.PS1

# V1.1 Updated 19-Sep-2025

# Check that we have the necessary Exchange Online module loaded
$ModulesLoaded = Get-Module | Select-Object -ExpandProperty Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {
  Write-Host "Connecting to Exchange Online..."
  Connect-ExchangeOnline -ShowBanner:$False -ErrorAction Stop
}

# Define some variables
$ExoPlan1 = "9aaf7827-d63c-4b61-89c3-182f06f82e5c"
$ExoArchiveAddOn = "176a09a6-7ec5-4039-ac02-b2791c6ba793"
$ExoPlan2 = "efb87545-963c-4e0d-99df-69c6916d9eb0" # See https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
$MailboxLimit = 50GB

Write-Host "Finding shared mailboxes..."
[array]$Mbx = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -Properties ProhibitSendReceiveQuota, ArchiveGuid, ArchiveName, ArchiveStatus, AutoExpandingArchiveEnabled, PrimarySmtpAddress, LitigationHoldEnabled -ResultSize Unlimited
If (!($Mbx)) {
  Write-Host "Can't find any shared mailboxes - exiting!" ; break
} Else {
  Write-Host ("Found {0} shared mailboxes" -f $Mbx.count)
}

[int]$i = 0
$SharedMbxReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
  $i++
  Write-Host ("Processing mailbox {0} ({1} of {2})" -f $M.DisplayName, $i, $Mbx.count)
  $NeedsLicense = $false; $ExoArchiveLicense = $false; $ExoPlan2License = $false; $ArchiveStats = $null
  $MailboxOverSize = $false; $ExoPlan1License = $false; $ArchiveMbxSize = $null

  $MbxStats = Get-ExoMailboxStatistics -Identity $M.ExternalDirectoryObjectId
  $MbxSize = [math]::Round(($MbxStats.TotalItemSize.Value.toBytes() / 1GB),5)
  If ($M.ArchiveStatus -ne "None") { #Mailbox has an archive
      $ArchiveStats = Get-ExoMailboxStatistics -Archive -Identity $M.ExternalDirectoryObjectId 
      If ($ArchiveStats) {       
        $ArchiveMbxSize = [math]::Round(($ArchiveStats.TotalItemSize.Value.toBytes() / 1GB),5)
      }
  }

  # Check licenses assigned to the mailbox's Entra ID account
  $Licenses = Get-MgUserLicenseDetail -UserId $M.ExternalDirectoryObjectId | Select-Object -ExpandProperty ServicePlans | Where-Object {$_.ProvisioningStatus -eq "Success"} | Sort-Object ServicePlanId -Unique
  If ($Licenses) { # The mailbox has some licenses
    If ($ExoArchiveAddOn -in $Licenses.ServicePlanId) { 
      $ExoArchiveLicense = $true 
    }
    If ($ExoPlan2 -in $Licenses.ServicePlanId) { 
        $ExoPlan2License = $true 
    }
    If ($ExoPlan1 -in $Licenses.ServicePlanId) { 
        $ExoPlan1License = $true
    }
  }

  # If the mailbox has an active archive, it needs an Exchange Online Plan 2 license unless it has an Exchange Online Archiving add-on license
  If ($M.ArchiveStatus -eq "Active") {
    If ($ExoPlan2License -eq $false) { 
      $NeedsLicense = $true
    }
    If ($ExoPlan1License -eq $true -and $ExoArchiveLicense -eq $true) { 1
      $NeedsLicense = $false 
    }
  }
  # Mailbox is on litigation hold and it doesn't have an Exchange Online Plan 2 license
  If ($M.LitigationHoldEnabled -eq $true -and $ExoPlan2License -eq $false)  { 
    $NeedsLicense = $true 
  }
  # Mailbox is over the 50GB usual quota for shared mailboxes and doesn't have an Exchange Online Plan 2 license
  If ($MbxStats.TotalItemSize.value -gt $MailboxLimit -and $ExoPlan2License -eq $false) { 
    $MailboxOverSize = $true
    $NeedsLicense = $true
  } Else {
    $MailboxOverSize = $false
  }

  $ReportLine = [PSCustomObject][Ordered]@{  # Write out details of the private channel and its members
      Mailbox               = $M.DisplayName
      UPN                   = $M.PrimarySmtpAddress
      NeedsLicense          = $NeedsLicense
      MailboxSizeGB         = $MbxSize
      MailboxItems          = $MbxStats.ItemCount
      MailboxOverSize       = $MailboxOverSize
      Archive               = $M.ArchiveStatus
      ArchiveSizeGB         = $ArchiveMbxSize
      ArchiveItems          = $ArchiveStats.ItemCount
      LitigationHoldEnabled = $M.LitigationHoldEnabled
      ExoPlan1              = $ExoPlan1License
      ExoPlan2              = $ExoPlan2License
      ExoArchiveAddOn       = $ExoArchiveLicense  
  }       
  $SharedMbxReport.Add($ReportLine) 

} #End ForEach $Mbx

[array]$NeedsLicenseList = $SharedMbxReport | Where-Object {$_.NeedsLicense -eq $True}
Write-Host " "
Write-Host ("{0} shared mailboxes scanned" -f $Mbx.count)
If ($NeedsLicenseList) {
  Write-Host ("{0} mailbox(es) need an Exchange Online Plan 2 license" -f $NeedsLicenseList.count)
  Write-Host ("Licenses must be assigned to the following mailboxes: {0}" -f $NeedsLicenseList.Mailbox -join ", " )
  $NeedsLicenseList | Format-Table Mailbox, MailboxOverSize, Archive, LitigationHoldEnabled
} Else {
  Write-Host "Congratulations! All of your shared mailboxes are properly licensed."
}

$SharedMbxReport | Out-GridView

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
