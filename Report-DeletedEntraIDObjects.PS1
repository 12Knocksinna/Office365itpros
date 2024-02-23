# Report-DeletedEntraIDObjects
# Example of how to report the soft-deleted objects in the Entra ID recycle bin
# V1.0 23-Feb-2024
# Ref https://techcommunity.microsoft.com/t5/microsoft-entra-blog/keeping-track-of-object-deletions-in-microsoft-entra-id/ba-p/4053415
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-DeletedEntraIDObjects.PS1

# Tested with SDK V2.15
Connect-MgGraph -NoWelcome -Scopes Directory.Read.All

$CSVOutputFile = "C:\temp\softDeletedObjects.csv"
$Now = Get-Date
[array]$DeletedGroups = Get-MgDirectoryDeletedItemAsGroup -All
[array]$DeletedSPs = Get-MgDirectoryDeletedItemAsServicePrincipal -All
[array]$DeletedAUs = Get-MgDirectoryDeletedItemAsAdministrativeUnit -All
[array]$DeletedApps = Get-MgDirectoryDeletedItemAsApplication -All
[array]$DeletedUsers = Get-MgDirectoryDeletedItemAsUser -All `
     -Property Id, DisplayName, DeletedDateTime, UserPrincipalName

$DeletedObjects =  [System.Collections.Generic.List[Object]]::new()

# Process deleted groups
ForEach ($Item in $DeletedGroups) {
    [string]$DeletedDateTime = $null
    $TimeTillRemoval = $null
    $PermanentRemovalDue = $null
  
    If ($Item.DeletedDateTime) {
        [datetime]$DeletedDateTime = Get-Date($Item.deletedDateTime)
        $PermanentRemovalDue = Get-Date($DeletedDateTime).AddDays(30)
        $TimeTillRemoval = $PermanentRemovalDue - $Now
    }
    If ($Item.CreatedDateTime) {
        [datetime]$CreatedDateTime = Get-Date ($Item.createdDatetime)
    }
    $ReportLine = [PSCustomObject]@{ 
         Group                      = $Item.displayName
         Id                         = $Item.id
         Created                    = Get-Date($CreatedDateTime) -format 'dd-MMM-yyyy HH:mm'
         Deleted                    = Get-Date($DeletedDateTime) -format 'dd-MMM-yyyy HH:mm'
         'Permanent Deletion due'   = Get-Date($PermanentRemovalDue) -format 'dd-MMM-yyyy HH:mm'
         DaysRemaining              = $TimeTillRemoval.Days 
         Type                       = "Group"
    } 
    $DeletedObjects.Add($ReportLine) 
}
# Deleted users
ForEach ($Item in $DeletedUsers) {
    [string]$DeletedDateTime = $null
    $TimeTillRemoval = $null
    $PermanentRemovalDue = $null
  
    If ($Item.DeletedDateTime) {
        [datetime]$DeletedDateTime = Get-Date($Item.deletedDateTime)
        $PermanentRemovalDue = Get-Date($DeletedDateTime).AddDays(30)
        $TimeTillRemoval = $PermanentRemovalDue - $Now
    }
    If ($Item.CreatedDateTime) {
        [datetime]$CreatedDateTime = Get-Date ($Item.createdDatetime)
    }
    $ReportLine = [PSCustomObject]@{ 
         Group                      = $Item.displayName
         Id                         = $Item.id
         Created                    = Get-Date($CreatedDateTime) -format 'dd-MMM-yyyy HH:mm'
         Deleted                    = Get-Date($DeletedDateTime) -format 'dd-MMM-yyyy HH:mm'
         'Permanent Deletion due'   = Get-Date($PermanentRemovalDue) -format 'dd-MMM-yyyy HH:mm'
         DaysRemaining              = $TimeTillRemoval.Days 
         Type                       = "User"
    } 
    $DeletedObjects.Add($ReportLine) 
}
# Deleted service principals
ForEach ($Item in $DeletedSPs) {
    [string]$DeletedDateTime = $null
    $TimeTillRemoval = $null
    $PermanentRemovalDue = $null
  
    If ($Item.DeletedDateTime) {
        [datetime]$DeletedDateTime = Get-Date($Item.deletedDateTime)
        $PermanentRemovalDue = Get-Date($DeletedDateTime).AddDays(30)
        $TimeTillRemoval = $PermanentRemovalDue - $Now
    }
    If ($Item.CreatedDateTime) {
        [datetime]$CreatedDateTime = Get-Date ($Item.createdDatetime)
    }
    $ReportLine = [PSCustomObject]@{ 
         Group                      = $Item.displayName
         Id                         = $Item.id
         Created                    = Get-Date($CreatedDateTime) -format 'dd-MMM-yyyy HH:mm'
         Deleted                    = Get-Date($DeletedDateTime) -format 'dd-MMM-yyyy HH:mm'
         'Permanent Deletion due'   = Get-Date($PermanentRemovalDue) -format 'dd-MMM-yyyy HH:mm'
         DaysRemaining              = $TimeTillRemoval.Days 
         Type                       = "Service Principal"
    } 
    $DeletedObjects.Add($ReportLine) 
}
# Deleted administrative units
ForEach ($Item in $DeletedAUs) {
    [string]$DeletedDateTime = $null
    $TimeTillRemoval = $null
    $PermanentRemovalDue = $null
  
    If ($Item.DeletedDateTime) {
        [datetime]$DeletedDateTime = Get-Date($Item.deletedDateTime)
        $PermanentRemovalDue = Get-Date($DeletedDateTime).AddDays(30)
        $TimeTillRemoval = $PermanentRemovalDue - $Now
    }
    If ($Item.CreatedDateTime) {
        [datetime]$CreatedDateTime = Get-Date ($Item.createdDatetime)
    }
    $ReportLine = [PSCustomObject]@{ 
         Group                      = $Item.displayName
         Id                         = $Item.id
         Created                    = Get-Date($CreatedDateTime) -format 'dd-MMM-yyyy HH:mm'
         Deleted                    = Get-Date($DeletedDateTime) -format 'dd-MMM-yyyy HH:mm'
         'Permanent Deletion due'   = Get-Date($PermanentRemovalDue) -format 'dd-MMM-yyyy HH:mm'
         DaysRemaining              = $TimeTillRemoval.Days 
         Type                       = "Administrative Unit"
    } 
    $DeletedObjects.Add($ReportLine) 
}
# Deleted applications
ForEach ($Item in $DeletedApps) {
    [string]$DeletedDateTime = $null
    $TimeTillRemoval = $null
    $PermanentRemovalDue = $null
  
    If ($Item.DeletedDateTime) {
        [datetime]$DeletedDateTime = Get-Date($Item.deletedDateTime)
        $PermanentRemovalDue = Get-Date($DeletedDateTime).AddDays(30)
        $TimeTillRemoval = $PermanentRemovalDue - $Now
    }
    If ($Item.CreatedDateTime) {
        [datetime]$CreatedDateTime = Get-Date ($Item.createdDatetime)
    }
    $ReportLine = [PSCustomObject]@{ 
         Group                      = $Item.displayName
         Id                         = $Item.id
         Created                    = Get-Date($CreatedDateTime) -format 'dd-MMM-yyyy HH:mm'
         Deleted                    = Get-Date($DeletedDateTime) -format 'dd-MMM-yyyy HH:mm'
         'Permanent Deletion due'   = Get-Date($PermanentRemovalDue) -format 'dd-MMM-yyyy HH:mm'
         DaysRemaining              = $TimeTillRemoval.Days 
         Type                       = "Application"
    } 
    $DeletedObjects.Add($ReportLine) 
}

If ($DeletedObjects.count -eq 0) { 
    Write-Host "No deleted items can be found - exiting"; break
} Else {
    Write-Host ""
    Write-Host ("Count of deleted applications         {0}" -f $DeletedApps.count)
    Write-Host ("Count of deleted administrative unuts {0}" -f $DeletedAUs.count)
    Write-Host ("Count of deleted groups               {0}" -f $DeletedGroups.count)
    Write-Host ("Count of deleted service principals   {0}" -f $DeletedSPs.count)
    Write-Host ("Count of deleted users                {0}" -f $DeletedUsers.count)
    Write-Host ""
    Write-Host ("Total count of deleted items          {0}" -f $DeletedObjects.count)
}

$DeletedObjects | Sort-Object {$_.PermanentDeleteOn -as [datetime]} | Out-GridView
$DeletedObjects | Export-Csv -Path $CSVOutputFile -Encoding UTF8 -NoTypeInformation  

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
