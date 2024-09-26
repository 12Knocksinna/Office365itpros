# Report-SharingOneDriveFiles.PS1
# A script to show how to report the permissions for shared files found in OneDrive for Business accounts
# V1.0 24-Sep-2024

# Requires an Entra ID app with the following application permissions:
# User.Read.All         Read User Profiles 
# Sites.Read.All        Read SharePoint sites, including OneDrive for Business sites
# Group.Read.All        Read all groups (to resolve group information in sharing links)
# GroupMember.Read.All  Read all group members (to resolve group information in sharing links)

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-SharingOneDriveFiles.PS1

function Get-DriveItems {
    [CmdletBinding()]
    param (
        [Parameter()]
        $Drive,
        [Parameter()]
        $FolderId
    )
    # Get data for a folder and its children
    [array]$Data = Get-MgDriveItemChild -DriveId $Drive -DriveItemId $FolderId -All
    # Split the data into files and folders
  
    [array]$Folders = $Data | Where-Object {$_.folder.childcount -gt 0} | Sort-Object Name
    $Global:TotalFolders = $TotalFolders + $Folders.Count
    [array]$Files = $Data | Where-Object {$null -ne $_.file.mimetype} 
    $TotalFiles = $TotalFiles + $Files.Count
    [array]$SharedItems = $Files | Where-Object {$null -ne $_.Shared.Scope}

    If ($null -eq $Folder.Name) {
        $FolderName = "the OneDrive account root"
    } Else {
        $FolderName = ("folder: {0}" -f $Folder.Name)
    }
    # Process the files
    If (!($SharedItems)) {
        Write-Host ("No shared files found in {0}" -f $FolderName)
        Return
    }
    ForEach ($File in $SharedItems) {  
        If ($File.LastModifiedDateTime) {
            $LastModifiedDateTime = Get-Date $File.LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
        } Else {
            $LastModifiedDateTime = $null
        }
        If ($File.CreatedDateTime) {
            $FileCreatedDateTime = Get-Date $File.CreatedDateTime -format 'dd-MMM-yyyy HH:mm'
        }
        [array]$Permissions = Get-MgDriveItemPermission -DriveId $Drive -DriveItemId $File.Id `
            -Property Roles, GrantedTo, HasPassword, ExpirationDateTime, Invitation, InheritedFrom, Id, Link, GrantedToV2 | Sort-Object Roles
        [array]$GrantedPermissions = $Permissions | Where-Object {$_.Roles[0] -ne 'owner'}
        If ($GrantedPermissions) {
            ForEach ($Permission in $GrantedPermissions) {
                $Role = ($Permission.Roles -join ",").trim()
                Switch ($Role) {
                    "read" {
                        $Scope = "View only"
                    }
                    "write" {
                        $Scope = "Edit"
                    }
                    default {
                        $Scope = "View only"
                    }
                }                  
                $GrantedTo = $null; $GrantedToId = $null
                # Use GrantToV2 if it's available
                If ($Permission.GrantedToV2.User.DisplayName) {
                    $GrantedTo = $Permission.GrantedToV2.User.DisplayName
                    $GrantedToId = $Permission.Grantedtov2.SiteUser.LoginName.Split("|")[2]
                } Else {
                    $GrantedTo = $Permission.GrantedTo.User.DisplayName
                    $GrantedToId = $Permission.GrantedTo.User.Id
                }
                # Handle the different types of sharing link
                $LinkExpired = $null
                If ($Permission.link.scope) {
                    If ($Permission.ExpirationDateTime) {
                        If ($Permission.ExpirationDateTime -lt [datetime]::Now) {
                            $ExpirationDate = ("Link expired on: {0}" -f $Permission.ExpirationDateTime)
                            $LinkExpired = $true
                        } Else { 
                            $ExpirationDate = ("Link expires on: {0}" -f $Permission.ExpirationDateTime) 
                            $LinkExpired = $false
                        }
                    } Else {
                        $ExpirationDate = $null
                    }
                    $GrantedToId = "N/A"
                    Switch ($Permission.link.scope) {
                        "anonymous" {
                            $GrantedTo = "Anyone link"
                            $GrantedToId = "Anyone with the link"
                        }
                        "organization" {
                            $GrantedTo = "Organization link"
                            $GrantedToId = "Anyone in the tenant"
                        } 
                        "existingAccess" {
                            $GrantedTo = "Existing access link"
                        }
                        "users" {
                            $GrantedTo = "Specific set of users"
                            $GrantedToId = "Users with the link"
                        }
                    } 
                }
                # Resolve values returned for accounts that no longer exist and guest accounts
                If (Test-Numeric ($GrantedToId)) {
                    $GrantedToId = "User account removed from tenant"
                } ElseIf ($GrantedToId -like "*#EXT#*") {
                    $GrantedToId = Get-MgUser -UserId $GrantedToId | Select-Object -ExpandProperty Mail
                } ElseIf ([guid]::TryParse($GrantedToId, $([ref][guid]::Empty))) {
                    # The identifier is a GUID rather than a UPN, so the sharing is with a group. Check the hash table to see if we've seen the group before. 
                    # If so, use the data from the table. If not, fetch the group members and add the data to the table.
                    [array]$GroupMembers = $null
                    $GroupId = $GrantedToId
                    If ($GroupsHash[$GroupId]) {
                        $GrantedTo = ("Group: {0}:" -f ($GrantedTo))
                        $GrantedToId = $GroupsHash[$GroupId]
                    } Else {
                        [array]$GroupMembers = Get-MgGroupMember -GroupId $GroupId
                        $GrantedTo = ("Group: {0}:" -f ($GrantedTo))
                        $GroupMembersNames = ("Members: {0}" -f ($GroupMembers.additionalProperties.displayName -join ","))
                        $GroupsHash.Add($GroupId, $GroupMembersNames)
                        $GrantedToId = $GroupMembersNames
                    }
                }
                If ($Permission.Link.PreventDownload) {
                    $PreventDownload = $true
                } Else {
                    $PreventDownload = $false
                }
                If ($null -ne $GrantedTo) {
                    $ReportLine = [PSCustomObject]@{
                        Account             = $Site.DisplayName
                        ItemName            = $File.Name
                        Folder              = $File.parentreference.name
                        'Access granted to' = $GrantedTo
                        'Effective scope'   = $GrantedToId
                        Permission          = $Scope.trim()
                        Size                = (FormatFileSize $File.Size)
                        Created             = $FileCreatedDateTime
                        Author              = $File.CreatedBy.User.DisplayName
                        HasPassword         = $Permission.HasPassword
                        'Expiration date'   = $ExpirationDate
                        'Link Expired'      = $LinkExpired
                        'Prevents Download' = $PreventDownload
                        LastModified        = $LastModifiedDateTime
                        'Last modified by'  = $File.LastModifiedBy.User.DisplayName
                        ItemUrl             = $File.WebUrl
                        SiteUrl             = $Site.WebUrl
                    }
                    $Report.Add($ReportLine)
                    Start-Sleep -Milliseconds 50
                }
            }
        }
       
    }

    # Process the folders
    ForEach ($Folder in $Folders) {
        Write-Host ("Processing folder {0} ({1} files/size {2})" -f $Folder.Name, $Folder.folder.childcount, (FormatFileSize $Folder.Size))
        Get-DriveItems -Drive $Drive -FolderId $Folder.Id
    }

}
function FormatFileSize {
    # Format File Size nicely
    param (
            [parameter(Mandatory = $true)]
            $InFileSize
        ) 
    
    If ($InFileSize -lt 1KB) { # Format the size of a document
        $FileSize = $InFileSize.ToString() + " B" 
    } 
    ElseIf ($InFileSize -lt 1MB) {
        $FileSize = $InFileSize / 1KB
        $FileSize = ("{0:n2}" -f $FileSize) + " KB"
    } 
    Elseif ($InFileSize -lt 1GB) {
        $FileSize = $InFileSize / 1MB
        $FileSize = ("{0:n2}" -f $FileSize) + " MB" 
    }
    Elseif ($InFileSize -ge 1GB) {
        $FileSize = $InFileSize / 1GB
        $FileSize = ("{0:n2}" -f $FileSize) + " GB" 
    }
    Return $FileSize
} 

function Test-Numeric ($Value) {
    return $Value -match "^[\d\.]+$"
}

# Define the values required to connect. These are for the Entra ID app used to authentciate with the Graph
# You must replace these values with your own.
$AppId = "72bd89d6-060a-43c9-8063-c281d8f8b685"
$Tenantid = 'a662313f-14fc-43a2-9a7a-d2e27f4f3478'
# Make sure that the certificate is in date and won't expire soon!
$Thumbprint = "2C9529B1FFD08BCD483A5D98807E47A472C5318"
Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome

# Find SharePoint sites (use Get-MgAllSite to handle multi-geo tenants)
Write-Host "Looking for OneDrive for Business sites..."
[array]$Sites = Get-MgSite -All -PageSize 500 -Property DisplayName, WebUrl, IsPersonalSite, CreatedByUser, CreatedDateTime, Description, Name, id
# Reduce the set to OneDrive sites
[array]$OneDriveSites = $Sites | Where-Object {$_.IsPersonalSite -eq $true}
If ($OneDriveSites.Count -eq 0) {
    Write-Host "No OneDrive for Business sites found"
    Break
} Else {
    Write-Host ("Found {0} OneDrive for Business sites" -f $OneDriveSites.Count)
    Write-Host "Some of the sites might be for users who have been removed from the tenant. The script will skip these sites during processing."
}

Write-Host "Finding user account information to validate OneDrive sites..."
# Fetch licensed users so that we can check OneDrive sites to find ones for current users
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
    -ConsistencyLevel eventual -CountVariable Records -All -PageSize 999 | Sort-Object displayName
$UserHash = @{}
ForEach ($User in $Users) {
    $UserHash.Add($User.DisplayName, $User.UserPrincipalName)
}

# Set up the report variables
$Global:Report = [System.Collections.Generic.List[Object]]::new()
$Global:GroupsHash = @{}
[int]$i = 0

# Go and process
ForEach ($Site in $OneDriveSites) {
    If ($UserHash[$Site.name]) {
        $Global:TotalFiles = 0
        $i++
        Write-Host ("Processing OneDrive site for {0} {1}/{2}" -f $Site.DisplayName, $i, $OneDriveSites.Count) -ForegroundColor Yellow
        Try {
            [array]$Drives = Get-MgSiteDrive -SiteId $Site.id
            $Drive = $Drives[0]
            Get-DriveItems -Drive $Drive.Id -FolderId "root"
        }
        Catch {
            Write-Host ("Error processing OneDrive site {0}. The account might be locked or the user might never have used OneDrive." -f $Site.DisplayName)
            Continue
        }
        # Brief pause before we process the next account
        Start-Sleep -Seconds 2
    }
}

# Build a list of users who share files
[array]$UsersWhoShare = $Report | Sort-Object Account -Unique | Select-Object Account
ForEach ($U in $UsersWhoShare) {
    $UPN = $UserHash[$U.Account]
    If ($UPN) {
        Add-Member -InputObject $U -MemberType NoteProperty -Name "UPN" -Value $UPN
    }
    [int]$NumberFiles = 0
    [int]$AnyOneLinks = 0
    $NumberFiles = $Report | Where-Object {$_.Account -eq $U.Account} | Sort-Object ItemName -Unique | Measure-Object | Select-Object -ExpandProperty Count
    $AnyOneLinks = $Report | Where-Object {$_.Account -eq $U.Account -and $_.'Access granted to' -eq "Anyone link"} | Measure-Object | Select-Object -ExpandProperty Count
    $OrganizationLinks = $Report | Where-Object {$_.Account -eq $U.Account -and $_.'Access granted to' -eq "Organization link"} | Measure-Object | Select-Object -ExpandProperty Count
    Add-Member -InputObject $U -MemberType NoteProperty -Name "Shared Files" -Value $NumberFiles
    Add-Member -InputObject $U -MemberType NoteProperty -Name "Anyone Links" -Value $AnyoneLinks
    Add-Member -InputObject $U -MemberType NoteProperty -Name "Organization Links" -Value $OrganizationLinks
}

Write-Host "Summary of sharing activity in OneDrive for Business sites"
$UsersWhoShare | Format-table -AutoSize
Write-Host ""

If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $SummaryOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\OneDrive Files Sharing Summary Report.xlsx"
    If (Get-Item $SummaryOutputFile -ErrorAction SilentlyContinue) {
        Remove-Item $SummaryOutputFile -ErrorAction SilentlyContinue
    }
    $UsersWhoShare | Export-Excel -Path $SummaryOutputFile -WorksheetName "OneDrive Files Sharing Summary" `
        -Title "OneDrive Files Sharing Summary Report" -TitleBold -TableName "OneDriveSharingSummary"  
    $DetailedOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\OneDrive Files Sharing Report.xlsx"
    If (Get-Item $DetailedOutputFile -ErrorAction SilentlyContinue) {
        Remove-Item $DetailedOutputFile -ErrorAction SilentlyContinue
    }
    $Report | Export-Excel -Path $DetailedOutputFile -WorksheetName "OneDrive Files Sharing Details" `
        -Title "OneDrive for Business Files Sharing Details" -TitleBold -TableName "OneDriveSharingDetails" 
} Else {
    $SummaryCSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\OneDrive Files Sharing Summary Report.CSV"
    $UsersWhoShare | Export-Csv -Path $SummaryCSVOutputFile -NoTypeInformation -Encoding Utf8
    $DetailedCSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\OneDrive Files Sharing Detailed Report.CSV"
    $Report | Export-Csv -Path $DetailedCSVOutputFile -NoTypeInformation -Encoding Utf8
}

If ($ExcelGenerated) {
    Write-Host ("Excel worksheets generated in your Downloads folder: {0}, {1}" -f $SummaryOutputFile, $DetailedOutputFile)
} Else {
    Write-Host ("CSV files generated in your Downloads folder: {0}, {1}" -f $SummaryCSVOutputFile, $DetailedCSVOutputFile)
} 

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
