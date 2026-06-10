# Find-InactiveLicensedUserAccounts.PS1
# A script to find licensed user accounts that have not been active for a specified period of time. The script
# examines the baseline time for refresh tokens and the last successful sign-in date to determine if an account is inactive. The script also checks if the account is a room mailbox, as these can have licenses assigned but may not be used for sign-in.
# Application needs User.Read.All, Place.Read.All and LicenseAssignment.Read.All permissions in the Graph

# V1.0 7-June-2026
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-InactiveLicensedUserAccounts.PS1

Connect-MgGraph -Scopes "User.Read.All", "Place.Read.All", "LicenseAssignment.Read.All" -NoWelcome

# Find licensed users...
Try {
    [array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
        -Property Id, displayName, userPrincipalName, assignedLicenses, SignInSessionsValidFromDateTime, SignInactivity, licenseAssignmentStates, AccountEnabled `
        -ConsistencyLevel eventual -CountVariable Records -All -PageSize 500 | Sort-Object displayName
    Write-Host ("Fetched {0} licensed user accounts." -f $Records) -ForegroundColor Green
} Catch {
    Write-Host ("Error fetching user data: {0}" -f $_.Exception.Message)
    Break
}

# Get SKU data for the tenant
Try {
    [array]$SKUs = Get-MgSubscribedSKU | Select-Object SkuId, SkuPartNumber
    $SKUHash = @{}
    ForEach ($SKU in $SKUs) {
        $SKUHash.Add($SKU.SkuId, $SKU.SkuPartNumber)
    }
} Catch {
    Write-Host ("Error fetching SKU data: {0}" -f $_.Exception.Message)
    Break
}

# Find Places (room mailboxes - this won't find them all, but it's a start)... This data is used to mark licensed accounts as rooms if they are room
# mailboxes with licenses (required for archive, increased quota...).
Try {
    [array]$Places = Get-MgPlaceAsRoom -All -PageSize 500 | Sort-Object displayName
} catch {
    Write-Host ("Error fetching place data: {0}" -f $_.Exception.Message)
    Continue
}

# Date to check against is 90 days ago, but this can be modified as needed. 
# The SignInSessionsValidFromDateTime property is the date and time that the user's sign-in sessions are valid from. 
# If the user has not signed in since that date, then they are considered inactive.
$CutoffDate = (Get-Date).AddDays(-90).ToUniversalTime()
$Report = [System.Collections.Generic.List[Object]]::new()  

ForEach ($User in $Users) {
   
    If (($User.SignInSessionsValidFromDateTime) -and ($User.SignInSessionsValidFromDateTime -lt $CutoffDate) -and `
    (($null -eq $User.SignInActivity.LastSuccessfulSignInDateTime) -or `
    ([datetime]$User.SignInActivity.LastSuccessfulSignInDateTime -le $User.SignInSessionsValidFromDateTime))) {   
    
        If ($Places.DisplayName -contains $User.DisplayName) {
            $UserType = "Room"
        } else {
            $UserType = "User"
        }
        # Resolve product SKUs to friendly names for output
        [array]$ProductFriendlyNames = $null
        [array]$UserLicenses = $User.AssignedLicenses.SkuId
        ForEach ($License in $User.AssignedLicenses.SkuId) {
            If ($SKUHash.ContainsKey($License)) {
                $ProductFriendlyNames += $SKUHash[$License]
            } Else {
                $ProductFriendlyNames += "Unknown SKU (" + $License + ")"
            }
        }

        $ReportLine = [PSCustomObject][Ordered]@{
            UserPrincipalName   = $User.UserPrincipalName
            DisplayName         = $User.DisplayName
            SignInSessionsValidFromDateTime = if ($null -ne $User.SignInSessionsValidFromDateTime) { Get-Date $User.SignInSessionsValidFromDateTime -format "dd-MMM-yyyy HH:mm" } else { $null }
            SignInActivity      = If ($null -ne $User.SignInActivity.LastSuccessfulSignInDateTime) { Get-Date $User.SignInActivity.LastSuccessfulSignInDateTime -format "dd-MMM-yyyy HH:mm" } else { $null }
            DaysSinceLastSignIn = If ($null -ne $User.SignInActivity.LastSuccessfulSignInDateTime) { [int]((Get-Date) - (Get-Date $User.SignInActivity.LastSuccessfulSignInDateTime)).TotalDays } else { $null }    
            DaysSinceBaseline = [int]((Get-Date) - $User.SignInSessionsValidFromDateTime).TotalDays
            AccountEnabled      = $User.AccountEnabled
            UserType            = $UserType
            Licenses            = $ProductFriendlyNames -Join ", "
            SKUs                = $UserLicenses
        }
        $Report.Add($ReportLine)
    }
}

$Report | Out-GridView -Title "Inactive licensed user accounts" 

# Set up to use the ImportExcel module to generate Excel worksheets if it is available. If not, the report will be generated in CSV format instead.
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
}

# Generate the output file in either Excel worksheet or CSV format, depending on if the ImportExcel module is available. If interactive, create
# the file in the user's Downloads folder, otherwise create it in the current directory.
If ($ExcelGenerated) {
    $OutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Inactive Users.xlsx"
    Remove-Item -Path $OutputFile -ErrorAction SilentlyContinue
    $Report | Export-Excel -Path $OutputFile -WorksheetName "Inactive Users" -Title ("Inactive Users {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "InactiveUsers" -AutoSize -AutoFilter -FreezeTopRow
} Else {
    $OutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Inactive Users.csv"
    Remove-Item -Path $OutputFile -ErrorAction SilentlyContinue
    $Report | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding Utf8
}

Write-Host ("Report generated: {0}" -f $OutputFile) -ForegroundColor Green

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# The Automating Microsoft 365 with PowerShell eBook is available from https://o365itpros.gumroad.com/l/M365PS in EPUB and PDF format 
# or in paperback at https://amzn.to/3L4IBmU

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 