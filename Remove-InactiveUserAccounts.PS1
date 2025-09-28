# Remove-InactiveUserAccounts.PS1
# An Azure Automation runbbook to remove inactive user accounts from Entra ID
# 
# V1.0 27-Sep0-2025
# Github Link:


# Define theidentifier for the group holding details of accounts we want to review for inactivity. This could be a dynamic group.
$TargetGroupId = '1cfc9ab3-d230-45f6-a81f-8e32de9ad95b' 

# Thresholds for inactivity (number of days)    
$DaysSinceLastSignInThreshold = 90

# If no target group, process all licensed users
Try {
    [array]$TargetAccounts = Get-MgGroupMember -GroupId $TargetGroupId -All -PageSize 500 | Select-Object -ExpandProperty Id
} Catch {
    Write-Output "Error retrieving group containing target accounts: $_"
    Write-Output "Checking for all user accounts instead"
    Try {
        [array]$TargetAccounts = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member' and accountEnabled ne false" -ConsistencyLevel eventual `
        -CountVariable UsersFound -Property id, displayName, userprincipalname, usertype, signInActivity, SignInSessionsValidFromDateTime, LastPasswordChangeDateTime, passwordPolicies `
        -All -PageSize 500 -Sort displayName -ErrorAction Stop | Select-Object -ExpandProperty Id
    } Catch {
        Write-Output "Error retrieving all user accounts: $_"
        Break    
    }
}

Write-Output "Found $($TargetAccounts.Count) target accounts to check for inactivity"

# Before checking any account for inactivity, find the disabled accounts from the last run of the script. We don't want to reprocess these accounts and will delete them 
# at the end of the script
Try {
    [array]$DisabledAccounts = Get-MgUser -Filter "accountEnabled eq false and onPremisesExtensionAttributes/extensionAttribute10 eq 'Inactive' and assignedLicenses/`$count ne 0 and userType eq 'Member'" `
        -All -PageSize 500 -ConsistencyLevel eventual -CountVariable DisabledAccountsFound 
} Catch {
    Write-Output "Error retrieving disabled accounts: $_"
    $DisabledAccounts = @()
}

$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Id in $TargetAccounts) {

    $DaysSinceLastSignIn = $null; $DaysSinceLastSuccessfulSignIn = $null

    Try {
        $User = Get-MgUser -UserId $Id -Property id, displayName, userprincipalname, department, signInActivity, SignInSessionsValidFromDateTime, accountEnabled -ErrorAction Stop
    } Catch {
        Write-Output ("Error retrieving user with identifier {0}: $_" -f $Id)
        Continue
    }
      
    If (!([string]::IsNullOrWhiteSpace($User.signInActivity.lastSuccessfulSignInDateTime))) {
        [datetime]$LastSuccessfulSignIn = $User.signInActivity.lastSuccessfulSignInDateTime
        $DaysSinceLastSuccessfulSignIn = (New-TimeSpan $LastSuccessfulSignIn).Days 
    }
    If (!([string]::IsNullOrWhiteSpace($User.signInActivity.lastSignInDateTime))) {
        [datetime]$LastSignIn = $User.signInActivity.lastSignInDateTime
        $DaysSinceLastSignIn = (New-TimeSpan $LastSignIn).Days
    }    

    If ($DaysSinceLastSuccessfulSignIn -ge $DaysSinceLastSignInThreshold -or $DaysSinceLastSignIn -ge $DaysSinceLastSignInThreshold) {
        Write-Output ("User {0} ({1}) is inactive. Last sign-in {2} days ago, last successful sign-in {3} days ago" -f $User.displayName, $User.userPrincipalName, `
            $DaysSinceLastSignIn, $DaysSinceLastSuccessfulSignIn)

        # Collect details for reporting
        $ReportItem = [PSCustomObject]@{
            DisplayName                 = $User.DisplayName
            UserPrincipalName           = $User.UserPrincipalName
            Department                  = $User.Department
            AccountEnabled              = $User.AccountEnabled
            LastSignInDateTime          = $LastSignIn
            DaysSinceLastSignIn         = $DaysSinceLastSignIn
            LastSuccessfulSignInDateTime = $LastSuccessfulSignIn
            DaysSinceLastSuccessfulSignIn = $DaysSinceLastSuccessfulSignIn
        }
        $Report.Add($ReportItem)

        # Disable the user account
        Try {
            Update-MgUser -UserId $User.Id -accountEnabled:$false -OnPremisesExtensionAttributes @{'extensionAttribute10' = 'Inactive'} -ErrorAction Stop
            Write-Output ("Disabled user account {0} ({1})" -f $User.DisplayName, $User.UserPrincipalName)
        } Catch {
            Write-Output ("Error disabling user account {0} ({1}): $_" -f $User.DisplayName, $User.UserPrincipalName)
        }
    } Else {
        Write-Output ("User {0} ({1}) is active. Last sign-in {2} days ago, last successful sign-in {3} days ago" -f $User.displayName, $User.userPrincipalName, `
            $DaysSinceLastSignIn, $DaysSinceLastSuccessfulSignIn)
    }   
}

$Report | Sort-Object DaysSinceLastSuccessfulSignIn -Descending | Select-Object DisplayName, UserPrincipalName, LastSignInDateTime | Format-Table -AutoSize

# Create an attachment for the email
$Report | Export-CSV InactiveUserAccounts.CSV -NoTypeInformation -Encoding UTF8
$Attachment = (Get-Location).Path + "\InactiveUserAccounts.CSV"
$EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($Attachment))
$MsgAttachments = @(
    @{
	"@odata.type" = "#microsoft.graph.fileAttachment"
	Name = ($Attachment -split '\\')[-1]
	ContentType = "application/vnd.ms-excel"
	ContentBytes = $EncodedAttachmentFile
	}
)

# Delete the accounts we disabled during the last run of the script
If ($DisabledAccounts) {
    ForEach ($Account in $DisabledAccounts) {
        Try {
            Set-Mailbox -Identity $Account.Id -LitigationHoldEnabled $true -LitigationHoldDate (Get-Date) -LitigationHoldOwner "Administrator Workflow"
            Remove-MgUser -UserId $Account.Id -ErrorAction Stop
            Write-Output ("Deleted user account {0} ({1})" -f $Account.DisplayName, $Account.UserPrincipalName)
        } Catch {
            Write-Output ("Error deleting user account {0} ({1}): $_" -f $Account.DisplayName, $Account.UserPrincipalName)
        }
    }
} Else {
    Write-Output "No accounts found that were disabled by the last run of the script"
}

$HtmlMsg = $null
$HtmlMsg ="<h2>Inactive user account report - $(Get-Date -Format 'dd-MMM-yyyy')</h2>"
$HtmlMsg += "<p>The inactive user accounts workflow found $($Report.Count) inactive accounts and disabled them</p>"
$HtmlMsg += $Report | Select-Object DisplayName, UserPrincipalName, LastSignInDateTime | ConvertTo-Html -Fragment

$HtmlMsg += "<p>The disabled accounts found by this workflow will be deleted the next time the inactive users workflow runs if no action is taken to re-enable them.</p>"

If ($DisabledAccounts) {
    $HtmlMsg += "<h3>Accounts deleted</h3>"
    $HtmlMsg += "<p>The following accounts were disabled by the last run of the script and have now been deleted:</p>"
    $HtmlMsg += $DisabledAccounts | Select-Object DisplayName, UserPrincipalName | ConvertTo-Html -Fragment
} Else {
    $HtmlMsg += "<p>No accounts were deleted because no accounts were found that had been disabled by the last run of the workflow.</p>"
}

$HtmlMsg += "<p>This is an automated message generated by the Remove-InactiveUserAccounts.PS1 script. More information about the script can be found in the Office 365 for IT Pros GitHub repository.</p>"

# Create and send a message to report what's hapened
$MsgFrom = 'Customer.Services@office365itpros.com'
$MsgSubject = "Inactive user account workflow report"
$ToRecipient = @{}
# Update the target email address here
$ToRecipient.Add("emailAddress",@{'address'="Tony.Redmond@office365itpros.com"})
[array]$MsgTo = $ToRecipient

# Construct the message body 	
$MsgBody = @{}
$MsgBody.Add('Content', "$($HtmlMsg)")
$MsgBody.Add('ContentType','html')

# Build the parameters to submit the message
$Message = @{}
$Message.Add('subject', $MsgSubject)
$Message.Add('toRecipients', $MsgTo)
$Message.Add('body', $MsgBody)
$Message.Add("attachments", $MsgAttachments)
$EmailParameters = @{}
$EmailParameters.Add('message', $Message)
$EmailParameters.Add('saveToSentItems', $true)
$EmailParameters.Add('isDeliveryReceiptRequested', $true)

# Send the message
Try {
  Write-Host "Sending email to $($MsgTo.emailAddress.address)" -ForegroundColor Yellow
  Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters
} Catch {
  Write-Host "Failed to send email to $($MsgTo.emailAddress.address)" -ForegroundColor Red
}

Write-Output "Processing complete for inactive user workflow"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.