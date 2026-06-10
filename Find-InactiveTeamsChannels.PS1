# Find-InactiveTeamsChannels.PS1
# This script identifies inactive channels in Microsoft Teams based on the last activity date of a user message posted to the channel.

# V1.0 5-June-2026
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Find-InactiveTeamsChannels.PS1

# Necessary permissions Team.ReadBasic.All, Channel.ReadBasic.All, ChannelMessage.Read.All, ChannelMember.Read.All
# if you want the script to archive channels automatically, you'll need ChannelSettings.ReadWrite.All
# before you can run Invoke-MgArchiveTeamChannel -TeamId $teamId -ChannelId $channelId

[array]$RequiredScopes = "Team.ReadBasic.All", "Channel.ReadBasic.All", "ChannelMessage.Read.All", "ChannelMember.Read.All", "Mail.Send"
$Interactive = $false

# Determine if we're interactive or not
If ([Environment]::UserInteractive) { 
    # We're running interactively...
    Clear-Host
    Write-Host "Script running interactively... connecting to the Graph" -ForegroundColor Yellow
    Write-Host "Access is only possible to channels in Teams where the signed-in account is a member."
    Connect-MgGraph -NoWelcome
    $Interactive = $true
    # Email address to use when sending email from interactive session
    $MsgFrom = (Get-MgContext).Account    
} Else { 
    # We're not, so likely in Azure Automation
    Write-Output "Executing the runbook to send email about incomplete tasks..." 
    Connect-MgGraph -Identity -NoWelcome
    # Email address to use when sending email from Azure Automation - update for your tenant
    $MsgFrom = "no-reply@office365itpros.com"
}
# Add the recipient address to receive the report of inactive channels. Update this for your needs. 
$EmailRecipient = "Customer.Services@Office365itpros.com"

# Check that we have the right permissions - in Azure Automation, we assume that the automation account has the right permissions
If ($Interactive) {
    [int]$RequiredScopesCount = $RequiredScopes.Count
    [string[]]$CurrentScopes = (Get-MgContext).Scopes
    [string[]]$RequiredScopes = $RequiredScopes

    $CheckScopes =[object[]][Linq.Enumerable]::Intersect($RequiredScopes,$CurrentScopes)
    If ($CheckScopes.Count -ne $RequiredScopesCount ) { 
        Write-Host ("To run this script, you need to connect to Microsoft Graph with the following scopes: {0}" -f $RequiredScopes) -ForegroundColor Red
        Break
    }
}

# To sign in interactively with app-only access
# $AppId = '6694eecf-adf8-4d56-bf73-bae26cd02ebd'
# $TenantId = '72f988bf-86f1-4111-9a1a-0123456789ab'
# $CertificateThumbprint = '8CD4AD070C7447BA767EAB5DA659A02E6411BB80'
# Connect-MgGraph -ClientId "app-id" -TenantId "tenant-id" -CertificateThumbprint "thumbprint"

# Threshold for inactivity in days
$ThresholdDays = 180
# Channel is deemed inactive if there has been no user activity for more than this number of days. Or we find no user messages in the channel.

Write-Host "Finding teams to check for inactive channels..." -ForegroundColor Green
Try {
    [array]$Teams = Get-MgTeam -All -ErrorAction Stop 
    Write-Host "Found $($Teams.Count) teams to check for inactive channels." -ForegroundColor Green
    $Teams = $Teams | Sort-Object DisplayName
} catch {
    Write-Host "Error fetching teams: $_" -ForegroundColor Red
    break
}

$Report = [System.Collections.Generic.List[Object]]::new()  
ForEach ($Team in $Teams) {
    Write-Host "Checking team: $($Team.DisplayName)" -ForegroundColor Cyan
    Try {
        [array]$Channels = Get-MgTeamChannel -TeamId $Team.Id -All `
            -Property Id, CreatedDateTime, DisplayName, Description, IsArchived, MembershipType -ErrorAction Stop 
        Write-Host "Found $($Channels.Count) channels in team: $($Team.DisplayName)." -ForegroundColor Green
        $Channels = $Channels | Sort-Object DisplayName
    } catch {
        Write-Host "Error fetching channels for team $($Team.DisplayName): $_" -ForegroundColor Red
        continue
    }

    ForEach ($Channel in $Channels) {
        $LastActivityDate = $null
        Write-Host "Checking channel: $($Channel.DisplayName)" -ForegroundColor Yellow
        Try {
            [array]$Messages = Get-MgTeamChannelMessage -TeamId $Team.Id -ChannelId $Channel.Id -Top 20 -ErrorAction Stop 
            if ($Messages.Count -gt 0) {
                [array]$UserMessages = $Messages | Where-Object {$_.body.content -ne "<systemeventmessage/>"}
                if ($UserMessages.Count -gt 0) {
                    $LastActivityDate = $UserMessages[0].CreatedDateTime
                    $LastActivityBy = $UserMessages[0].From.User.DisplayName
                    Write-Host "Last activity date for channel '$($Channel.DisplayName)' is: $LastActivityDate" -ForegroundColor Green
                } else {
                    Write-Host "No user messages found in channel '$($Channel.DisplayName)'. It may be inactive." -ForegroundColor Red
                    $ChannelStatus = "Inactive" 
                }
            } else {
                Write-Host "No messages found in channel '$($Channel.DisplayName)'. It may be inactive." -ForegroundColor Red
                $ChannelStatus = "Inactive"
            }
        } catch {
            Write-Host "Error fetching messages for channel '$($Channel.DisplayName)': $_" -ForegroundColor Red
            continue
        }
        Try {
            [array]$Members = Get-MgTeamChannelMember -TeamId $Team.Id -ChannelId $Channel.Id -All -ErrorAction Stop   
            [array]$Owners = $Members | Where-Object { $_.Roles -contains "owner" }
            [array]$Guests = $Members | Where-Object { $_.UserId -like "*#EXT#*" }
        } catch {
            Write-Host "Error fetching members for channel '$($Channel.DisplayName)': $_" -ForegroundColor Red
            continue
        }

        If ($LastActivityDate) { 
            [int]$DaysSinceLastActivity = (New-TimeSpan -Start $LastActivityDate -End (Get-Date)).Days
            $LastActivityDate = $LastActivityDate.ToString("dd-MMM-yyyy HH:mm")
            If ($DaysSinceLastActivity -gt $ThresholdDays) {
                $ChannelStatus = "Inactive"
            } Else {
                $ChannelStatus = "Active"
            }
        } Else {
            [string]$DaysSinceLastActivity = $null
            $ChannelStatus = "Unknown"
        } 

        If ($Channel.MembershipType -eq "unknownFutureValue") {
            $MembershipType = "Shared"
        } Else {
            $MembershipType = $Channel.MembershipType
        }
        $ReportLine = [PSCustomObject][Ordered]@{
            Team        = $Team.DisplayName
            Channel     = $Channel.DisplayName
            ChannelStatus   = $ChannelStatus
            Description     = $Channel.Description
            LastActivityDate = [string]$LastActivityDate
            LastActivityBy   = $LastActivityBy
            'Days since last activity' = [string]$DaysSinceLastActivity 
            IsArchived      = $Channel.IsArchived
            'Channel type'  = $MembershipType
            'Created date'  = $Channel.CreatedDateTime.ToString("dd-MMM-yyyy HH:mm")
            OwnerCount      = $Owners.Count
            MemberCount     = $Members.Count
            GuestCount      = $Guests.Count
            TeamId          = $Team.Id
            ChannelId       = $Channel.Id
        }
        $Report.Add($ReportLine) | Out-Null
    }
}   

# Basic Statistics
$TotalChannels = $Report.Count
$InactiveChannels = $Report | Where-Object { $_.ChannelStatus -eq "Inactive" } | Measure-Object | Select-Object -ExpandProperty Count
$ActiveChannels = $TotalChannels - $InactiveChannels
$InactivePercentage = if ($TotalChannels -gt 0) { [math]::Round(($InactiveChannels / $TotalChannels) * 100, 2) } else { 0 }
Write-Host "Total channels checked: $TotalChannels" -ForegroundColor Green
Write-Host "Inactive channels: $InactiveChannels ($InactivePercentage%)" -ForegroundColor Green     
Write-Host "Active channels: $ActiveChannels" -ForegroundColor Green   
$AverageDaysSinceLastActivity = If ($InactiveChannels -gt 0) { 
    [math]::Round(($Report | Where-Object { $_.ChannelStatus -eq "Inactive" } | Measure-Object -Property 'Days since last activity' -Average).Average, 2) } else { 0 }
Write-Host "Average days since last activity for inactive channels: $AverageDaysSinceLastActivity" -ForegroundColor Green 
$LongestInactiveChannel = $Report | Where-Object { $_.ChannelStatus -eq "Inactive" } | Sort-Object {$_.'Days since last activity' -as [int]} -Descending | Select-Object -First 1
If ($LongestInactiveChannel) {
    Write-Host ("The longest inactive channel is {0} in the team {1} created on {2} with {3} days since last activity." -f $LongestInactiveChannel.Channel, $LongestInactiveChannel.Team, $LongestInactiveChannel.'Created date', $LongestInactiveChannel.'Days since last activity') -ForegroundColor Green
}   

# Set up to use the ImportExcel module to generate Excel worksheets if it is available. If not, the report will be generated in CSV format instead.
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
}

# Generate the output file in either Excel worksheet or CSV format, depending on if the ImportExcel module is available. If interactive, create
# the file in the user's Downloads folder, otherwise create it in the current directory.
If ($ExcelGenerated) {
    If ($Interactive) {
        $OutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveTeamsChannels.xlsx"
    } Else {
        $OutputFile = "InactiveTeamsChannels.xlsx"
    }
    Remove-Item -Path $OutputFile -ErrorAction SilentlyContinue
    $Report | Export-Excel -Path $OutputFile -WorksheetName "Inactive Teams Channels" -Title ("Inactive Teams Channels {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "InactiveTeamsChannels" -AutoSize -AutoFilter -FreezeTopRow
    If ($Interactive) {
        $OutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\InactiveTeamsChannels.csv"
    } Else {
        $OutputFile = "InactiveTeamsChannels.csv"
    }
    Remove-Item -Path $OutputFile -ErrorAction SilentlyContinue
    $Report | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding Utf8
}

Write-Host "Report generated in: $OutputFile" -ForegroundColor Green

$Attachment = (Get-Item -Path $OutputFile).Name
$EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($OutputFile))

$MsgAttachments = @(
    @{
	"@odata.type" = "#microsoft.graph.fileAttachment"
	Name = ($Attachment -split '\\')[-1]
	ContentType = "application/vnd.ms-excel"
	ContentBytes = $EncodedAttachmentFile
	}
)
 
$ToRecipientList   = @( $EmailRecipient )
[array]$MsgToRecipients = Add-MessageRecipients -ListOfAddresses $ToRecipientList
$MsgSubject = "Inactive Teams Channels Report - $((Get-Date).ToString("dd-MMM-yyyy"))"
$HtmlHead = "<p><h2>Inactive Teams Channels Report</h2><p>"
$HtmlHead += "<p>Total channels checked:<b> $TotalChannels</b></p>"
$HtmlHead += "<p>Inactive channels: <b> $InactiveChannels</b> ($InactivePercentage%)</p>"
$HtmlHead += "<p>Active channels: <b> $ActiveChannels</b></p>"     
$htmlHead += "<p>Average days since last activity for inactive channels: <b>$AverageDaysSinceLastActivity</b></p>"
If ($LongestInactiveChannel) {
    $HtmlHead += ("<p>The channel that has been inactive for the longest period is <i><b>{0}</i></b> in the <b>{1}</b> team. The channel was created on {2} and it is <b><u>{3}</u></b> days since the last noted activity.</p>" -f $LongestInactiveChannel.Channel, $LongestInactiveChannel.Team, $LongestInactiveChannel.'Created date', $LongestInactiveChannel.'Days since last activity')
}   
$htmlHead += "<p>Channels are deemed inactive if users have not posted messages to the channel in the last $ThresholdDays days. This does not mean that other channel resources have not been used in that time. However, you should check the inactive channels and decide to keep, archive, or remove the channels.</p>"
$HtmlHead += "<h3>The following Teams channels have been identified as potentially inactive.</h3></p>"
$HtmlBody = $Report | Where-Object { $_.ChannelStatus -eq "Inactive" } | `
    Select-Object Team, Channel, 'Created date', 'Days since last activity' | ConvertTo-Html -Fragment 
$HtmlMsg = "</body></html><p>" + $HtmlHead + $HtmlBody + "<p>"
# Construct the message body
$MsgBody = @{
  Content = "$($HtmlMsg)"
  ContentType = 'html'  
}

$Message = @{}
$Message.Add("subject", $MsgSubject)
$Message.Add("toRecipients", $MsgToRecipients)
$Message.Add("body", $MsgBody)
$Message.Add("attachments", $MsgAttachments)
$EmailParameters = @{}
$EmailParameters.Add("message", $Message)
$EmailParameters.Add("saveToSentItems", $True)
$EmailParameters.Add("isDeliveryReceiptRequested", $True)

# And send the message using the parameters that we've filled in
Try {
    Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters
    Write-Output ("Message containing information about inactive Teams channels sent to {0}!" -f $EmailRecipient)
} Catch {
    Write-Output ("Error sending email: $_")
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# The Automating Microsoft 365 with PowerShell eBook is available from https://o365itpros.gumroad.com/l/M365PS in EPUB and PDF format 
# or in paperback at https://amzn.to/3L4IBmU

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
