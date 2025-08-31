# Report-LastAppAccessedByUsers.PS1
# Use the Entra ID sign-in audit log and the Microsoft 365 audit log to report the last Microsoft 365 app
# accessed by users. The script generates a report in HTML and either CSV or Excel format. If run interactively,
# the script outputs the location of the reports to the console. If run in Azure Automation, the script emails the reports
# to a specified recipient.

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-LastAppAccessedByUsers.PS1

# V1.0 27-Aug-2025

Param
(
  [Parameter (Mandatory= $true)]
  [string]$DestinationEmailAddress
)

Function Update-MessageAttachments {
    [cmdletbinding()]
    Param(
        [array]$ListOfAttachments 
    )
    # Function to create an array of message attachments
    [array]$MsgAttachments = $null
    ForEach ($File in $ListOfAttachments) {
        $ConvertedContent = [Convert]::ToBase64String([IO.File]::ReadAllBytes($File))
        $FileExtension = [System.IO.Path]::GetExtension($File) 
        Switch ($FileExtension) {
            ".pdf" {
                $ContentType = "application/pdf"
            }
            ".docx" {
                $ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            }
            ".xlsx" {
                $ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }   
            ".pptx" {
                $ContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            }   
            ".jpg" {
                $ContentType = "image/jpeg"
            }   
            ".png" {
                $ContentType = "image/png"
            }   
            default {
                $ContentType = "application/octet-stream"
            }
        }
        $AttachmentDetails = @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            Name = $File
            ContentType = $ContentType
            ContentBytes = $ConvertedContent
        }
        $MsgAttachments += $AttachmentDetails
    }
    Return $MsgAttachments  
}

# Flag to let script code know if we're running interactively or within Azure Automation
$Interactive = $false

If ([Environment]::UserInteractive) { 
    # We're running interactively...
    Write-Host "Script running interactively... connecting to the Graph" -ForegroundColor Yellow
    Connect-MgGraph -NoWelcome
    $Interactive = $true
    [array]$Modules = Get-Module | Select-Object -ExpandProperty Name
    If ("ExchangeOnlineManagement" -Notin $Modules) {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowBanner:$false 
    }
} Else { 
    # We're not, so likely in Azure Automation
    Write-Host "Running the script to identify the last app accessed by Users" 
    Connect-MgGraph -Identity
    $Tenant = Get-MgOrganization
    # Connect with a managed identity
    $TenantDomain = $Tenant.VerifiedDomains | Where-Object {$_.isDefault -eq $true} | Select-Object -ExpandProperty Name
    Connect-ExchangeOnline -ManagedIdentity -Organization $TenantDomain
    $CurrentFolder = (Get-Location).Path
}

# Check that we have the right permissions - in Azure Automation, we assume that the automation account has the right permissions
If ($Interactive) {
    [string[]]$CurrentScopes = (Get-MgContext).Scopes
    [string[]]$RequiredScopes = @('AuditLog.Read.All','User.Read.All','Mail.Send')

    $CheckScopes =[object[]][Linq.Enumerable]::Intersect($RequiredScopes,$CurrentScopes)
    If ($CheckScopes.Count -ne 3) { 
        Write-Host ("To run this script, you need to connect to Microsoft Graph with the following scopes: {0}" -f $RequiredScopes) -ForegroundColor Red
        Disconnect-Graph
        Break
    }
}

# Make sure that we have a good email address to send the output message to
 If (!($DestinationEmailAddress)) {
        $DestinationEmailAddress = 'Ken.Bowers@office365itpros.com'
}

# Get licensed member accounts for the tenant, excluding utility accounts
[array]$Users = Get-MgUser -All -PageSize 500 -Filter "usertype eq 'member' and accountenabled eq true and employeeType ne 'Utility' and assignedLicenses/`$count ne 0" `
    -Sort displayName -ConsistencyLevel Eventual -CountVariable Count -Property DisplayName, UserPrincipalName, id, employeeType, userType, signinactivity, accountEnabled

# $Users = $Users | Where-Object {$_.DisplayName -NotMatch "admin|test|service|breakglass|break glass|monitor|audit|helpdesk|it dept|itdept|it team|itteam|exchange admin|exchangeadmin|sharepoint admin|sharepointadmin|onedrive admin|onedriveadmin|teams admin|teamsadmin"}

# Switch to control whether we check Entra ID sign-in logs or just use the signinActivity property of the user object
$CheckEntra = $true

If ($Users) {
    Write-Host "Found $($Users.Count) user accounts to process."
    $Report = [System.Collections.Generic.List[Object]]::new()
} Else {
    Write-Host "No user accounts found to process."
    Break
}

# Loop through the users to see what we can find
[datetime]$StartProcessing = Get-Date
[int]$i = 0
ForEach ($User in $Users) {
    $i++
    Write-Host ("Checking sign-in records and audit records for {0} {1}/{2}" -f $User.DisplayName, $i, $Users.Count)
    $LastAppSignIn = $null; $LastAppSignInDate = $null; $AuditLogTimeStamp = $null; $AuditLogApps = $null; $UserId = $null; $AuditLogActions = $null; $AuditLogTimeStamp = $null
    [array]$AuditLogAppsPrint = $null
    [array]$Logs = $null; [array]$AuditLogs = $null
    If ($CheckEntra) {
        # Check Entra ID sign-in logs
        $UserId = $User.Id
        [array]$Logs = Get-MgAuditLogSignIn -Filter "UserId eq '$UserId'" -Top 3
        If ($Logs) {
            [array]$AppNames = $Logs.AppDisplayName | Sort-Object -Unique
            $LastAppSignIn = $AppNames -join "; "
            $LastAppSignInDate = Get-Date ($Logs[0].CreatedDateTime) -format 'dd-MMM-yyyy HH:mm:ss'
        } Else {
            $LastAppSignIn = "No Entra ID sign-in records found"
            $LastAppSignInDate = "No sign-ins over the last 30 days"
        }
    } Else { # Rely on the signinActivity property of the user object
        $LastAppSignIn = "Entra ID sign-in records not checked"
        If ($User.signinActivity.lastSignInDateTime) {
        $LastAppSignInDate = Get-Date ($User.signinActivity.lastSignInDateTime) -format 'dd-MMM-yyyy HH:mm:ss'
        } Else {
            $LastAppSignInDate = "Never signed in"
        }
    }
    # Check Microsoft 365 audit log
    [array]$AuditLogs = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date) -UserIds $User.UserPrincipalName -ResultSize 3 -Formatted  
    If ($AuditLogs) {
        [array]$AuditLogApps = $AuditLogs.RecordType | Sort-Object -Unique
        [array]$AuditLogActions = $AuditLogs.Operations | Sort-Object -Unique
        $AuditLogTimeStamp = Get-Date ($AuditLogs[0].CreationDate) -format 'dd-MMM-yyyy HH:mm:ss'
       
        ForEach ($App in $AuditLogApps) {
            # Modify the record type contained in the audit records to make it more understandable - Audit log record types are defined at https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-schema?WT.mc_id=M365-MVP-9501#auditlogrecordtype
            # Not all record types are covered here, so feel free to add more as needed
            Switch -Wildcard ($App) {
                "Exchang*" {            
                    $AuditLogAppsPrint = "Exchange Online"
                }
                "AzureActiveDirectory*" {
                    $AuditLogAppsPrint += "Entra ID"
                }
                "SharePoint*" {
                    $AuditLogAppsPrint += "SharePoint Online"
                }
                "CopilotInteraction" {
                    $AuditLogAppsPrint += "Microsoft 365 Copilot"
                }   
                "MicrosoftTeams" {
                    $AuditLogAppsPrint += "Microsoft Teams"
                }
                "MIP*" {
                    $AuditLogAppsPrint += "Microsoft Information Protection"
                }
                "Compliance*" {
                    $AuditLogAppsPrint += "Microsoft Purview Compliance"
                }
                "Planner*" {
                    $AuditLogAppsPrint += "Microsoft Planner"
                }
                "Power*" {
                    $AuditLogAppsPrint += "Microsoft Power Platform"
                }
                "Project*" {
                    $AuditLogAppsPrint += "Microsoft Project"
                }
            }   
        }
    }

    If ($Logs.count -gt 0 -or $AuditLogs.count -gt 0) {
        $Report.Add([PSCustomObject][ordered]@{
            UserPrincipalName       = $User.UserPrincipalName
            Name                    = $User.DisplayName
            LastAppSignIn           = $LastAppSignIn
            LastEntraSignInDate     = $LastAppSignInDate
            AuditLogApp             = $AuditLogAppsPrint -join ", "
            AuditLogAction          = $AuditLogActions -join ", "
            AuditLogTimeStamp       = $AuditLogTimeStamp
        })
    }
}
[datetime]$EndProcessing = Get-Date
$TimeRequired = $EndProcessing - $StartProcessing
$Minutes = [math]::Floor($TimeRequired.TotalSeconds / 60)
$Seconds = [math]::Round($TimeRequired.TotalSeconds % 60, 2)
If ($Interactive) {
    Write-Host ("Total processing time for {0} users: {1}m {2}s" -f $users.count, $Minutes, $Seconds) -ForegroundColor Cyan
    Write-Host ("Average required per user {0} seconds" -f [math]::Round($TimeRequired.TotalSeconds / $users.count, 2)) -ForegroundColor Cyan
} Else {
    Write-Output ("Total processing time for {0} users: {1}m {2}s" -f $users.count, $Minutes, $Seconds) 
    Write-Output ("Average required per user {0} seconds" -f [math]::Round($TimeRequired.TotalSeconds / $users.count, 2)) 
}

# Create a nice HTML report
# Generate sortable HTML table with type-aware sorting - use number as the type for numeric values, date for dates, and string for text
$HtmlHeader = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Last App Access by User Accounts</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; background: #f4f6f8; color: #222; }
h1 { background: #0078d4; color: #fff; padding: 16px; border-radius: 6px 6px 0 0; margin-bottom: 20px; }
table { width: 100%; background: #fff; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-collapse: collapse; }
th, td { padding: 12px; text-align: left; }
th { background: #e5eaf1; cursor: pointer; position: relative; }
th:hover { background: #d0e7fa; }
th::after { content: '↕'; position: absolute; right: 8px; opacity: 0.5; }
tr:nth-child(even) { background: #f0f4fa; }
tr:hover { background: #d0e7fa; }
</style>
<script>
function parseValue(val, type) {
    if(type === 'number') return parseFloat(val.replace(/,/g,'')) || 0;
    if(type === 'date') return new Date(val);
    return val.toLowerCase();
}
function sortTable(n, type) {
    var table = document.getElementById('UserAppStats');
    var rows = Array.from(table.rows).slice(1);
    var dir = table.getAttribute('data-sortdir'+n) === 'asc' ? 'desc' : 'asc';
    rows.sort(function(a, b) {
        var x = parseValue(a.cells[n].innerText, type);
        var y = parseValue(b.cells[n].innerText, type);
        if(x < y) return dir === 'asc' ? -1 : 1;
        if(x > y) return dir === 'asc' ? 1 : -1;
        return 0;
    });
    rows.forEach(function(row) { table.tBodies[0].appendChild(row); });
    table.setAttribute('data-sortdir'+n, dir);
}
</script>
</head>
<body>
<h1>Last App Access by Users</h1>
<table id="UserAppStats">
<thead>
<tr>
<th onclick="sortTable(0,'string')">UserPrincipalName</th>
<th onclick="sortTable(1,'string')">Name</th>
<th onclick="sortTable(2,'string')">LastAppSignIn</th>
<th onclick="sortTable(3,'date')">LastEntraSignInDate</th>
<th onclick="sortTable(4,'string')">AuditLogApp</th>
<th onclick="sortTable(5,'string')">AuditLogAction</th>
<th onclick="sortTable(6,'date')">AuditLogTimeStamp</th>
</tr>
</thead>
<tbody>
"@

$HtmlRows = foreach ($Row in $Report ) {
    "<tr><td>$($row.UserPrincipalName)</td><td>$($row.Name)</td><td>$($row.LastAppSignIn)</td><td>$($row.LastEntraSignInDate)</td><td>$($row.AuditLogApp)</td><td>$($row.AuditLogAction)</td><td>$($row.AuditLogTimeStamp)</td></tr>"
}

$HtmlFooter = @"
</tbody>
</table>
</body>
</html>
"@

#Generate the full HTML content and save it to a file
If ($Interactive) {
    $HTMLReportFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\LastAppAccessbyUsers.html"
} Else {
    $HTMLReportFile = $CurrentFolder + "\LastAppAccessbyUsers.html"
}

$HTMLFile = $HtmlHeader + ($HtmlRows -join "`n") + $HtmlFooter
$HTMLFile | Out-File -FilePath $HTMLReportFile -Encoding utf8

If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    If ($Interactive) {
        $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\LastAppAccessbyUsers.xlsx"
        If (Test-Path $ExcelOutputFile) {
            Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
        } 
    } Else {
        # Simple file name when run in Azure Automation
        $ExcelOutputFile = $CurrentFolder +  "\LastAppAccessbyUsers.xlsx"
    }
    $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName "Last App Access byUsers" -Title ("Last App Access by Users Report {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "LastAppAccessByUsers" 
} Else {
    If ($Interactive) {
        $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\LastAppAccessbyUsers.CSV"
    } Else {
        $CSVOutputFile = $CurrentFolder + "\LastAppAccessbyUsers.CSV"
    }
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}
 
# If we're running interactively, the script outputs details to the console and tells the user where to find the reports. If not,
# we assume that the code is running in Azure Automation and we send an email with the reports attached.
If ($Interactive) {
    If ($ExcelGenerated) {
        Write-Host ("An Excel report about the last app access by users is available in {0}" -f $ExcelOutputFile)
    } Else {
        Write-Host ("A CSV report about the last app access by users is available in {0}" -f $CSVOutputFile)
    }
    Write-Host "Data for last app access by users is also available in a HTML format at $HTMLReportFile" 

} Else {
          
    If ($ExcelGenerated) {
        [array]$InputAttachments = $ExcelOutputFile, $HTMLReportFile  
    } Else {
        [array]$InputAttachments = $HTMLReportFile, $CSVOutputFile
    }

    [array]$MsgAttachments = Update-MessageAttachments -ListOfAttachments $InputAttachments
    # Define who’s sending the message - change this to an appropriate account in your tenant
    $MsgFrom = 'Customer.Services@office365itpros.com'
    # Build the array of a single TO recipient detailed in a hash table - change this to the appropriate recipient for your tenant
    $ToRecipient = @{}
    $ToRecipient.Add("emailAddress",@{'address'=$DestinationEmailAddress})
    [array]$MsgTo = $ToRecipient
    # Define the message subject
    $MsgSubject = "Important: Last Apps Accessed by Users Report"
    # Create the HTML content
    $HtmlMsg = "</body></html><p>The output files for the <b>Last Apps Accessed by Users Report</b> are attached to this message. Please review the information at your convenience</p>"
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
        Send-MgUserMail -UserId $MsgFrom -BodyParameter $EmailParameters -ErrorAction Stop
        Write-Output ("Last User Access to Apps reports emailed to {0}" -f $ToRecipient.emailAddress.address)
    } Catch {
        Write-Output "Unable to send email"
    }
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 