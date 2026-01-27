# Report-AdaptiveScopes.PS1
# Demonstrate the use of the Get-AdaptiveScope and Get-AdaptiveScopeMembers cmdlets to report on adaptive scopes for Purview

# V1.0 27-Jan-2026
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-AdaptiveScopes.PS1

[array]$Modules = Get-Module | Select-Object -ExpandProperty Name
If (!($Modules -contains "ExchangeOnlineManagement")) {
    Connect-ExchangeOnline -ShowBanner:$false
}

Write-Output "Finding adaptive scopes in the organization..."
[array]$AdaptiveScopes = Get-AdaptiveScope | Sort-Object Name

If ($AdaptiveScopes) {
    Write-Output ("Preparing to analyze {0} adaptive scopes..." -f $AdaptiveScopes.count)
} Else {
    Write-Output "No adaptive scopes found in the organization - exiting"
    Break
}

$Report = [System.Collections.Generic.List[Object]]::new()
$ReportDetail = [System.Collections.Generic.List[Object]]::new()
ForEach ($Scope in $AdaptiveScopes) {
    Write-Output ("Processing scope {0}..." -f $Scope.Name)
    [array]$Members = Get-AdaptiveScopeMembers -Identity $Scope.Name -ErrorAction SilentlyContinue
    If ($Members) {
        # Only take non-removed members into account for the count, and drop 1 to account for the summary entry
        $MemberCount = ($Members | Where-Object {$_.State -ne "Removed"} | Measure-Object | Select-Object -ExpandProperty Count) - 1
    } Else {
        $MemberCount = 0
    }

    If ($Scope.FilterConditions) {
        $Filter = $Scope.FilterConditions
    } Else {
        $Filter = $Scope.RawQuery
    }

    $ReportLine = [PSCustomObject] @{
        Name                = $Scope.Name
        Workload            = $Scope.Workload
        Description         = $Scope.Comment
        filter              = $Filter
        CreatedBy           = $Scope.CreatedBy
        CreatedDate         = Get-Date $Scope.WhenCreated -format 'dd-MMM-yyyy HH:mm'
        LastModifiedBy      = $Scope.LastModifiedBy
        LastModifiedDate    = Get-Date $Scope.WhenChanged -format 'dd-MMM-yyyy HH:mm'
        MemberCount         = $MemberCount
        Id                  = $Scope.ImmutableId
    }
    $Report.Add($ReportLine)

    [int]$i = 0  
    ForEach ($Member in $Members) {       
        If ($i -eq 0 -or $Member.State -eq "Removed") { 
            # Skip the first entry (it's the summary) and any removed members
            $i++
            Continue 
        } Else {
            $Warning = $null
            # If a mailbox or group, check that it still exists
            Switch ($Member.LocationType) {
                "Group" {
                    $Group = Get-UnifiedGroup -Identity $Member.ObjectId -ErrorAction SilentlyContinue
                    If (!($Group)) {
                        # Group no longer exists
                        $Warning = "Microsoft 365 Group cannot be found"
                    }
                }
                "User" {
                    $Mailbox = Get-Mailbox -Identity $Member.ObjectId -ErrorAction SilentlyContinue
                    If (!($Mailbox)) {
                    # Mailbox no longer exists
                        $Warning = "Mailbox cannot be found"
                    }
                }
                Default {
                    # Do nothing for other types
                }
            }
            # Get the timestamp and make it into a date
            [string]$AddedDateString = $Member.EventDateTime.Split("+")[0].trim()
            [datetime]$ParsedEventDate = [DateTime]::ParseExact($AddedDateString, 'MM/dd/yyyy HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)

            $DetailLine = [PSCustomObject] @{
                ScopeName           = $Scope.Name
                MemberType          = $Member.LocationType
                MemberIdentity      = $Member.ObjectId
                DisplayName         = $Member.DisplayName
                UserPrincipalName   = $Member.Upn
                DateAdded           = Get-Date $ParsedEventDate -format 'dd-MMM-yyyy HH:mm'
                Notes               = $Warning
            }
            $ReportDetail.Add($DetailLine)   
        }
    }
}

$ReportFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\AdaptiveScopesReport.html"

$Rundate = Get-Date -format 'dd-MMM-yyyy HH:mm'
$Orgname = (Get-OrganizationConfig).Name
$HTMLHead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 32px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 26px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 20px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
       H4{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.admin{background: #B7EB83;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Adaptive Scopes Report</h1></p>
           <p><h2><b>For the " + $Orgname + " tenant</b></h2></p>
           <p><h3>Generated: " + $RunDate + "</h3></p></div>"

$HtmlReport = "<body>"
ForEach ($Scope in $Report) {
    # Generate a HTML section for each scope
    $HtmlScope = "<h2>Scope: " + $Scope.Name + " (Members: " + $Scope.MemberCount + ")</h2>"
    $HtmlScope = $HtmlScope + "<p><b><h3>Adaptive Scope Properties:</h3></b><p>Workload: " + $Scope.Workload + `
     "<br>Description: " + $Scope.Description + "<br>Created by: <b>" + $Scope.CreatedBy + "</b> on <b>" + $($Scope.CreatedDate) + `
     "</b><br>Last Modified By: <b>" + $Scope.LastModifiedBy + "</b> on <b>" + $($Scope.LastModifiedDate) + "</b><br>Filter Conditions: " +  $Scope.filter + "</p>"
    
    # Include the membership details
    $ScopeMembers = $ReportDetail | Where-Object {$_.ScopeName -eq $Scope.Name}
    If ($ScopeMembers) {
        $HtmlScope = $HtmlScope + "<br><b>Members:</b><br>"
        $HtmlScope = $HtmlScope + ($ScopeMembers | Select-Object MemberType, DisplayName, UserPrincipalName, DateAdded, Notes | ConvertTo-Html -Fragment -As Table)
    } Else {
        $HtmlScope = $HtmlScope + "<br><i>No members found for this adaptive scope</i><br>"
    }
    $HtmlReport = $HtmlReport + $HtmlScope
}   

$HtmlReport = $HtmlHead + $HtmlReport + "</body></html>"
$HtmlReport | Out-File -FilePath $ReportFile -Encoding UTF8

Write-Output ("Report saved to {0}" -f $ReportFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 