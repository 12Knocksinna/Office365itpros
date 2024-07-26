# ReportDLMembershipsCounts-MgGraph.PS1
# A script to demonstrate how to use the Microsoft Graph PowerShell SDK to report Exchange Online distribution list members
# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/ReportDLMembershipsCounts-MgGraph.PS1

Function Get-DLOwners {
[cmdletbinding()]
# Return information about DL Owners
param(
   [parameter(Mandatory=$True,ValueFromPipeline=$true)]
   $DLManagedBy
)

# Unpack the set of DL owners and make sure that they resplve into display names
$Users = [System.Collections.Generic.List[Object]]::new()
ForEach ($O in $DLManagedBy) {
   $User = Get-ExoRecipient -Identity $O -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DisplayName
   If ($User) { 
      $UserLine = [PSCustomObject][Ordered]@{  
       User  = $User }
      $Users.Add($UserLine) }
}
 
$OwnerNames = $Users.User -join ", "
Return $OwnerNames
}

# Connect to Exchange Online
$Modules = Get-Module | Select-Object -ExpandProperty Name
If ('ExchangeOnlineManagement' -notin $Modules) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -SkipLoadingCmdletHelp
}

# Connect to the Microsoft Graph
Connect-MgGraph -Scope Group.Read.All, GroupMember.Read.All

# Find all distribution lists - this is most easily done with Exchange Online - make sure that we exclude room lists
Write-Host "Finding Exchange Online Distribution Lists..."
$DLs = Get-DistributionGroup -ResultSize Unlimited -Filter {RecipientTypeDetails -ne "Roomlist"} | Select-Object DisplayName, ExternalDirectoryObjectId, ManagedBy
If (!($DLs)) { 
   Write-Host "No distribution lists found... sorry! "; break 
 } Else { 
   Write-Host ("{0} distribution lists found" -f $DLs.count) 
}

$Report = [System.Collections.Generic.List[Object]]::new()
$DLCSVOutput = "c:\temp\DLMemberCounts.CSV"
[int]$DLNumber = 0
ForEach ($DL in $DLs) { 
  $DLNumber++
  $ProgressBar = "Processing distribution list " + $DL.DisplayName + " (" + $DLNumber + " of " + $DLs.Count + ")" 
  Write-Progress -Activity "Analzying membership of distribution list " -Status $ProgressBar -PercentComplete ($DLNumber/$DLs.Count*100)
    [array]$Members = Get-MgGroupTransitiveMember -GroupId $DL.ExternalDirectoryObjectId
    [array]$MemberData = $Members.AdditionalProperties
    If ($MemberData.Count -eq 0) { # For whatever reason, no members in this group, so we zeroize everything to prepare for the next group
        $CountOfMembers = 0
        [array]$TenantMembers = @() 
        [array]$GroupMembers = @() 
        [array]$GuestMembers = @() 
        [array]$ContactMembers = @()
        $MemberNames = $Null   
     } Else  { 
       [int]$CountOfMembers = $MemberData.Count
       [array]$TenantMembers = $MemberData | Where-Object { $_.'@odata.type' -eq "#microsoft.graph.user" -and $_.userType -eq 'Member' }
       [array]$GuestMembers = $MemberData | Where-Object { $_.'@odata.type' -eq "#microsoft.graph.user" -and $_.userType -eq 'Guest' }
       [array]$ContactMembers = $MemberData | Where-Object { $_.'@odata.type' -eq "#microsoft.graph.orgContact" }
       [array]$GroupMembers = $MemberData | Where-Object { $_.'@odata.type' -eq "#microsoft.graph.group" }
       [string]$MemberNames = $MemberData.displayName -join ", "
    }
    $DLOwners = Get-DLOwners $DL.ManagedBy
    $ReportLine = [PSCustomObject][Ordered]@{  
       DLName             = $DL.DisplayName
       ManagedBy          = $DLOwners
       "Members"          = $CountOfMembers
       "Tenant Members"   = $TenantMembers.Count
       "Guest Members"    = $GuestMembers.Count
       "Group Members"    = $GroupMembers.Count
       "Mail Contacts"    = $ContactMembers.Count
       "Member names"     = $MemberNames
       Identity           = $DL.ExternalDirectoryObjectId 
      }
    $Report.Add($ReportLine) 
}

# Create output files (HTML and CSV)
$DLCSVOutput = "c:\temp\DLData.CSV"
$ReportFile = "c:\temp\DLData.html"
# Create the HTML report
$OrgName = (Get-MgOrganization).DisplayName
$CreationDate = Get-Date -format dd-MMM-yyyy
$Version = "1.0"
$htmlhead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.pass{background: #B7EB83;}
	   td.warn{background: #FFF275;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Distribution List Manager Report</h1></p>
           <p><h2><b>For the " + $Orgname + " organization</b></h2></p>
           <p><h3>Generated: " + (Get-Date -format g) + "</h3></p></div>"

$htmlbody1 = $Report | ConvertTo-Html -Fragment

$htmltail = "<p>Report created for: " + $OrgName + "</p>" +
             "<p>Created: " + $CreationDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of distribution lists found:    " + $DLs.Count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Distribution List Manager Report<b> " + $Version + "</b>"	

$htmlreport = $htmlhead + $htmlbody1 + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8

Write-Host ("All done. {0} distribution lists analyzed. CSV file is available at {1} and a HTML report at {2}" -f $DLs.Count, $DLCSVOutput, $ReportFile)
$Report | Out-GridView
$Report | Export-CSV -NoTypeInformation $DLCSVOutput

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
