# Report-MailUsageDomains.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-MailUsageDomains.PS1
# An example of how to read the mail usage data for a tenant using the Microsoft Graph PowerShell SDK
# and report user activity for 7, 30, 90, and 180 day intervals. Plus report the domains that send most emails.

# Connect to the Graph
Connect-MgGraph -Scopes ReportSettings.ReadWrite.All, Directory.Read.All, Reports.Read.All

Write-Output "Finding user accounts to process..."
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All | Sort-Object UserPrincipalName

# Check tenant usage data obfuscation setting. If it's true (names are concealed)
# reset it for this script
If ((Get-MgBetaAdminReportSetting).DisplayConcealedNames -eq $True) {
    $Parameters = @{ displayConcealedNames = $False }
    Update-MgBetaAdminReportSetting -BodyParameter $Parameters
    $DisplayConcealedNames = $True
}

Write-Output "Fetching mail activity usage report data..."
Get-MgReportEmailActivityUserDetail -Period 'D7' -OutFile 'c:\temp\d7.csv'
Get-MgReportEmailActivityUserDetail -Period 'D30' -OutFile 'c:\temp\d30.csv'
Get-MgReportEmailActivityUserDetail -Period 'D90' -OutFile 'c:\temp\d90.csv'
Get-MgReportEmailActivityUserDetail -Period 'D180' -OutFile 'c:\temp\d180.csv'
# Import the data into arrays
[array]$D7Data = Import-CSV 'c:\temp\d7.csv' | Sort-Object 'User Principal Name'
[array]$D30Data = Import-CSV 'c:\temp\d30.csv' | Sort-Object 'User Principal Name'
[array]$D90ata = Import-CSV 'c:\temp\d90.csv' | Sort-Object 'User Principal Name'
[array]$D180Data = Import-CSV 'c:\temp\d180.csv' | Sort-Object 'User Principal Name'

# Process mailboxes
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    [array]$D7Email = $D7Data | Where-Object {$_.'User Principal Name' -eq $User.UserPrincipalName}
    [array]$D30Email = $D30Data| Where-Object {$_.'User Principal Name' -eq $User.UserPrincipalName}
    [array]$D90Email = $D90ata | Where-Object {$_.'User Principal Name' -eq $User.UserPrincipalName}
    [array]$D180Email = $D180Data | Where-Object {$_.'User Principal Name' -eq $User.UserPrincipalName}

    If ($D7Email.'Report Refresh Date') {
      $ReportDate = Get-Date($D7Email.'Report Refresh Date') -format dd-MMM-yyyy
    } Else {
      $ReportDate = $Null }

    If ($D7Email.'Last Activity Date') {
      $LastActivityDate = Get-Date($D7Email.'Last Activity Date') -format  dd-MMM-yyyy   
    } Else {                              
      $LastActivityDate = $Null }  
      
    If ([string]::IsNullOrWhiteSpace($User.Mail)) {
       $Domain = $User.UserPrincipalName.Split('@')[1]
    }  Else {
       $Domain = $User.Mail.Split('@')[1]
    }
    
    $ReportLine = [PSCustomObject] @{  
        User            = $User.UserPrincipalName
        Name            = $User.DisplayName      
        'Data Date'     = $ReportDate            
        'Last Activity' = $LastActivityDate      
        'D7 Mail In'    = $D7Email.'Receive Count'
        'D7 Mail Out'   = $D7Email.'Send Count'         
        'D30 Mail In'   = $D30Email.'Receive Count'  
        'D30 Mail Out'  = $D30EMail.'Send Count'     
        'D90 Mail In'   = $D90Email.'Receive Count'   
        'D90 Mail Out'  = $D90Email.'Send Count'      
        'D180 Mail In'  = $D180Email.'Receive Count'  
        'D180 Mail Out' = $D180Email.'Send Count'     
        Domain          = $Domain } 
    $Report.Add($ReportLine)               
} # End Foreach user    
 # Show the data sorted by the volume of outbound mail sent by a user                                
$Report | Sort-Object {$_.'D180 Mail Out' -as [int]} -Descending | Out-GridView  
    
# Reset tenant obfuscation settings to True if that's what they were before
If ($DisplayConcealedNames -eq $True) {
    $Parameters = @{ displayConcealedNames = $True }
    Update-MgBetaAdminReportSetting -BodyParameter $Parameters
    $DisplayConcealedNames = $Null
}

# Analyze domains             
[array]$Domains = Get-MgDomain | Select-Object -ExpandProperty Id 
$OutputData = [System.Collections.Generic.List[Object]]::new()     
ForEach ($Domain in $Domains) {                         
    $DomainData = $Report | Where-Object {$_.Domain -eq $Domain}  
    $DomainSendCount = ($DomainData.'D180 Mail out' | Measure-Object -Sum).Sum 
    $DomainOutput = [PSCustomObject] @{   
       'Domain'     = $Domain          
       'Send Count' = $DomainSendCount } 
    $OutputData.Add($DomainOutput)                                                          
 }                                                                                                        

# Display the domain data
$OutputData | Sort-Object 'Send Count' -Descending

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 
