Write-Output "Finding user accounts to process..."
[array]$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All | Sort-Object UserPrincipalName

Write-Output "Fetching mail activiy usage report data..."
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
        Domain         = $User.Mail.Split('@')[1]}    
    $Report.Add($ReportLine)               
    }                                  
    $Report | Out-GridView         
# Analyze domains             
[array]$Domains = Get-MgDomain | Select-Object -ExpandProperty Id 
$OutputCounts = [System.Collections.Generic.List[Object]]::new()     
ForEach ($Domain in $Domains) {                         
    $DomainData = $Report | Where-Object {$_.Domain -eq $Domain}  
    $DomainSendCount = ($DomainData.'D180 Mail out' | Measure-Object -Sum).Sum 
    $DomainOutput = [PSCustomObject] @{   
       'Domain'     = $Domain          
       'Send Count' = $DomainSendCount } 
    $OutputCounts.Add($DomainOutput)                                                          
 }                                                         >                                               