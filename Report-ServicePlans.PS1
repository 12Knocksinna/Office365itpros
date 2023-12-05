# Report-ServicePlans.PS1
# Script to report the service plans found in subscriptions for a Microsoft 365 tenant
# V1.0 30-Nov-2023
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ServicePlans.PS1

[string]$RunDate = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
$Version = "1.0"
$CSVOutputFile = "c:\temp\Microsoft365LicenseServicePlans.CSV"
$ReportFile = "c:\temp\Microsoft365LicenseServicePlans.html"

# Connect to the Graph and get information about the subscriptions in the tenant
Connect-MgGraph -Scopes Directory.Read.All -NoWelcome
# Get the basic information about tenant subscriptions
[array]$Skus = Get-MgSubscribedSku

# The $ProductInfoDataFile variable points to the CSV file downloaded from Microsoft from
# https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
# It's used to resolve SKU and service plan code names to human-friendly values
Write-Output "Loading product data..."
$ProductInfoDataFile = "C:\Temp\Product names and service plan identifiers for licensing.csv"
If (!(Test-Path -Path $ProductInfoDataFile)) {
    Write-Host "No product information data file available - product and service plan names will not be resolved"
    $ProductData = $false
} Else {
    $ProductData = $true
}

If ($ProductData) {
# If the product data file is available, use it to populate some hash tables to use to resolve SKU and service plan names
    [array]$ProductData = Import-CSV $ProductInfoDataFile
    [array]$ProductInfo = $ProductData | Sort-Object GUID -Unique
    # Create Hash table of the SKUs used in the tenant with the product display names from the Microsoft data file
    $TenantSkuHash = @{}
        ForEach ($P in $SKUs) { 
            $ProductDisplayName = $ProductInfo | Where-Object {$_.GUID -eq $P.SkuId} | `
                Select-Object -ExpandProperty Product_Display_Name
            If ($Null -eq $ProductDisplayName) {
                $ProductDisplayname = $P.SkuPartNumber
            }
            $TenantSkuHash.Add([string]$P.SkuId, [string]$ProductDisplayName) 
        }
# Extract service plan information and build a hash table
    [array]$ServicePlanData = $ProductData | Select-Object Service_Plan_Id, Service_Plan_Name, Service_Plans_Included_Friendly_Names | `
        Sort-Object Service_Plan_Id -Unique
    $ServicePlanHash = @{}
    ForEach ($SP in $ServicePlanData) { 
        $ServicePlanHash.Add([string]$SP.Service_Plan_Id,[string]$SP.Service_Plans_Included_Friendly_Names)
    }
}

# Generate a report about the subscriptions used in the tenant
Write-Host "Generating product subscription information..."
$SkuReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Sku in $Skus) {
    $AvailableUnits = ($Sku.PrepaidUnits.Enabled - $Sku.ConsumedUnits)
    If ($ProductData) {
        $SkuDisplayName = $TenantSkuHash[$Sku.SkuId]
    } Else {
        $SkuDisplayName = $Sku.SkuPartNumber
    }
    $DataLine = [PSCustomObject][Ordered]@{
        'Sku Part Number'   = $SkuDisplayName
        SkuId               = $Sku.SkuId
        'Active Units'      = $Sku.PrepaidUnits.Enabled
        'Warning Units'     = $Sku.PrepaidUnits.Warning
        'Consumed Units'    = $Sku.ConsumedUnits
        'Available Units'   = $AvailableUnits
    }
    $SkuReport.Add($Dataline)
}
# Get the renewal data
$Uri = "https://graph.microsoft.com/beta/directory/subscriptions"
[array]$SkuData = Invoke-MgGraphRequest -Uri $Uri -Method Get
# Put the renewal information into a hash table
$SkuHash = @{}
ForEach ($Sku in $SkuData.Value) { $SkuHash.Add($Sku.SkuId,$Sku.nextLifecycleDateTime) }
# Update the report with the renewal information
ForEach ($R in $SkuReport) {
  $DaysToRenew = $Null
  $SkuRenewalDate = $SkuHash[$R.SkuId]
  If ($SkuRenewalDate) {
    $SkuRenewalDate = (Get-Date $SkuRenewalDate -format "dd-MMM-yyy")
  }
  $R | Add-Member -NotePropertyName "Renewal date" -NotePropertyValue $SkuRenewalDate -Force 
  If ($SkuRenewalDate) {
   $DaysToRenew = (New-TimeSpan $SkuRenewalDate).Days
   $R | Add-Member -NotePropertyName "Days to renewal" -NotePropertyValue $DaysToRenew -Force 
 }
}

# Now process the service plans in the subscriptions
Write-Host "Extracting service plan information..."
$ServicePlanReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($Sku in $Skus) {
    [array]$ServicePlans = $Sku.ServicePlans
    ForEach ($SP in $ServicePlans) {
        If ($ProductData) {
            $SkuDisplayName = $TenantSkuHash[$Sku.SkuId]
            $ServicePlanDisplayName = $ServicePlanHash[$SP.servicePlanId]
        } Else {
            $SkuDisplayName = $Sku.SkuPartNumber
            $ServicePlanDisplayName = $SP.servicePlanName
        }
        $SPDataLine = [PSCustomObject][Ordered]@{
            SkuId               = $Sku.SkuId
            Sku                 = $Sku.SkuPartNumber
            'SKU Name'          = $SkuDisplayName
            'Service Plan'      = $SP.servicePlanId
            'Service Plan Name' = $ServicePlanDisplayName
        }
        $ServicePlanReport.Add($SPDataLine)
    }
}

Write-Host "Generating report..."
$OrgName  = (Get-MgOrganization).DisplayName
# Create the HTML report. First, define the header.
$HTMLHead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.pass{background: #B7EB83;}
	   td.warn{background: #E3242B;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Microsoft 365 Subsriptions and Service Plan Report</h1></p>
           <p><h2><b>For the " + $Orgname + " tenant</b></h2></p>
           <p><h3>Generated: " + $RunDate + "</h3></p></div>"

# This section highlights subscriptions that have less than 3 remaining licenses.
# Idea from https://stackoverflow.com/questions/37662940/convertto-html-highlight-the-cells-with-special-values
# First, convert the output SKU Report to HTML and then import it into an XML structure
$HTMLTable = $SkuReport | ConvertTo-Html -Fragment
[xml]$XML = $HTMLTable
# Create an attribute class to use, name it, and append to the XML table attributes
$TableClass = $XML.CreateAttribute("class")
$TableClass.Value = "AvailableUnits"
$XML.table.Attributes.Append($TableClass) | Out-Null
# Conditional formatting for the table rows. The number of available units is in table row 6, so we update td[5]
ForEach ($TableRow in $XML.table.SelectNodes("tr")) {
    # each TR becomes a member of class "tablerow"
    $TableRow.SetAttribute("class","tablerow")
    ## If row has TD and TD[5] is 3 or less
    If (($TableRow.td) -and ([int]$TableRow.td[5] -le 3))  {
        ## tag the TD with the class "warn" defined in the heading
        $TableRow.SelectNodes("td")[5].SetAttribute("class","warn")
    }
}
# Wrap the output table with a div tag
$HTMLBody = [string]::Format('<div class="tablediv">{0}</div>',$XML.OuterXml)

[string]$HTMLSkuSeparator = "<p><h2>Service Plans in Microsoft 365 Subscriptions</h2></p>"
[string]$HTMLSkuOutput = $null

# For each SKU, extract its service plans from the list created earlier and convert it into a HTML segment
ForEach ($Sku in $Skus) {
    If ($ProductData) {
        $SkuDisplayName = $TenantSkuHash[$Sku.SkuId]
    } Else {
        $SkuDisplayName = $Sku.SkuPartNumber
    }
    $SkuServicePlans = $ServicePlanReport | Where-Object {$_.SkuId -eq $Sku.SkuId} | Sort-Object 'Service Plan Name'
    $HTMLSkuHeader = "<h3>Service Plans for product <b>" + $SkuDisplayName + "<b></h3><p>"
    $HTMLSkuContent = $SkuServicePlans | ConvertTo-HTML -Fragment
    $HTMLSkuOutput = $HTMLSkuOutput + $HTMLSkuHeader + $HTMLSkuContent + "<p><p>"
}

# End stuff to output
$HTMLTail = "<p>Report created for the " + $OrgName + " tenant on " + $RunDate + "<p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+  
             "<p>Number of subscriptions found:             " + $SkuReport.Count + "</p>" +
             "<p>-----------------------------------------------------------------------------------------------------------------------------</p>"+
             "<p>Microsoft 365 Subscriptions and Service Plan Report<b> " + $Version + "</b>"	

$HTMLReport = $HTMLHead + $HTMLBody + $HTMLSkuSeparator + $HTMLSkuOutput + $HTMLtail
$HTMLReport | Out-File $ReportFile  -Encoding UTF8

$SkuReport | Export-CSV -NoTypeInformation $CSVOutputFile

Write-Host ""
Write-Host "All done. Output files are" $CSVOutputFile "and" $ReportFile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
