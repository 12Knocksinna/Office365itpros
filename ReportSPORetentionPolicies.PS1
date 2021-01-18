# ReportSPORetentionPolicies.PS1
# Example for Chapter 19 of Office 365 for IT Pros
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportSPORetentionPolicies.PS1
# Must connect to the Compliance Center endpoint to make this script work!
$Report = [System.Collections.Generic.List[Object]]::new()
# Fetch a set of retention policies that apply to SharePoint and aren't to publish labels
$Policies = (Get-RetentionCompliancePolicy -ExcludeTeamsPolicy -DistributionDetail -RetentionRuleTypes | ? {$_.SharePointLocation -ne $Null -and $_.RetentionRuleTypes -ne "Publish"}) 
# Process Each policy
ForEach ($P in $Policies) {
        $Duration = $Null
        Write-Host "Processing retention policy" $P.Name
        $Rule = Get-RetentionComplianceRule -Policy $P.Name 
        $Settings = "Simple"
        $Duration = $Rule.RetentionDuration
        # Check whether a rule is for advanced settings - either a KQL query or sensitive data types
        If (-not [string]::IsNullOrEmpty($Rule.ContentMatchQuery) -and -not [string]::IsNullOrWhiteSpace($Rule.ContentMatchQuery)) {
              $Settings = "Advanced/KQL" }
        Elseif (-not [string]::IsNullOrEmpty($Rule.ContentContainsSensitiveInformation) -and -not [string]::IsNullOrEmpty($Rule.ContentContainsSensitiveInformation)) {
             $Settings = "Advanced/Sensitive Data" }
        # Handle retention policy that retains and doesn’t do anything else
        If ($Rule.RetentionDuration -eq $Null -and $Rule.ApplyComplianceTag -ne $Null) {
           $Duration = (Get-ComplianceTag -Identity $Rule.ApplyComplianceTag | Select -Expandproperty RetentionDuration) }
        $RetentionAction = $Rule.RetentionComplianceAction
        If ([string]::IsNullOrEmpty($RetentionAction)) {
           $RetentionAction = "Retain" }
        If ($P.SharePointLocation.Name -eq "All") {
              $ReportLine = [PSCustomObject]@{
              PolicyName        = $P.Name
              SiteName          = "All SharePoint Sites"
              SiteURL           = "All SharePoint Sites"
              RetentionTime     = $Rule.RetentionDurationDisplayHint
              RetentionDuration = $Duration
              RetentionAction   = $RetentionAction 
              Settings           = $Settings}
             $Report.Add($ReportLine) } 
        If ($P.SharePointLocationException -ne $Null) {
               $Locations = ($P | Select -ExpandProperty SharePointLocationException)
               ForEach ($L in $Locations) {
                  $Exception = "*Exclude* " + $L.DisplayName
                  $ReportLine = [PSCustomObject]@{
                    PolicyName = $P.Name
                    SiteName   = $Exception
                    SiteURL    = $L.Name }
                $Report.Add($ReportLine) }
        }
        ElseIf ($P.SharePointLocation.Name -ne "All") {
           $Locations = ($P | Select -ExpandProperty SharePointLocation)
           ForEach ($L in $Locations) {
               $ReportLine = [PSCustomObject]@{
                  PolicyName        = $P.Name
                  SiteName          = $L.DisplayName
                  SiteURL           = $L.Name 
                  RetentionTime     = $Rule.RetentionDurationDisplayHint
                  RetentionDuration = $Duration
                  RetentionAction   = $RetentionAction
                  Settings          = $Settings}
               $Report.Add($ReportLine)}                    
          }
}
$Policy = @{e={$_.PolicyName}; l= "Policy";width=30}
$Report | Sort SiteName | Format-Table $Policy, @{e={$_.SiteName};l="Site"}, @{e={$_.RetentionDuration};l="Duration"}, @{e={$_.RetentionAction};l="Action"} -AutoSize

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
