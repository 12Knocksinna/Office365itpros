
# Compute-EntraIDPremiumBackupObjects.PS1
# Microsoft Entra Backup and Recovery Premium cost estimator
# Current preview price: $0.025 USD per protected object/month

# V1.0 25-Nov-2026
# GitHub link:

# See https://office365itpros.com/2026/11/26/microsoft-365-backup-entra/ for more information about the
# Microsoft Entra Backup and Recovery Premium license and how to use this script to estimate the cost of protecting your Microsoft Entra objects.

function Get-AppRoleAssignmentCount {
    [CmdletBinding()]
    param ()

    [long]$AssignmentCount = 0
    [long]$ServicePrincipalCount = 0

   #$ServicePrincipals = Get-MgServicePrincipal -All -Property Id, displayName

    ForEach ($ServicePrincipal in $ServicePrincipals) {
        $ServicePrincipalCount++
        Write-Progress -Activity "Counting app role assignments" -Status ("Processing {0}" -f $ServicePrincipal.DisplayName) `
            -CurrentOperation ("{0:N0} service principals processed; {1:N0} assignments found" -f $ServicePrincipalCount, $AssignmentCount)
        Try {
            $Assignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $ServicePrincipal.Id -All -Property Id -ErrorAction Stop
            $AssignmentCount += $Assignments.Count
        } Catch {
            Write-Warning ("Failed to retrieve app role assignments for {0} ({1}): {2}" -f  $ServicePrincipal.DisplayName, $ServicePrincipal.Id, $_.Exception.Message)
            return
        }
    }
    Write-Progress -Activity "Counting app role assignments" -Completed

    $AssignmentCount
}

function Get-EntraObjectCount {
    param (
        [string]$Name
    )

    $Headers = @{ConsistencyLevel = "eventual"}
    Write-Host ("Counting {0}..." -f $Name)

    # Call the appropriate Microsoft Graph API endpoint based on the object type
    Switch ($Name) {
        "Users" { 
            $Count = Get-MgUserCount -ConsistencyLevel eventual
        }
        "Groups" { 
            $Uri = "https://graph.microsoft.com/v1.0/groups?`$count=true"
            $Count = (Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers $Headers).'@odata.count'
            # Alternative method using the built-in cmdlet (uncomment if needed)
            # $Count = (Get-MgGroup -All -Property Id).count
        }
        "Applications" { 
            $Count = (Get-MgApplication -All -Property Id).count
        }
        "Service Principals" { 
            $Script:ServicePrincipals = Get-MgServicePrincipal -All 
            $Count = $ServicePrincipals.Count
        }
        "Conditional Access Policies" { 
            $Count = (Get-MgIdentityConditionalAccessPolicy -All -Property Id).count
        }
        "Named Locations" { 
            $Count = (Get-MgIdentityConditionalAccessNamedLocation -All -Property Id).count
        }
        "Authentication Strength Policies" { 
            $Count = (Get-MgPolicyAuthenticationStrengthPolicy -All -Property Id).count
        }
        "Delegated Permission Grants" {
            $Uri = "https://graph.microsoft.com/v1.0/oauth2PermissionGrants"
            $Count = 0
            Do {
                $Result = Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers $Headers
                $Count += $Result.value.Count
                $Uri = $Result.'@odata.nextLink'
            } While ($Uri)            
        }
        "Agent Identities" {
            $Uri = "https://graph.microsoft.com/beta/servicePrincipals/microsoft.graph.agentIdentity"
            $Count = 0
            Do {
                $Result = Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers $Headers
                $Count += $Result.value.Count
                $Uri = $Result.'@odata.nextLink'
            } While ($Uri)
        }
        default { 
            Write-Host ("Unknown object type: {0}" -f $Name)
            return 
        }
    }

    $Report.Add([PSCustomObject]@{
        ObjectType  = $Name
        ObjectCount = [int64]$Count
        MonthlyCost = [math]::Round(([int64]$Count * $PricePerObject), 2)
    })
}

Connect-MgGraph -Scopes @("User.ReadBasic.All","Group.Read.All","Application.Read.All","Policy.Read.All","Policy.Read.ConditionalAccess","DelegatedPermissionGrant.Read.All",
    "AppRoleAssignment.ReadWrite.All") -NoWelcome

$PricePerObject = 0.025
$Report = [System.Collections.Generic.List[Object]]::new()

# Collect data for the various object types that can be protected by Microsoft Entra Backup and Recovery Premium. 
# Each object type is counted, and the monthly cost is calculated based on the current preview price.
Get-EntraObjectCount -Name "Users" 
Get-EntraObjectCount -Name "Groups"
Get-EntraObjectCount -Name "Applications"
Get-EntraObjectCount -Name "Service Principals"
Get-EntraObjectCount -Name "Conditional Access Policies" 
Get-EntraObjectCount -Name "Named Locations" 
Get-EntraObjectCount -Name "Authentication Strength Policies" 
Get-EntraObjectCount -Name "Delegated Permission Grants" 
Get-EntraObjectCount -Name "Agent Identities"

# App Role Assignments are not directly retrievable via a single Microsoft Graph API call, so we need to iterate through 
# allrvice principals to count their app role assignments. This can be time-consuming for tenants with many service principals, but it is necessary to get an accurate count of app role assignments.
$AppRoleAssignmentCount = Get-AppRoleAssignmentCount

$Report.Add([PSCustomObject]@{
    ObjectType  = "App Role Assignments"
    ObjectCount = $AppRoleAssignmentCount
    MonthlyCost = [math]::Round(($AppRoleAssignmentCount * $PricePerObject), 2)})

    
$TotalObjects = ($Report | Measure-Object ObjectCount -Sum).Sum
$MonthlyCost = [math]::Round(($TotalObjects * $PricePerObject), 2)
$AnnualCost = [math]::Round(($MonthlyCost * 12), 2)

$Report | Format-Table ObjectType, ObjectCount, MonthlyCost -AutoSize

Write-Host
Write-Host ("Total protected objects: {0:N0}" -f $TotalObjects)
Write-Host ("Estimated monthly cost: `${0:N2} USD" -f $MonthlyCost)
Write-Host ("Estimated annual cost:  `${0:N2} USD" -f $AnnualCost)

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 