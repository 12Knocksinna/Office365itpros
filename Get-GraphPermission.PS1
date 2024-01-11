<#
.SYNOPSIS
This function retrieves the possible Microsoft Graph permissions used in a PowerShell script.
.DESCRIPTION
The Get-GraphScriptPermission function takes a scriptblock as input and parses it to extract a list of command elements and their associated parameters. It then filters the list to include only the commands sourced from Microsoft.Graph. For each Microsoft.Graph command, it calls the Find-MgGraphCommand function to retrieve the permissions associated with that command. The function returns an object containing the command, its source, verb, noun, and the list of permissions with their names and whether they require admin privileges.
.PARAMETER Script
The scriptblock to be analyzed for Microsoft Graph commands and their permissions.
.EXAMPLE
$script = {
    Get-MgUser -Filter "Department eq 'Sales'"
    New-MgGroup -DisplayName 'Marketing Group' -Description 'Group for marketing team'
    Get-MGApplication -Filter "DisplayName eq 'My Application'"
}
Get-GraphScriptPermission -Script $script
This example retrieves the Microsoft Graph permissions used in the provided scriptblock.
.OUTPUTS
The function returns an object with the following properties for each Microsoft.Graph command:
- Cmdlet: The name of the command.
- Source: The source of the command.
- Verb: The verb of the command.
- Type: The noun of the command.
- Scopes: A list of scopes associated with the command. Each object has the following properties:
    - Name: The name of the permission.
    - IsAdmin: Indicates whether the permission requires admin privileges.

# https://github.com/12Knocksinna/Office365itpros/blob/master/Get-GraphPermission.PS1
#>
function Get-GraphScriptPermission {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [scriptblock] $Script
    )
    
    begin {

        $ast = [System.Management.Automation.Language.Parser]::ParseInput($Script.ToString(), [ref]$null, [ref]$null)
        [array]$commandElementList = $null
        # Extract a list of command elements and their associated parameters from the AST
        [array]$CommandElementList = $ast.FindAll({$args[0].GetType().Name -like 'CommandAst'}, $true) | ForEach-Object {
            [pscustomobject]@{
                Cmdlet                  = $Cmdlet = $_.CommandElements[0].Value
                Source                  = (Get-Command -Name $Cmdlet).Source
                Verb                    = (Get-Command -Name $Cmdlet).Verb
                Type                    = (Get-Command -Name $Cmdlet).Noun
                AllPrivileges   = $null
            }
        }
    }
    
    process {

        $GraphScopeReport = [System.Collections.Generic.List[Object]]::new()
        [array]$CommandList = $CommandElementList | Where-Object Source -like 'Microsoft.Graph*'
        ForEach ($GraphCommand in $CommandList) { 
            [array]$ScopeOutput = $null
            [array]$Scopes =  (Find-MgGraphCommand -Command $GraphCommand.Cmdlet | `
                	Select-Object -ExpandProperty Permissions | Sort-Object Name, isAdmin -Unique)
            ForEach ($Scope in $Scopes)  {
                $ScopeInfo = ("{0} (admin: {1})" -f $Scope.Name, $Scope.isAdmin)
                [array]$ScopeOutput += $ScopeInfo
            }
            [string]$ScopeOutput = $ScopeOutput -Join ", "
            $DataLine = [PSCustomObject][Ordered]@{
                Cmdlet          = $GraphCommand.cmdlet
                Source          = $GraphCommand.Source
                Verb            = $GraphCommand.Verb
                Type            = $GraphCommand.Type
                Scopes          = $ScopeOutput 
            }
            $GraphScopeReport.Add($DataLine) 
        }
        $GraphScopeReport | Sort-Object Cmdlet -Unique
    }
    
    end {
    }
}

# Original version from https://gist.github.com/HCRitter/d7017ce1eeb66689b14a108423d383ab. See
# https://www.linkedin.com/pulse/get-graphscriptpermission-christian-ritter-gsaie/ for more information.
# Amended in several places to make the list of scopes a string rather than an array