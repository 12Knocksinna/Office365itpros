# ConvertDLtoO365Group.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ConvertDLtoO365Group.PS1
#    .SYNOPSIS
#    Converts a distribution list to an Microsoft 365 Group.
#   
#    .PARAMETER InputGroup
#    The name (or alias) of the input distribution group to process
#   
#    .EXAMPLE
#    .\ConvertDLtoO365Group -InputGroup Managers
#
#  First rudimentary version: 12-May-2015  Tony Redmond
#  Moved to GitHub from the TechnNet Gallery in May 2020
#

param(
    [parameter(Position=1,Mandatory=$true,ValueFromPipeline=$false,HelpMessage="Input the DL to convert")][string]$InputGroup
    )

# First, check to see whether a Group of the same name already exists
$Test = Get-Recipient -Identity $InputGroup -ErrorAction SilentlyContinue
if ($Test -eq $Null) 
   { 
    Write-Host "Can't find that object in Azure Active Directory. Sorry!"
    EXIT
   } 

$Test = Get-UnifiedGroup -Identity $InputGroup -ErrorAction SilentlyContinue

if ($Test -ne $null)
 { 
  Write-Host "A Microsoft 365 Group with the same alias already exists (" $InputGroup "). Exiting..."
  EXIT
}

# Check to see whether the input group is a plain old Exchange DL
$InputDL = Get-DistributionGroup -Identity $InputGroup
$DlType = $InputDL.RecipientTypeDetails

if ($DlType -ne "MailUniversalDistributionGroup")
{ 
  Write-Host $InputGroup " is a" $DlType "group and cannot be converted to a Microsoft 365 Group"
  EXIT
}

# Collect some other information
Write-Host "Collecting some information..."
$DisplayName = $InputDL.DisplayName
[Array]$ManagedBy = $InputDL.ManagedBy
[Array]$Members = Get-DistributionGroupMember -Identity $InputGroup

# Filter so we know the DL members that can be added to the Microsoft 365 Group
[Array]$NewO365GroupMembers = $Members | Where-Object {$_.RecipientTypeDetails –eq “UserMailbox”}

# Create a new alias to give to the DL so that we can reuse its alias for the new Microsoft 365 Group
$O365GroupAlias = $InputDL.Alias
$NewAlias = $O365GroupAlias + "-convert"
$NewDLDisplayName = $InputDL.DisplayName + " (Converted to O365 Group)"

# Figure out the default accepted domain for the tenant so that we can create a new email address for the DL
Foreach ($Domain in (Get-AcceptedDomain)) {
 if ($Domain.Default -eq "True") {
    $DefaultDomain = $Domain.DomainName 
    }
}

$OldDLAddress = $InputDL.PrimarySMTPAddress
$OldDLName = $OldDLAddress.Split("@")[0]
$NewDLAddress = $OldDLName+"-Convert@"+$DefaultDomain

# Tell the admin if there are any potential problems adding recipients to the group
[decimal]$BadType = 0
[Array]$BadMembers = @()
[string]$NewBad
Foreach ($Member in $Members) {
   $MType = (Get-Recipient -Identity $Member.alias -ErrorAction SilentlyContinue).RecipientTypeDetails 
      If ($MType -ne "UserMailbox") {
         $BadType = $BadType + 1 
         $NewBad = $Member.Name + " (" + $Mtype +")"
         $BadMembers += $NewBad
      }
   }

# Phew... Now that we know what we're going to do, let's tell the admin and ask them if it's OK to proceed...
Write-Host " "
Write-Host "Now ready to convert the Distribution Group" $DisplayName "to a new Microsoft 365 Group"
Write-Host "The Alias and email address for the Distribution Group will be switched to the Microsoft 365 Group"
Write-Host "The new alias for the Distribution Group will be" $NewAlias "and its email address will be" $NewDLAddress
Write-Host "The Distribution Group will be renamed to" $NewDLDisplayName "and it will be hidden from the GAL"
if ($BadType -gt 0) {
   Write-Host "Warning!" $BadType "recipients of" $Members.Count "will be dropped from the Microsoft 365 Group because their recipient type is unsupported"
   Write-Host "The recipients that will be dropped are:"
   $BadMembers
   Write-Host " "
    }
 else {
   Write-Host $Members.Count "recipients will be transferred to the new Microsoft 365 Group"
 }
Write-Host $ManagedBy.Count "manager(s) for the $DisplayName group will be transferred"
Write-Host " "

# Simple confirmation to make sure that we are all set to go
$Confirmation = Read-Host "Do you want to proceed with the conversion?"
if ($Confirmation -eq "N") 
  {
    EXIT
  }

Write-Host "Processing..."

# Prepare for the conversion by first switching the alias so we can reuse it for the new Microsoft 365 Group
Set-DistributionGroup $InputGroup -Alias $NewAlias
$InputDL.Alias = $NewAlias

# Check to see whether the DL is closed or has a member restriction. If it is, we offer the user the option to make the new group private
# First, initialize $Answer with 1 because a public group is the default
$Answer = 1
if ($InputDL.MemberJoinRestriction -eq "Closed" -or $InputDL.MemberJoinRestriction -eq "ApprovalRequired") {
   $Message = "The source distribution group is closed. Do you want the Microsoft 365 Group to be private?"
   $Caption = "Select the type of Microsoft 365 group to create";
   $PrivateOption= new-Object System.Management.Automation.Host.ChoiceDescription "&Private Group","Private Group";
   $PublicOption = new-Object System.Management.Automation.Host.ChoiceDescription "&Open (Public)","Open (Public)";
   $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($PrivateOption,$PublicOption);
   $Answer = $host.ui.PromptForChoice($Caption,$Message,$Choices,0)
}

# The alias for the source DL is updated so we can now create a new Microsoft 365 Group based on the old alias
# When Microsoft updates Exchange Online to support creation of private and public groups, this code will be updated
# to use the following code for private groups:
# New-UnifiedGroup -Alias $O365GroupAlias -DisplayName $DisplayName -AccessType Private
# and for public groups
# New-UnifiedGroup -Alias $O365GroupAlias -DisplayName $DisplayName -AccessType Public
# for the moment, a public group is all we can create

Write-Host "Creating new Microsoft 365 Group:" $DisplayName "   This might take a moment..."
If ($Answer -eq 1) {
    # Create a public group
    New-UnifiedGroup -Alias $O365GroupAlias -DisplayName $DisplayName
    }
  else {
    # Create a pruvare group
    New-UnifiedGroup -Alias $O365GroupAlias -DisplayName $DisplayName
}

# Now move all of the settings that we can from the old DL to the new group. We're doing this in a number of commands just
# to break up processing
# Note that we set AutoSubscribeNewMembers to $True to mimic the behaviour of a traditional DL	

Set-UnifiedGroup -Identity $O365GroupAlias -CustomAttribute1 $InputDL.CustomAttribute1 `
    -CustomAttribute2 $InputDL.CustomAttribute2  `
    -CustomAttribute3 $InputDL.CustomAttribute3  `
    -CustomAttribute4 $InputDL.CustomAttribute4  `
    -CustomAttribute5 $InputDL.CustomAttribute5  `
    -CustomAttribute6 $InputDL.CustomAttribute6  `
    -CustomAttribute7 $InputDL.CustomAttribute7 `
    -CustomAttribute8 $InputDL.CustomAttribute8  `
    -CustomAttribute9 $InputDL.CustomAttribute9  `
    -CustomAttribute10 $InputDL.CustomAttribute10  `
    -CustomAttribute11 $InputDL.CustomAttribute11  `
    -CustomAttribute12 $InputDL.CustomAttribute12  `
    -CustomAttribute13 $InputDL.CustomAttribute13  `
    -CustomAttribute14 $InputDL.CustomAttribute14  `
    -CustomAttribute15 $InputDL.CustomAttribute15  

Set-UnifiedGroup -Identity $O365GroupAlias -ExtensionCustomAttribute1 $InputDL.ExtensionCustomAttribute1 `
    -ExtensionCustomAttribute2 $InputDL.ExtensionCustomAttribute2 `
    -ExtensionCustomAttribute3 $InputDL.ExtensionCustomAttribute3 `
    -ExtensionCustomAttribute4 $InputDL.ExtensionCustomAttribute4 `
    -ExtensionCustomAttribute5 $InputDL.ExtensionCustomAttribute5 `
    -RequireSenderAuthenticationEnabled $InputDL.RequireSenderAuthenticationEnabled `
    -RejectMessagesFromSendersOrMembers $InputDL.RejectMessagesFromSendersOrMembers `
    -AcceptMessagesOnlyFromSendersOrMembers $InputDL.AcceptMessagesOnlyFromSendersOrMembers `
    -HiddenFromAddressListsEnabled $InputDL.HiddenFromAddressListsEnabled `
    -AutoSubscribeNewMembers:$True `
    -MailTip $InputDL.MailTip `
    -MailTipTranslations $InputDL.MailTipTranslations

# We add members of the DG as both members and subscribers of the Office 365 group so that they receive contributions to conversations via email
# as well as having access to group resources. Note that there is a bug in Office 365 groups that prevents subscriber records being added, which is
# why the silentlycontinue error handler is place 

Write-Host "Adding members to the new Microsoft 365 Group"
# Using the filtered array created earlier to only add the mailbox members of the DL who can be added
Foreach ($Member in $NewO365GroupMembers) { 
   $MType = (Get-Recipient -Identity $Member.alias -ErrorAction SilentlyContinue).RecipientTypeDetails 
      If ($MType -eq "UserMailbox") {
       Add-UnifiedGroupLinks -Identity $O365GroupAlias -Links $Member.Alias -LinkType "Members"
       Add-UnifiedGroupLinks -Identity $O365GroupAlias -Links $Member.Alias -LinkType "Subscribers" -ErrorAction SilentlyContinue
  }
}

# Add-UnifiedGroupLinks –Identity $O365GroupAlias –LinkType Members –Links $NewO365GroupMembers.Alias 

# Now add all of the managers - they have to be members of the group first, so we make sure by adding them as a member too
Write-Host "Adding managers"
Foreach ($Manager in $ManagedBy) {
   Add-UnifiedGroupLinks -Identity $O365GroupAlias -Links (Get-Recipient $Manager).Alias -LinkType "Members" -ErrorAction SilentlyContinue
   Add-UnifiedGroupLinks -Identity $O365GroupAlias -Links (Get-Recipient $Manager).Alias -LinkType "Owners" 
 }

# To make sure that messages sent to the old DL now go to the new Microsoft 365 Group, we switch the primary email address of the old DL to the new group
Write-Host "Switching email addresses to route messages to new Microsoft 365 Group..."

Set-DistributionGroup -Identity $NewAlias -PrimarySMTPAddress $NewDLAddress 
Set-DistributionGroup -Identity $NewAlias -EmailAddresses $NewDLAddress
Write-Host "Waiting for directory sychronization to complete before switching addresses..."
Sleep -Seconds 8

Set-UnifiedGroup -Identity $O365GroupAlias -EmailAddresses $OldDLAddress

# Now set the old DL to be hidden from the GAL
Set-DistributionGroup -Identity $NewAlias -HiddenFromAddressListsEnabled $True -DisplayName $NewDLDisplayName

# All done - tell the user what we have done
$DisplayName = (Get-UnifiedGroup -Identity $O365GroupAlias).DisplayName
Write-Host "Details of new Microsoft 365 Group alias:" $O365GroupAlias " Display name:" $DisplayName
Write-Host "Email address for the new group is set to: " (Get-UnifiedGroup -Identity $O365GroupAlias).PrimarySmtpAddress
Write-Host " "
Write-Host "Group Members" 
Get-UnifiedGroupLinks -Identity $O365GroupAlias -LinkType Members | Format-Table DisplayName
Write-Host "Group Managers"
Get-UnifiedGroupLinks -Identity $O365GroupAlias -LinkType Owners | Format-Table DisplayName
Write-Host " "

Write-Host "Process complete. New Microsoft 365 Group created:" (Get-UnifiedGroup -Identity $O365GroupAlias).DisplayName
Write-Host "Note that the old distribution group is hidden from GAL and is now called:" (Get-DistributionGroup -Identity $NewAlias).DisplayName

# Done and dusted...
EXIT

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
