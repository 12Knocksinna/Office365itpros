# Report-EntraIDGovernanceCosts
# Read Entra ID audit logs to find records where guest accounts use ID Governance chargeable features and report same.

# V1.0 22-Jul-2025
# GitHub Link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-EntraIDGovernanceCosts.PS1



# Unsupported record types: PlannerTask, Discovery, AzureActiveDirectoryStsLogon
# unsupported operations: Add member to group, MoveToDeletedItems
# Session Id is only available after a successful authentication, so some records will not have a session ID.

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.petri.com. See our post about the Office 365 for IT Pros repository # https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.