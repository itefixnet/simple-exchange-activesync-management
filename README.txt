Adeas.ps1 is a Powershell script which implements a simple Exchange ActiveSync management solution based on AD-groups and ActiveSync mailbox policies. Each AD-group corresponds to a level implemented by an ActiveSysnc mailbox policy.

Following operations are available:

- Activate user(s)
- Deactivate user(s)
- Disable ActiveSync for all users
- Generate usage reports
- Remove stale device partnerships

Usage info from the script:
 
    adeas.ps1 [-activate] [-deactivate] [-stats] [-activesyncoff] [-staledays <days>] [-all] [-user <user> [-level <level>] ] [-help] [-verbose]
 
    This script implements a simple Activesync management solution based on AD-groups and Activesync mailbox policies.
 
    -activate: Enables ActiveSync for user(s).     
 
       Examples:
          adeas.ps1 -activate -all
             Enables ActiveSync and activates corresponding mailbox policies for all users.
 
          adeas.ps1 -activate -level 2
             Enables ActiveSync and activates the corresponding mailbox policy for level 2 users.
 
          adeas.ps1 -activate -user myuser -level 1 -verbose
             Enables ActiveSync, adds 'myuser' to the level 1 group and activates the corresponding
             the mailbox policy as well. Produce progress messages during operations.
 
    -deactivate: Disables Activesync for user(s).
 
        Examples:
           adeas.ps1 -deactivate -all
           adeas.ps1 -deactivate -level 1
           adeas.ps1 -deactivate -user myuser
 
    -activesyncoff: Turns off ActiveSync for all mailboxes in the organization
 
        Example:
           adeas.ps1 -activesyncoff
 
    -stats: Generate reports/statistics
 
        Examples:
           adeas.ps1 -stats -all
           adeas.ps1 -stats -level 2
           adeas.ps1 -stats -user myuser
 
    -staledays: Removes ActiveSync devices which have not been synchronized within a specified number of days.
 
        Example:
           adeas.ps1 -staledays 60 -verbose
              Removes devices which have not been synchronized during last 60 days.