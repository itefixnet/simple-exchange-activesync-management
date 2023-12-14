#
# adeas.ps1 - A simple Exchange ActiveSync management solution
#
# v1.0 - Initial version, sep 2011, itefix.net
# v2.0 - Add option to remove stale partnerships, oct 2011, itefix.net
#
 
# This script implements an Exchange ActiveSync management solution
# based on AD-groups and ActiveSync mailbox policies. Each AD-group corresponds
# to an Activesync level implemented as an ActiveSysnc mailbox policy
#
# Following operations are available:
#
#   - Activate user(s) 
#   - Deactivate user(s)
#   - Turn off ActiveSync for the whole organization
#   - Generate usage reports
#   - Remove stale device partnerships
#  
 
param
(
    [switch]$activate,
    [switch]$deactivate,
    [int]$staledays,
    [switch]$stats,
    [switch]$activesyncoff,
    [string]$user,
    [int]$level,
    [switch]$all,
    [switch]$help,
    [switch]$verbose
)
 
Import-Module activedirectory
#Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
 
##### CUSTOMIZE START! #######
# AD-groups and ActiveSync policies
#
 
# Example 1: An one-level organization : Every ActiveSync user is
# a member of an AD-group and uses the default Activesync mailbox
# policy
 
#$aDGroups =
#@(
#    "**UNUSED**",
#    "AD_Activesync_Users",
#)
#
#$activeSyncPolicies =
#@(
#    "**UNUSED**",
#    "Default",
#)
 
# Example 2: A three-level role-based organization:
# 1-Standard, 2-Managers, 3-Sales
#
 
#$aDGroups =
#@(
#    "**UNUSED**",
#    "AD_Activesync_Standard",
#    "AD_Activesync_Managers",
#    "AD_Activesync_Sales",
#)
#
#$activeSyncPolicies =
#@(
#    "**UNUSED**",
#    "Default",
#    "Managers",
#    "Sales"
#)
 
# Example 3: A three-level device capability-based
# organization:
# 1-LowSecurity, 2-MediumSecurity, 3-HighSecurity
#
 
#$aDGroups =
#@(
#    "**UNUSED**",
#    "AD_Activesync_LowSecurity",
#    "AD_Activesync_MediumSecurity",
#    "AD_Activesync_HighSecurity",
#)
#
#$activeSyncPolicies =
#@(
#    "**UNUSED**",
#    "LowSecurity",
#    "MediumSecurity",
#    "HighSecurity"
#)
 
$aDGroups =
@(
 
)
 
$activeSyncPolicies =
@(
 
)
 
##### CUSTOMIZE END! #######
 
$stats_user_format = "{0,-8} {1,-8} {2,-20} {3,-36} {4,-20}"
 
#####
# ActivateUser: Activates a user for an ActiveSync policy
 
function ActivateUser
{
    param (
        [string]$puser,
        [string]$ppolicy
    )
 
    if ($verbose) { write-host "Activating user $puser for policy $ppolicy..." }
 
    try
    {
        Set-CASMailbox $puser -ActiveSyncEnabled:$true -ActiveSyncMailboxPolicy:$ppolicy
    }
 
    catch
    {
        write-host "User activation problem: $_"
    }
}
 
#####
# DeactivateUser: Dectivates a user
 
function DeactivateUser
{
    param (
        [string]$puser
    )
 
    if ($verbose) { write-host "Dectivating user $puser for ActiveSync ..." }
 
    try
    {
        Set-CASMailbox $puser -ActiveSyncEnabled:$false
    }
 
    catch
    {
        write-host "User dectivation problem: $_"
    }
}
 
#####
# StatsUser: Generates device/user list for a user
 
function StatsUser
{
    param (
        [string]$suser,
        [int]$slevel
    )
 
    if (!$suser -or $suser -eq "") { return }
 
    if ($verbose) { write-host "Processing $suser" }
 
    try
    {
        $stats = Get-ActiveSyncDeviceStatistics -Mailbox $suser
 
        foreach ($stat in $stats)
        {       
            if ($stat.LastSuccessSync -ne $null)
            {
                $stats_user_format -f $suser, $slevel, $stat.DeviceType, $stat.DeviceID, $stat.LastSuccessSync
            }
        }
    }
 
    catch
    {
        write-host "ActiveSync stats problem: $_" 
    }  
}
 
#####
# RemoveStalePartnership: Remove stale Activesync partnerships
# for a user.
 
function RemoveStalePartnership
{
    param (
        [string]$suser,
        [int]$sdays
    )
 
    if (!$suser -or $suser -eq "") { return }
 
    if ($verbose) { write-host "Checking stale Activesync partnerships for $suser ($sdays days)" }
 
    try
    {
        $partnerships = Get-ActiveSyncDeviceStatistics -Mailbox $suser
 
        foreach ($partner in $partnerships)
        {      
            if ($partner.Identity -and $partner.LastSuccessSync -le (Get-Date).AddDays($sdays))
            {
                if ($verbose) { write-host "Removing partnership " $partner.Identity }
                Remove-ActiveSyncDevice -Identity $partner.Identity -confirm:$false
            }
        }
    }
 
    catch
    {
        write-host "ActiveSync remove stale activesync partnership problem: $_" 
    }  
}
 
#####
# GroupMembers: return members of a group for an ActiveSync policy
 
function GroupMembers ($group)
{
 
    # Get group members
    try
    {
        $groupMembers = Get-ADGroupMember $group
    }
 
    catch
    {
        write-host "AD Group problem: $_"
        exit 1
    }
 
    $groupMembers
}
 
#####
# PrintHelp: Viser hjelp
 
function PrintHelp
{
    write-host @"
 
Usage:
 
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
 
"@
}
 
#
# Main program
#
 
if ($help)
{
    PrintHelp
 
# (De)Activate user(s)
#   
} elseif ($activate -or $deactivate) {
 
    # Activate a user
    if ($activate -and $user -and $level)
    {
        # make sure that user is member of only one group 
        for ($i=1; $i -le $adGroups.length - 1; $i++)
        {
            if ($level -eq $i)
            {
                if ($verbose) { write-host "Adding user $user to the group " $adGroups[$i] }
 
                try { Add-ADGroupMember -identity $adGroups[$i] -members $user }
                catch { write-host $_ }
 
            } else {
                if ($verbose) { write-host "Removing user $user from the group " $adGroups[$i] }           
 
                try { Remove-ADGroupMember -identity $adGroups[$i] -members $user -confirm:$false }
                catch { write-host $_ }
            }
         }
 
         ActivateUser $user $activeSyncPolicies[$level]
 
    # Deactivate a user
    } elseif ($deactivate -and $user) {
 
        # Remove user from all groups 
        for ($i=1; $i -le $adGroups.length - 1; $i++)
        {
            if ($verbose) { write-host "Removing user $user from the group " $adGroups[$i] }           
 
            try { Remove-ADGroupMember -identity $adGroups[$i] -members $user -confirm:$false }
            catch { write-host $_ }
         }
 
         DeactivateUser $user  
 
    # (De)Activate all/specific level users
    } elseif ($all -or ($level -and !$user)) {
 
        for ($i=1; $i -le $activeSyncPolicies.length - 1; $i++)
        {
            # Process only specified level if a level is specified
            if ($level -and ($level -ne $i)) { continue }
 
            if ($verbose) {
                if ($activate) {write-host "Activating level $i users ..." }
                if ($deactivate) {write-host "Deactivating level $i users ..." }
            }
 
            # Process group members  
            foreach ($groupMember in GroupMembers($adGroups[$i]))
            { 
                if ($activate)
                {
                    ActivateUser $groupMember.Name $activeSyncPolicies[$i]
 
                } elseif ($deactivate) {
                    DeactivateUser $groupMember.Name
 
                }
            }  
        }
     } else {
 
        PrintHelp
     }
 
# Produce stats
} elseif ($stats) {
 
    $stats_user_format -f "BID", "Level", "DeviceType", "DeviceID", "LastSuccessSync"
    $stats_user_format -f "---", "-----", "----------", "--------", "---------------"
 
    # Stats for a specific user
    if ($user) {
 
        $wlevel = 0
 
        # try to find level of the user
        for ($i=1; $i -le $activeSyncPolicies.length - 1; $i++)
        {   
            foreach ($groupMember in GroupMembers($adGroups[$i]))
            {                 
                if ($groupMember.Name -eq $user)
                {
                    $wlevel = $i
                    break
                }
            }  
        }
 
        StatsUser $user $wlevel
 
    # Stats for a level
    } elseif ($level) {
 
        # Process group members for the level 
         foreach ($groupMember in GroupMembers($adGroups[$level]))
         {
            StatsUser $groupMember.Name $level
         }
 
    # Stats for all/level users      
    } elseif ($all) {
 
        for ($i=1; $i -le $activeSyncPolicies.length - 1; $i++)
        {                              
            # Process group members  
            foreach ($groupMember in GroupMembers($adGroups[$i]))
            {
                StatsUser $groupMember.Name $i
            }
        }
     } else {
        PrintHelp
     }
 
# Turn ActiveSync off for all mailboxes
} elseif ($activesyncoff) {
 
    $res = read-host "Do you really want to DISABLE ActiveSync for ALL users (y/n)?"
 
    if ($res -eq "Y")
    {
        get-Mailbox -ResultSize:unlimited | set-CASMailbox -ActiveSyncEnabled:$false -WarningAction SilentContinue
        #write-host "Y is selected"
    }
 
# Remove stale ActiveSync device partnerships
} elseif ($staledays) {
 
    for ($i=1; $i -le $activeSyncPolicies.length - 1; $i++)
    {                              
        # Process group members  
        foreach ($groupMember in GroupMembers($adGroups[$i]))
        {
            # add - before days for backward operation
            RemoveStalePartnership $groupMember.Name "-$staledays"
        }
     }
 
# Improper option combination. Get some help !! 
} else {
    PrintHelp
}