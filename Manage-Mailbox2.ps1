<#
.SYNOPSIS
    Efficiently lists Exchange Online distribution groups you're a member of and allows adding/removing members.

.DESCRIPTION
    Uses Get-Recipient to efficiently find groups where you're already a member,
    then filters for groups containing arrowEEdot in the DisplayName.

.PARAMETER UserPrincipalName
    The azure user@domain.name.whatever of the user running the script.

.PARAMETER DryRun
    If specified, simulates actions without making changes.

.EXAMPLE
    .\Manage-Mailbox.ps1 -UserPrincipalName "j.bloggs@domain.com"
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$UserPrincipalName,

    [switch]$DryRun
)

# Suppress validation warnings globally
$WarningPreference = 'SilentlyContinue'

function Ensure-ExchangeModule {
    $moduleName = "ExchangeOnlineManagement"
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Host "Installing Exchange Online module..." -ForegroundColor Yellow
        try {
            Install-Module $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        } catch {
            Write-Error "Module install failed $_"
            exit
        }
    }
    Import-Module $moduleName -ErrorAction Stop
}

function Connect-Exchange {
    Write-Verbose "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ErrorAction Stop
}

function Get-ManagedGroupsByMembership {
    Write-Verbose "Starting optimized scan for mail-enabled security groups..."

    $matchingGroups = @()

    try {
        # Method 1: Get user's group memberships directly (most efficient)
        Write-Verbose "Getting group memberships for $UserPrincipalName..."
        
        # Get the user object first
        $user = Get-Recipient -Identity $UserPrincipalName -ErrorAction Stop
        Write-Verbose "Found user: $($user.DisplayName)"

        # Get groups where this user is a member using reverse lookup
        $userGroups = Get-Recipient -Filter "Members -eq '$($user.DistinguishedName)'" -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited
        
        Write-Verbose "Found $($userGroups.Count) groups where user is a member"

        # Filter for groups starting with arrowEEdot and collect patterns for Owners groups
        $ownerGroupBases = @()
        
        foreach ($group in $userGroups) {
            $emailMatch = $group.DisplayName -like ">EE.*"
            
            if ($emailMatch) {
                Write-Verbose "Matched group: $($group.DisplayName)"
                Write-Host "Found manageable group: $($group.DisplayName)" -ForegroundColor Green
                $matchingGroups += $group
                
                # Check if this is an Owners group and create patterns for related groups
                if ($group.DisplayName -like ">EE.*.Owners.*") {
                    # Replace .Owners. with SendAs and SendOnBehalf to create search patterns
                    $sendAsPattern = $group.DisplayName -replace '\.Owners\.', '.SendAs.'
                    $sendOnBehalfPattern = $group.DisplayName -replace '\.Owners\.', '.SendOnBehalf.'
                    $ownerGroupBases += @{
                        'SendAs' = $sendAsPattern
                        'SendOnBehalf' = $sendOnBehalfPattern
                    }
                    Write-Verbose "Found owner group, will search for: $sendAsPattern and $sendOnBehalfPattern"
                }
            } else {
                Write-Verbose "Skipping group $($group.DisplayName) - doesn't match EE filter"
            }
        }

        # Now find the corresponding SendAs and SendOnBehalf groups for each owner group
        if ($ownerGroupBases.Count -gt 0) {
            Write-Verbose "Looking for related SendAs and SendOnBehalf groups for $($ownerGroupBases.Count) owner groups..."
            
            foreach ($groupPatterns in $ownerGroupBases) {
                Write-Verbose "Searching for related groups: $($groupPatterns.SendAs) and $($groupPatterns.SendOnBehalf)"
                
                # Search for SendAs group
                try {
                    $sendAsGroup = Get-DistributionGroup -Identity $groupPatterns.SendAs -RecipientTypeDetails MailUniversalSecurityGroup -ErrorAction SilentlyContinue
                    if ($sendAsGroup) {
                        Write-Host "Found related SendAs group: $($sendAsGroup.DisplayName)" -ForegroundColor Green
                        $matchingGroups += $sendAsGroup
                    } else {
                        Write-Verbose "SendAs group $($groupPatterns.SendAs) not found"
                    }
                } catch {
                    Write-Verbose "SendAs group $($groupPatterns.SendAs) not found or not accessible: $_"
                }
                
                # Search for SendOnBehalf group
                try {
                    $sendOnBehalfGroup = Get-DistributionGroup -Identity $groupPatterns.SendOnBehalf -RecipientTypeDetails MailUniversalSecurityGroup -ErrorAction SilentlyContinue
                    if ($sendOnBehalfGroup) {
                        Write-Host "Found related SendOnBehalf group: $($sendOnBehalfGroup.DisplayName)" -ForegroundColor Green
                        $matchingGroups += $sendOnBehalfGroup
                    } else {
                        Write-Verbose "SendOnBehalf group $($groupPatterns.SendOnBehalf) not found"
                    }
                } catch {
                    Write-Verbose "SendOnBehalf group $($groupPatterns.SendOnBehalf) not found or not accessible: $_"
                }
            }
        }

    } catch {
        Write-Warning "Direct group membership lookup failed: $_"
        Write-Verbose "Falling back to alternative method..."
        
        # Fallback Method: Use Graph API approach via Get-DistributionGroup
        try {
            $filteredGroups = @()
            
            # Search for groups with arrowEEdot in name
            $EEGroups = Get-DistributionGroup -Filter "DisplayName -like '>EE.*'" -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited -ErrorAction SilentlyContinue
            $filteredGroups += $EEGroups
            
            # Remove duplicates
            $filteredGroups = $filteredGroups | Sort-Object Identity -Unique
            
            Write-Verbose "Fallback method found $($filteredGroups.Count) potential groups to check"

            $ownerGroupBases = @()
            
            # Check membership only for these filtered groups
            foreach ($group in $filteredGroups) {
                try {
                    Write-Verbose "Checking membership in $($group.DisplayName)..."
                    $members = Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop
                    
                    $match = $members | Where-Object { $_.PrimarySmtpAddress -eq $UserPrincipalName }
                    
                    if ($match) {
                        Write-Verbose "User $UserPrincipalName is a member of $($group.DisplayName)"
                        Write-Host "Found manageable group: $($group.DisplayName)" -ForegroundColor Green
                        $matchingGroups += $group
                        
                        # Check if this is an Owners group and create patterns for related groups
                        if ($group.DisplayName -like ">EE.*.Owners.*") {
                            # Replace .Owners. with SendAs and SendOnBehalf to create search patterns
                            $sendAsPattern = $group.DisplayName -replace '\.Owners\.', '.SendAs.'
                            $sendOnBehalfPattern = $group.DisplayName -replace '\.Owners\.', '.SendOnBehalf.'
                            $ownerGroupBases += @{
                                'SendAs' = $sendAsPattern
                                'SendOnBehalf' = $sendOnBehalfPattern
                            }
                            Write-Verbose "Found owner group, will search for: $sendAsPattern and $sendOnBehalfPattern"
                        }
                    }
                } catch {
                    Write-Verbose "Skipping group $($group.DisplayName) due to error: $_"
                }
            }
            
            # Find related SendAs and SendOnBehalf groups in fallback mode
            if ($ownerGroupBases.Count -gt 0) {
                foreach ($groupPatterns in $ownerGroupBases) {
                    # Look for corresponding groups in the already retrieved list
                    $relatedGroups = $filteredGroups | Where-Object { 
                        $_.DisplayName -eq $groupPatterns.SendAs -or $_.DisplayName -eq $groupPatterns.SendOnBehalf
                    }
                    
                    foreach ($relatedGroup in $relatedGroups) {
                        Write-Host "Found related group: $($relatedGroup.DisplayName)" -ForegroundColor Green
                        $matchingGroups += $relatedGroup
                    }
                }
            }
            
        } catch {
            Write-Error "Both direct lookup and fallback method failed: $_"
            return @()
        }
    }

    # Remove any duplicates that might have slipped through
    $matchingGroups = $matchingGroups | Sort-Object Identity -Unique
    
    Write-Verbose "Finished scanning. Found $($matchingGroups.Count) manageable groups."
    return $matchingGroups
}

function Show-GroupMembers($group) {
    try {
        Write-Host "`nCurrent members of $($group.DisplayName):" -ForegroundColor Cyan
        $members = Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited
        
        if ($members.Count -eq 0) {
            Write-Host "  No members found" -ForegroundColor Gray
        } else {
            $members | ForEach-Object { 
                Write-Host "  - $($_.DisplayName) ($($_.PrimarySmtpAddress))" -ForegroundColor White 
            }
        }
        Write-Host "  Total: $($members.Count) members" -ForegroundColor Yellow
    } catch {
        Write-Warning "Could not retrieve members for $($group.DisplayName): $_"
    }
}

function Search-Users($searchTerm) {
    try {
        Write-Verbose "Searching for users matching '$searchTerm'..."
        $users = Get-Recipient -Filter "DisplayName -like '*$searchTerm*' -or PrimarySmtpAddress -like '*$searchTerm*'" -RecipientTypeDetails UserMailbox -ResultSize 10
        return $users
    } catch {
        Write-Warning "User search failed: $_"
        return @()
    }
}

function Prompt-Action {
    Write-Host "=========================================================================" -ForegroundColor Magenta
    Write-Host "Choose an action:" -ForegroundColor Cyan
    Write-Host "[1] Add member to a group"
    Write-Host "[2] Remove member from a group"
    Write-Host "[3] View group members"
    Write-Host "[4] Refresh group list"
    Write-Host "[Q] Quit"
    return Read-Host "Enter your choice"
}

function Prompt-GroupSelection($groups) {
    Write-Host "`nGroups you can manage:" -ForegroundColor Green
    for ($i = 0; $i -lt $groups.Count; $i++) {
        Write-Host "  $($i+1). $($groups[$i].DisplayName)" -ForegroundColor White
    }
    
    do {
        $selection = Read-Host "`nEnter group number (1-$($groups.Count)) or 'c' to cancel"
        if ($selection -eq 'c') { return $null }
        
        if ([int]$selection -ge 1 -and [int]$selection -le $groups.Count) {
            return $groups[$selection - 1]
        } else {
            Write-Warning "Invalid selection. Please enter a number between 1 and $($groups.Count)."
        }
    } while ($true)
}

function Prompt-UserEmail {
    $email = Read-Host "Enter the email address of the user, or search term to find users"
    
    # If it looks like an email, use it directly
    if ($email -match '^[^@]+@[^@]+\.[^@]+$') {
        return $email
    }
    
    # Otherwise, search for users (still buggy)
    $users = Search-Users $email
    if ($users.Count -eq 0) {
        Write-Warning "No users found matching '$email'. Please try a different search term."
        return $null
    }
    
    if ($users.Count -eq 1) {
        $user = $users[0]
        Write-Host "Found user: $($user.DisplayName) ($($user.PrimarySmtpAddress))" -ForegroundColor Green
        $confirm = Read-Host "Use this user? (y/n)"
        if ($confirm -eq 'y') {
            return $user.PrimarySmtpAddress
        }
        return $null
    }
    
    # Multiple users found
    Write-Host "`nFound multiple users:" -ForegroundColor Green
    for ($i = 0; $i -lt $users.Count; $i++) {
        Write-Host "  $($i+1). $($users[$i].DisplayName) ($($users[$i].PrimarySmtpAddress))" -ForegroundColor White
    }
    
    do {
        $selection = Read-Host "`nSelect user (1-$($users.Count)) or 'c' to cancel"
        if ($selection -eq 'c') { return $null }
        
        if ([int]$selection -ge 1 -and [int]$selection -le $users.Count) {
            return $users[$selection - 1].PrimarySmtpAddress
        } else {
            Write-Warning "Invalid selection."
        }
    } while ($true)
}

function Add-Member($group, $email) {
    if ($DryRun) {
        Write-Host "[DryRun] Would add $email to $($group.DisplayName)" -ForegroundColor Yellow
    } else {
        try {
            Write-Host "Adding $email to $($group.DisplayName)..." -ForegroundColor Yellow
            Add-DistributionGroupMember -Identity $group.Identity -Member $email -ErrorAction Stop
            Write-Host "Successfully added $email to $($group.DisplayName)" -ForegroundColor Green
        } catch {
            Write-Error "Failed to add member: $_"
        }
    }
}

function Remove-Member($group, $email) {
    if ($DryRun) {
        Write-Host "[DryRun] Would remove $email from $($group.DisplayName)" -ForegroundColor Yellow
    } else {
        try {
            Write-Host "Removing $email from $($group.DisplayName)..." -ForegroundColor Yellow
            Remove-DistributionGroupMember -Identity $group.Identity -Member $email -Confirm:$false -ErrorAction Stop
            Write-Host "Successfully removed $email from $($group.DisplayName)" -ForegroundColor Green
        } catch {
            Write-Error "Failed to remove member: $_"
        }
    }
}

# Main Execution
Write-Host "Exchange Online Group Management Tool" -ForegroundColor Magenta
Write-Host "=======================================================" -ForegroundColor Magenta

Ensure-ExchangeModule
Connect-Exchange

Write-Host "Scanning for manageable groups for: $UserPrincipalName" -ForegroundColor Cyan

$managedGroups = Get-ManagedGroupsByMembership

if (-not $managedGroups -or $managedGroups.Count -eq 0) {
    Write-Warning "No matching groups found where you're a member."
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

Write-Host "`nFound $($managedGroups.Count) manageable group(s)" -ForegroundColor Green

do {
    $choice = Prompt-Action
    switch ($choice) {
        "1" {
            $group = Prompt-GroupSelection $managedGroups
            if ($group) {
                $email = Prompt-UserEmail
                if ($email) {
                    Add-Member $group $email
                }
            }
        }
        "2" {
            $group = Prompt-GroupSelection $managedGroups
            if ($group) {
                Show-GroupMembers $group
                $email = Prompt-UserEmail
                if ($email) {
                    Remove-Member $group $email
                }
            }
        }
        "3" {
            $group = Prompt-GroupSelection $managedGroups
            if ($group) {
                Show-GroupMembers $group
            }
        }
        "4" {
            Write-Host "Refreshing group list..." -ForegroundColor Yellow
            $managedGroups = Get-ManagedGroupsByMembership
            Write-Host "Found $($managedGroups.Count) manageable group(s)" -ForegroundColor Green
        }
        { $_ -in @("Q", "q") } {
            Write-Host "Exiting..." -ForegroundColor Cyan
        }
        default {
            Write-Warning "Invalid choice. Please select 1, 2, 3, 4, or Q."
        }
    }
} while ($choice -notin @("Q", "q"))

Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Done, Bye!" -ForegroundColor Green
