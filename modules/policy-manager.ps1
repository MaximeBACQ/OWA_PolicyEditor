# ============================================================
# modules/policy-manager.ps1 — OWA Policy Manager
# ============================================================
# Main hub for listing, creating, assigning, and deleting
# OWA Mailbox Policies in Exchange Online.
#
# For each policy in the list we display:
#   - Policy name
#   - Number of assigned users (from Get-CasMailbox)
#   - Whether it is the default policy
#
# Actions available per policy:
#   View/Edit  — opens the viewer, then editor on request
#   Assign     — searchable list of mailboxes to assign policy
#   Reset      — remove custom policy assignment (→ default)
#   Delete     — non-default policies only, with confirmation
#   Create new — prompt for name, create, open in editor
# ============================================================

# ── Main policy manager screen ────────────────────────────────
function Show-PolicyManager {
    if (-not (Assert-Connected)) { return }

    while ($true) {
        # Fetch policy list with user counts
        Clear-Host
        Write-AppHeader -Subtitle "Policy Manager"
        Write-Ansi "${Gray}  Loading policies...${Reset}"

        try {
            $policyData = Get-AllPoliciesWithCounts
        } catch {
            Write-Ansi "${Red}  Error fetching policies: $($_.Exception.Message.Split([Environment]::NewLine)[0])${Reset}"
            Wait-KeyPress
            return
        }

        if ($policyData.Count -eq 0) {
            Write-Ansi "${Yellow}  No OWA Mailbox Policies found.${Reset}"
            Wait-KeyPress
            return
        }

        # Build menu items
        $menuItems = @()
        foreach ($p in $policyData) {
            $defaultTag  = if ($p.IsDefault) { "${Blue}[DEFAULT]${Reset}" } else { "         " }
            $userCount   = "${Cyan}$($p.UserCount.ToString().PadLeft(4))${Reset} users"
            $label       = "$defaultTag  $($p.Name.PadRight(40))  $userCount"
            $menuItems  += @{ Label = $label; Value = $p }
        }
        $menuItems += @{ Label = "${Gray}  ── Create new policy${Reset}"; Value = "CREATE" }
        $menuItems += @{ Label = "${Gray}  ── Back${Reset}";              Value = "BACK"   }

        # Show menu
        $selected = Invoke-PolicyListMenu -Items $menuItems -Title "Policy Manager  — Select a policy"

        if ($null -eq $selected -or $selected -eq "BACK") { return }
        if ($selected -eq "QUIT")                         { return }

        if ($selected -eq "CREATE") {
            Invoke-CreatePolicy
            continue
        }

        # Policy selected — show action sub-menu
        $policy = $selected
        $action = Show-PolicyActions -Policy $policy

        switch ($action) {
            "view"   { Show-PolicyViewer -PolicyName $policy.Name }
            "edit"   { Show-PolicyEditor -PolicyName $policy.Name }
            "assign" { Invoke-AssignPolicyToUser -PolicyName $policy.Name }
            "reset"  { Invoke-ResetUserToDefault }
            "delete" {
                if ($policy.IsDefault) {
                    Write-Ansi "${Red}  Cannot delete the default OWA policy.${Reset}"
                    Wait-KeyPress
                } else {
                    Invoke-DeletePolicy -Policy $policy
                }
            }
        }
    }
}

# ── Scrollable policy list with action keys ───────────────────
# Returns the selected Value or $null
function Invoke-PolicyListMenu {
    param([array]$Items, [string]$Title)

    $selected = 0
    $count    = $Items.Count

    while ($true) {
        Clear-Host
        Write-AppHeader -Subtitle $Title

        # Column headers
        Write-Ansi "  ${Gray}[DEFAULT]  Name                                        Users${Reset}"
        Write-Rule -Color $Gray

        for ($i = 0; $i -lt $count; $i++) {
            if ($i -eq $selected) {
                Write-Ansi "  ${Rev}${Cyan}${Bold}$($Items[$i].Label)${Reset}"
            } else {
                Write-Ansi "  $($Items[$i].Label)"
            }
        }

        Write-KeyHints -Hints @("↑↓ Navigate", "ENTER Select", "N New policy", "Q Back")

        $key = Read-Key
        switch ($key.Code) {
            "Up"    { if ($selected -gt 0)        { $selected-- } }
            "Down"  { if ($selected -lt $count-1) { $selected++ } }
            "Enter" { return $Items[$selected].Value }
            "Escape"{ return $null }
            { $_ -eq 'q' -or $_ -eq 'Q' } { return "BACK" }
            { $_ -eq 'n' -or $_ -eq 'N' } { return "CREATE" }
        }
    }
}

# ── Policy action sub-menu ────────────────────────────────────
# Returns action string or $null for back
function Show-PolicyActions {
    param([hashtable]$Policy)

    $actions = @(
        @{ Label = "  View (read-only)";      Value = "view"   }
        @{ Label = "  Edit parameters";       Value = "edit"   }
        @{ Label = "  Assign to user";        Value = "assign" }
        @{ Label = "  Reset user to default"; Value = "reset"  }
    )
    if (-not $Policy.IsDefault) {
        $actions += @{ Label = "  ${Red}Delete this policy${Reset}"; Value = "delete" }
    }
    $actions += @{ Label = "  Back"; Value = "back" }

    $defaultTag = if ($Policy.IsDefault) { " ${Blue}(Default Policy)${Reset}" } else { "" }
    return Show-Menu -Title "Actions — $($Policy.Name)$defaultTag" -Items $actions `
                     -Hints @("↑↓ Navigate", "ENTER Select", "Q Back")
}

# ── Fetch all policies and their user counts ──────────────────
function Get-AllPoliciesWithCounts {
    $policies    = Get-OwaMailboxPolicy -ErrorAction Stop
    $casMailboxes = @(Get-CasMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue)

    # Build policy → user count map
    $countMap = @{}
    foreach ($mbx in $casMailboxes) {
        $policyName = $mbx.OwaMailboxPolicy
        if ($policyName) {
            # Policy name in CasMailbox is the full identity; extract display name
            $shortName = $policyName -replace '^.*\\', ''  # strip "Org\Policy" prefix
            if ($countMap.ContainsKey($shortName)) { $countMap[$shortName]++ }
            else                                   { $countMap[$shortName] = 1 }
        }
    }

    $result = @()
    foreach ($p in $policies | Sort-Object -Property Name) {
        $name      = $p.Name
        $shortName = $name -replace '^.*\\', ''
        $result   += @{
            Name      = $name
            ShortName = $shortName
            IsDefault = [bool]$p.IsDefault
            UserCount = if ($countMap.ContainsKey($shortName)) { $countMap[$shortName] } else { 0 }
        }
    }
    return $result
}

# ── Create new policy ─────────────────────────────────────────
function Invoke-CreatePolicy {
    Clear-Host
    Write-AppHeader -Subtitle "Create New Policy"
    Write-Ansi ""
    Write-Ansi "  ${White}Enter a name for the new OWA Mailbox Policy:${Reset}"
    Write-Ansi "  ${Gray}(Letters, numbers, hyphens and spaces. No special characters.)${Reset}"
    Write-Ansi ""
    Write-Ansi "  ${White}Name: ${Reset}" -NoNewline

    $name = Invoke-LineInput -Default "" -Color $White
    if ($null -eq $name -or $name.Trim() -eq "") {
        Write-Ansi "${Yellow}  Cancelled.${Reset}"
        Start-Sleep -Milliseconds 600
        return
    }

    $name = $name.Trim()

    Write-Ansi ""
    Write-Ansi "${Gray}  Creating policy '${White}$name${Gray}'...${Reset}"

    try {
        New-OwaMailboxPolicy -Name $name -ErrorAction Stop | Out-Null
        Write-Ansi "${Green}  ✓ Policy '$name' created.${Reset}"
        Write-Ansi ""
        Write-Ansi "  Opening editor..."
        Start-Sleep -Milliseconds 800
        Show-PolicyEditor -PolicyName $name
    } catch {
        $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
        Write-Box -Title "Error" -BorderColor $Red -Lines @(
            "${Red}Failed to create policy:${Reset}",
            "",
            "${Yellow}$msg${Reset}"
        )
        Wait-KeyPress
    }
}

# ── Assign policy to user ─────────────────────────────────────
function Invoke-AssignPolicyToUser {
    param([string]$PolicyName)

    Clear-Host
    Write-AppHeader -Subtitle "Assign Policy — $PolicyName"
    Write-Ansi ""
    Write-Ansi "${Gray}  Loading mailboxes (this may take a moment)...${Reset}"

    try {
        $mailboxes = @(Get-CasMailbox -ResultSize Unlimited -ErrorAction Stop |
                       Select-Object DisplayName, PrimarySmtpAddress, OwaMailboxPolicy |
                       Sort-Object DisplayName)
    } catch {
        Write-Ansi "${Red}  Error: $($_.Exception.Message.Split([Environment]::NewLine)[0])${Reset}"
        Wait-KeyPress
        return
    }

    if ($mailboxes.Count -eq 0) {
        Write-Ansi "${Yellow}  No mailboxes found.${Reset}"
        Wait-KeyPress
        return
    }

    # Search-driven user picker
    $user = Invoke-UserPicker -Mailboxes $mailboxes -Title "Select user to assign policy: $PolicyName"
    if ($null -eq $user) { return }

    # Confirm
    Clear-Host
    Write-AppHeader -Subtitle "Confirm Assignment"
    Write-Ansi ""
    Write-Box -Title "Confirm Assignment" -BorderColor $Cyan -Lines @(
        "",
        "  ${White}User  :${Reset} $($user.DisplayName) ${Gray}<$($user.PrimarySmtpAddress)>${Reset}",
        "  ${White}Policy:${Reset} ${Cyan}$PolicyName${Reset}",
        ""
    )

    if (-not (Confirm-Action "Apply this assignment?")) {
        Write-Ansi "${Yellow}  Cancelled.${Reset}"
        Start-Sleep -Milliseconds 600
        return
    }

    Write-Ansi ""
    Write-Ansi "${Gray}  Assigning policy...${Reset}"
    try {
        Set-CasMailbox -Identity $user.PrimarySmtpAddress -OwaMailboxPolicy $PolicyName -ErrorAction Stop
        Write-Ansi "${Green}  ✓ Policy assigned to $($user.DisplayName).${Reset}"
    } catch {
        $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
        Write-Ansi "${Red}  Error: $msg${Reset}"
    }
    Wait-KeyPress
}

# ── Reset user policy to org default ─────────────────────────
function Invoke-ResetUserToDefault {
    Clear-Host
    Write-AppHeader -Subtitle "Reset User to Default Policy"
    Write-Ansi ""
    Write-Ansi "${Gray}  Loading mailboxes...${Reset}"

    try {
        $mailboxes = @(Get-CasMailbox -ResultSize Unlimited -ErrorAction Stop |
                       Select-Object DisplayName, PrimarySmtpAddress, OwaMailboxPolicy |
                       Sort-Object DisplayName)
    } catch {
        Write-Ansi "${Red}  Error: $($_.Exception.Message.Split([Environment]::NewLine)[0])${Reset}"
        Wait-KeyPress
        return
    }

    $user = Invoke-UserPicker -Mailboxes $mailboxes -Title "Select user to reset to default OWA policy"
    if ($null -eq $user) { return }

    Clear-Host
    Write-AppHeader -Subtitle "Confirm Reset"
    Write-Ansi ""
    Write-Box -Title "Confirm Reset" -BorderColor $Yellow -Lines @(
        "",
        "  ${White}User           :${Reset} $($user.DisplayName) ${Gray}<$($user.PrimarySmtpAddress)>${Reset}",
        "  ${White}Current policy :${Reset} ${Yellow}$($user.OwaMailboxPolicy)${Reset}",
        "  ${White}Action         :${Reset} Reset to organization default",
        ""
    )

    if (-not (Confirm-Action "Reset this user to the default OWA policy?")) {
        Write-Ansi "${Yellow}  Cancelled.${Reset}"
        Start-Sleep -Milliseconds 600
        return
    }

    Write-Ansi ""
    Write-Ansi "${Gray}  Resetting...${Reset}"
    try {
        # Fetch the organization default policy — it's the one where IsDefault = $true
        $defaultPolicy = Get-OwaMailboxPolicy -ErrorAction Stop |
                         Where-Object { $_.IsDefault -eq $true } |
                         Select-Object -First 1
        if (-not $defaultPolicy) {
            Write-Ansi "${Red}  Error: Could not determine the default OWA policy.${Reset}"
            Wait-KeyPress
            return
        }
        Set-CasMailbox -Identity $user.PrimarySmtpAddress -OwaMailboxPolicy $defaultPolicy.Name -ErrorAction Stop
        Write-Ansi "${Green}  ✓ $($user.DisplayName) assigned to default policy '$($defaultPolicy.Name)'.${Reset}"
    } catch {
        $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
        Write-Ansi "${Red}  Error: $msg${Reset}"
    }
    Wait-KeyPress
}

# ── Delete a policy ───────────────────────────────────────────
function Invoke-DeletePolicy {
    param([hashtable]$Policy)

    if ($Policy.IsDefault) {
        Write-Ansi "${Red}  The default policy cannot be deleted.${Reset}"
        Wait-KeyPress
        return
    }

    Clear-Host
    Write-AppHeader -Subtitle "Delete Policy"
    Write-Ansi ""

    Write-Box -Title "Confirm Deletion" -BorderColor $Red -Lines @(
        "",
        "  ${Red}${Bold}This action is IRREVERSIBLE.${Reset}",
        "",
        "  Policy : ${Cyan}$($Policy.Name)${Reset}",
        "  Users  : ${Yellow}$($Policy.UserCount) assigned user$(if($Policy.UserCount -ne 1){'s'})${Reset}",
        "",
        "  Assigned users will fall back to the default OWA policy.",
        ""
    )

    if (-not (Confirm-Action "Permanently delete '$($Policy.Name)'?")) {
        Write-Ansi "${Yellow}  Cancelled.${Reset}"
        Start-Sleep -Milliseconds 600
        return
    }

    Write-Ansi ""
    Write-Ansi "${Gray}  Deleting...${Reset}"

    try {
        Remove-OwaMailboxPolicy -Identity $Policy.Name -Confirm:$false -ErrorAction Stop
        Write-Ansi "${Green}  ✓ Policy '$($Policy.Name)' deleted.${Reset}"
    } catch {
        $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
        Write-Ansi "${Red}  Error: $msg${Reset}"
    }
    Wait-KeyPress
}

# ── Searchable user picker ────────────────────────────────────
# Returns selected mailbox object or $null
function Invoke-UserPicker {
    param(
        [array]$Mailboxes,
        [string]$Title
    )

    $filter      = ""
    $selected    = 0
    $viewOffset  = 0

    while ($true) {
        $ts         = Get-TermSize
        $viewHeight = $ts.H - 12

        # Filter mailboxes by search string
        $filtered = if ($filter) {
            $Mailboxes | Where-Object {
                $_.DisplayName -match [regex]::Escape($filter) -or
                $_.PrimarySmtpAddress -match [regex]::Escape($filter)
            }
        } else {
            $Mailboxes
        }

        $count = @($filtered).Count
        if ($selected -ge $count) { $selected = [Math]::Max(0, $count - 1) }
        $viewOffset = Sync-ViewOffset -SelIdx $selected -VOffset $viewOffset -VHeight $viewHeight

        Clear-Host
        Write-AppHeader -Subtitle $Title

        # Search bar
        Write-Ansi "  ${Gray}Search:${Reset} ${White}$filter${Cyan}▌${Reset}"
        Write-Rule -Color $Gray

        if ($count -eq 0) {
            Write-Ansi "${Yellow}  No mailboxes match '$filter'${Reset}"
        } else {
            for ($i = $viewOffset; $i -lt [Math]::Min($viewOffset + $viewHeight, $count); $i++) {
                $mbx  = @($filtered)[$i]
                $curr = [string]$mbx.OwaMailboxPolicy -replace '^.*\\', ''
                $policyTag = if ($curr) { "${Gray}[$curr]${Reset}" } else { "${Gray}[default]${Reset}" }

                if ($i -eq $selected) {
                    Write-Ansi "  ${Rev}${Cyan}${Bold} $($mbx.DisplayName.PadRight(30)) $($mbx.PrimarySmtpAddress.PadRight(35)) $policyTag ${Reset}"
                } else {
                    Write-Ansi "  ${White}$($mbx.DisplayName.PadRight(30))${Reset} ${Gray}$($mbx.PrimarySmtpAddress.PadRight(35))${Reset} $policyTag"
                }
            }
        }

        Write-KeyHints -Hints @("↑↓ Navigate", "Type to search", "ENTER Select", "ESC Cancel")

        $key = Read-Key
        switch ($key.Code) {
            "Up"        { if ($selected -gt 0) { $selected-- } }
            "Down"      { if ($selected -lt $count-1) { $selected++ } }
            "Enter"     {
                if ($count -gt 0) { return @($filtered)[$selected] }
            }
            "Escape"    { return $null }
            "Backspace" {
                if ($filter.Length -gt 0) {
                    $filter   = $filter.Substring(0, $filter.Length - 1)
                    $selected = 0
                }
            }
            default {
                $ch = $key.Char
                if ($ch -ge [char]32 -and $ch -le [char]126) {
                    $filter   += $ch
                    $selected  = 0
                }
            }
        }
    }
}
