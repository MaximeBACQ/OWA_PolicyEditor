# ============================================================
# modules/policy-editor.ps1 — OWA Policy Editor (vim-inspired)
# ============================================================
# Interactive editor for OWA Mailbox Policy parameters.
#
# Navigation:
#   ↑↓     — move between parameters
#   ENTER  — toggle boolean / open text input for others
#   CTRL+X — exit (with unsaved-changes dialog if needed)
#   B/Q    — back (same as CTRL+X)
#
# Modified params are marked with * and shown with before→after.
# No changes are written until the user confirms with [Y].
# ============================================================

# ── Main editor entry point (4-column layout) ─────────────────
# $PolicyName  : Identity string for Set/Get-OwaMailboxPolicy
# $FocusParam  : Optional param name to start cursor on
function Show-PolicyEditor {
    param(
        [string]$PolicyName,
        [string]$FocusParam = ""
    )

    if (-not (Assert-Connected)) { return }

    Clear-Host
    Write-AppHeader -Subtitle "Loading editor..."
    Write-Ansi "${Gray}  Fetching $PolicyName...${Reset}"

    try {
        $result  = Get-PolicyParams -PolicyName $PolicyName
        $typeMap = $result.TypeMap
    } catch {
        Write-Ansi "${Red}  Error: $_${Reset}"
        Wait-KeyPress
        return
    }

    # Working copy of values — $currentValues is mutated in-place as edits happen
    $currentValues  = $result.Params
    $originalValues = @{}
    foreach ($k in $currentValues.Keys) { $originalValues[$k] = $currentValues[$k] }
    $modifiedSet = @{}

    $colProps   = @("True", "False", "NotSet", "Values")
    $colData    = Build-ColItems -CurrentValues $currentValues -OriginalValues $originalValues -ModifiedSet $modifiedSet -TypeMap $typeMap
    $colOffsets = @(0, 0, 0, 0)
    $colCursors = @(0, 0, 0, 0)
    $activeCol  = 0

    # Set initial focus
    if ($FocusParam) {
        $loc = Find-ParamInCols -ColData $colData -ColProps $colProps -Name $FocusParam
        if ($null -ne $loc) {
            $activeCol = $loc.Col
            $colCursors[$activeCol] = $loc.Idx
        }
    } else {
        for ($c = 0; $c -lt 4; $c++) {
            if ($colData.($colProps[$c]).Count -gt 0) { $activeCol = $c; break }
        }
    }

    while ($true) {
        $ts         = Get-TermSize
        $w          = $ts.W
        $viewHeight = [Math]::Max(1, $ts.H - 8)
        $colW       = [Math]::Max(8, [int](($w - 3) / 4))
        $col3W      = [Math]::Max(8, $w - ($colW * 3) - 3)
        $widths     = @($colW, $colW, $colW, $col3W)

        # Sync offsets for all columns
        for ($c = 0; $c -lt 4; $c++) {
            $colOffsets[$c] = Sync-ColView -Cursor $colCursors[$c] -Offset $colOffsets[$c] -VHeight $viewHeight
        }

        [Console]::CursorVisible = $false
        Clear-Host

        $modCount  = $modifiedSet.Count
        $modSuffix = if ($modCount -gt 0) { "  ${Yellow}● $modCount unsaved change$(if($modCount -ne 1){'s'})${Reset}" } else { "" }
        Write-AppHeader -Subtitle "Editor  ›  $PolicyName$modSuffix"

        # Column headers
        $hLine = ""
        for ($c = 0; $c -lt 4; $c++) {
            $cnt    = $colData.($colProps[$c]).Count
            $hLine += Render-ColHeader -ColIdx $c -ItemCount $cnt -IsActive ($c -eq $activeCol) -Width $widths[$c]
            if ($c -lt 3) { $hLine += "${Gray}│${Reset}" }
        }
        Write-Ansi $hLine

        # Divider
        $dLine = ""
        for ($c = 0; $c -lt 4; $c++) {
            $bc     = if ($c -eq $activeCol) { $Cyan } else { $Gray }
            $dLine += "${bc}" + ('─' * $widths[$c]) + "${Reset}"
            if ($c -lt 3) { $dLine += "${Gray}┼${Reset}" }
        }
        Write-Ansi $dLine

        # Content rows
        for ($row = 0; $row -lt $viewHeight; $row++) {
            $rowLine = ""
            for ($c = 0; $c -lt 4; $c++) {
                $col     = $colData.($colProps[$c])
                $itemIdx = $colOffsets[$c] + $row
                $cW      = $widths[$c]
                if ($itemIdx -lt $col.Count) {
                    $isSel   = ($c -eq $activeCol -and $itemIdx -eq $colCursors[$c])
                    $cell    = Render-ColCell -Item $col[$itemIdx] -IsSelected $isSel -ColIdx $c -Width $cW
                } else {
                    $cell = ' ' * $cW
                }
                $rowLine += $cell
                if ($c -lt 3) { $rowLine += "${Gray}│${Reset}" }
            }
            Write-Ansi $rowLine
        }

        Write-KeyHints -Hints @("Tab/→← Switch col", "↑↓ Navigate", "PgUp/Dn Jump 5", "ENTER Edit/Toggle", "^X Exit", "B/Q Back")
        [Console]::CursorVisible = $true

        $key  = Read-Key
        $code = $key.Code
        $ctrl = $key.Ctrl

        # CTRL+X — exit with unsaved-changes dialog
        if ($ctrl -and ($code -eq 'x' -or $code -eq 'X')) {
            $action = Invoke-ExitDialog -ModCount $modifiedSet.Count -PolicyName $PolicyName `
                                        -CurrentValues $currentValues -OriginalValues $originalValues `
                                        -ModifiedSet $modifiedSet -TypeMap $typeMap
            if ($action -eq "EXIT") { return }
            $colData = Build-ColItems -CurrentValues $currentValues -OriginalValues $originalValues -ModifiedSet $modifiedSet -TypeMap $typeMap
            Clamp-ColCursors -ColData $colData -ColProps $colProps -Cursors $colCursors
        }
        elseif ($code -eq "Tab" -or $code -eq "Right") {
            $activeCol = ($activeCol + 1) % 4
        }
        elseif ($code -eq "Left") {
            $activeCol = ($activeCol + 3) % 4
        }
        elseif ($code -eq "Up") {
            if ($colCursors[$activeCol] -gt 0) {
                $colCursors[$activeCol]--
                $colOffsets[$activeCol] = Sync-ColView -Cursor $colCursors[$activeCol] -Offset $colOffsets[$activeCol] -VHeight $viewHeight
            }
        }
        elseif ($code -eq "Down") {
            $cnt = $colData.($colProps[$activeCol]).Count
            if ($colCursors[$activeCol] -lt $cnt - 1) {
                $colCursors[$activeCol]++
                $colOffsets[$activeCol] = Sync-ColView -Cursor $colCursors[$activeCol] -Offset $colOffsets[$activeCol] -VHeight $viewHeight
            }
        }
        elseif ($code -eq "PageUp") {
            $colCursors[$activeCol] = [Math]::Max(0, $colCursors[$activeCol] - 5)
            $colOffsets[$activeCol] = Sync-ColView -Cursor $colCursors[$activeCol] -Offset $colOffsets[$activeCol] -VHeight $viewHeight
        }
        elseif ($code -eq "PageDown") {
            $cnt = $colData.($colProps[$activeCol]).Count
            $colCursors[$activeCol] = [Math]::Min([Math]::Max(0, $cnt - 1), $colCursors[$activeCol] + 5)
            $colOffsets[$activeCol] = Sync-ColView -Cursor $colCursors[$activeCol] -Offset $colOffsets[$activeCol] -VHeight $viewHeight
        }
        elseif ($code -eq "Home") {
            $colCursors[$activeCol] = 0
            $colOffsets[$activeCol] = 0
        }
        elseif ($code -eq "End") {
            $cnt = $colData.($colProps[$activeCol]).Count
            $colCursors[$activeCol] = [Math]::Max(0, $cnt - 1)
            $colOffsets[$activeCol] = Sync-ColView -Cursor $colCursors[$activeCol] -Offset 0 -VHeight $viewHeight
        }
        elseif ($code -eq "Enter") {
            $col = $colData.($colProps[$activeCol])
            if ($col.Count -gt 0) {
                $item      = $col[$colCursors[$activeCol]]
                $paramName = $item.Name
                $oldValue  = $item.Value

                $newValue = Invoke-ParamEdit -Item $item -PolicyName $PolicyName

                if ($null -ne $newValue) {
                    $storeValue = if ($newValue -eq "") { $null } else { $newValue }
                    $oldNorm    = Normalize-ParamValue $oldValue
                    $newNorm    = Normalize-ParamValue $storeValue

                    if ($newNorm -ne $oldNorm) {
                        if (-not $modifiedSet.ContainsKey($paramName)) {
                            $modifiedSet[$paramName] = $originalValues[$paramName]
                        }
                        if ($newNorm -eq (Normalize-ParamValue $originalValues[$paramName])) {
                            $modifiedSet.Remove($paramName)
                        }
                        $currentValues[$paramName] = $storeValue

                        # Rebuild columns — param may move to a different column
                        $colData = Build-ColItems -CurrentValues $currentValues -OriginalValues $originalValues -ModifiedSet $modifiedSet -TypeMap $typeMap
                        Clamp-ColCursors -ColData $colData -ColProps $colProps -Cursors $colCursors

                        # Follow the param to its new column
                        $loc = Find-ParamInCols -ColData $colData -ColProps $colProps -Name $paramName
                        if ($null -ne $loc) {
                            $activeCol = $loc.Col
                            $colCursors[$activeCol] = $loc.Idx
                            $colOffsets[$activeCol] = Sync-ColView -Cursor $loc.Idx -Offset $colOffsets[$activeCol] -VHeight $viewHeight
                        }
                    }
                }
            }
        }
        elseif ($code -eq "b" -or $code -eq "B" -or $code -eq "Q" -or $code -eq "q" -or $code -eq "Escape") {
            $action = Invoke-ExitDialog -ModCount $modifiedSet.Count -PolicyName $PolicyName `
                                        -CurrentValues $currentValues -OriginalValues $originalValues `
                                        -ModifiedSet $modifiedSet -TypeMap $typeMap
            if ($action -eq "EXIT") { return }
            $colData = Build-ColItems -CurrentValues $currentValues -OriginalValues $originalValues -ModifiedSet $modifiedSet -TypeMap $typeMap
            Clamp-ColCursors -ColData $colData -ColProps $colProps -Cursors $colCursors
        }
    }
}

# ── Exit dialog (CTRL+X / B / Q) ─────────────────────────────
# Returns "EXIT" or "STAY"
function Invoke-ExitDialog {
    param(
        [int]$ModCount,
        [string]$PolicyName,
        [System.Collections.Specialized.OrderedDictionary]$CurrentValues,
        [hashtable]$OriginalValues,
        [hashtable]$ModifiedSet,
        [hashtable]$TypeMap
    )

    if ($ModCount -eq 0) { return "EXIT" }

    while ($true) {
        Clear-Host
        Write-AppHeader -Subtitle "Unsaved Changes"

        Write-Box -Title "Unsaved changes" -BorderColor $Yellow -Lines @(
            "",
            "  ${Yellow}You have ${Bold}$ModCount${Reset}${Yellow} unsaved change$(if($ModCount -ne 1){'s'}).${Reset}",
            "",
            "  ${Cyan}[S]${Reset} Save and exit      ${Cyan}[D]${Reset} Discard and exit      ${Cyan}[C]${Reset} Cancel",
            ""
        )

        Write-Ansi ""
        Write-KeyHints -Hints @("S Save+Exit", "D Discard", "C Cancel / stay in editor")

        $key = Read-Key
        switch ($key.Char) {
            { $_ -eq 's' -or $_ -eq 'S' } {
                # Show confirmation diff screen
                $saved = Invoke-SaveConfirmation -PolicyName $PolicyName `
                             -CurrentValues $CurrentValues -OriginalValues $OriginalValues `
                             -ModifiedSet $ModifiedSet -TypeMap $TypeMap
                if ($saved) { return "EXIT" }
                # User pressed N → back to editor
                return "STAY"
            }
            { $_ -eq 'd' -or $_ -eq 'D' } {
                return "EXIT"   # Discard — caller does nothing, just exits
            }
            { $_ -eq 'c' -or $_ -eq 'C' -or $_ -eq [char]27 } {
                return "STAY"
            }
        }
        if ($key.Code -eq "Escape") { return "STAY" }
    }
}

# ── Save-confirmation screen (before/after diff) ──────────────
# Applies changes via Set-OwaMailboxPolicy and logs them.
# Returns $true if saved, $false if user pressed N.
function Invoke-SaveConfirmation {
    param(
        [string]$PolicyName,
        [System.Collections.Specialized.OrderedDictionary]$CurrentValues,
        [hashtable]$OriginalValues,
        [hashtable]$ModifiedSet,
        [hashtable]$TypeMap
    )

    # Build change list for display
    $changes = @()
    foreach ($name in ($ModifiedSet.Keys | Sort-Object)) {
        $changes += @{
            Name     = $name
            OldValue = $OriginalValues[$name]
            NewValue = $CurrentValues[$name]
        }
    }

    while ($true) {
        Clear-Host
        Write-AppHeader -Subtitle "Confirm Changes"

        $count = $changes.Count
        $diffLines = @(
            "",
            "  ${White}${Bold}$count change$(if($count -ne 1){'s'}) will be applied to ${Cyan}$PolicyName${Reset}",
            ""
        )

        # Column widths
        $maxNameW = ($changes | Measure-Object { $_.Name.Length } -Maximum).Maximum
        if ($maxNameW -lt 30) { $maxNameW = 30 }

        foreach ($c in $changes) {
            $oldDisp = Format-ChangeValue $c.OldValue
            $newDisp = Format-ChangeValue $c.NewValue
            $namePad = $c.Name.PadRight($maxNameW + 1)
            $diffLines += "  ${White}$namePad${Reset}  $oldDisp  ${Gray}→${Reset}  $newDisp"
        }
        $diffLines += ""
        $diffLines += "  ${Cyan}[Y]${Reset} Apply    ${Cyan}[N]${Reset} Go back to editor"
        $diffLines += ""

        Write-Box -Title "Confirm changes" -BorderColor $Cyan -Lines $diffLines

        Write-Ansi ""
        Write-KeyHints -Hints @("Y Apply changes to Exchange Online", "N Back to editor")

        $key = Read-Key
        switch ($key.Char) {
            { $_ -eq 'y' -or $_ -eq 'Y' } {
                # Apply changes
                return Apply-PolicyChanges -PolicyName $PolicyName -Changes $changes -TypeMap $TypeMap
            }
            { $_ -eq 'n' -or $_ -eq 'N' -or $_ -eq [char]27 } {
                return $false
            }
        }
        if ($key.Code -eq "Escape") { return $false }
    }
}

# ── Apply changes to Exchange Online via Set-OwaMailboxPolicy ─
function Apply-PolicyChanges {
    param(
        [string]$PolicyName,
        [array]$Changes,
        [hashtable]$TypeMap
    )

    Clear-Host
    Write-AppHeader -Subtitle "Applying Changes..."
    Write-Ansi ""

    # Build Set-OwaMailboxPolicy parameter splatting hashtable
    $setParams = @{ Identity = $PolicyName }

    foreach ($c in $Changes) {
        $name = $c.Name
        $val  = $c.NewValue
        $type = if ($TypeMap.ContainsKey($name)) { $TypeMap[$name] } else { "" }

        # Cast to correct type to avoid Exchange cmdlet errors
        if ($type -match "Boolean" -or $val -is [bool]) {
            $val = [bool]$val
        } elseif ($type -match "Int" -and $val -match '^\d+$') {
            $val = [int]$val
        } elseif ($null -eq $val -or "$val" -eq "") {
            $val = $null
        }

        $setParams[$name] = $val
        Write-Ansi "  ${Gray}Setting ${White}$name${Gray} → ${Reset}$(Format-ChangeValue $val)"
    }

    Write-Ansi ""

    try {
        Set-OwaMailboxPolicy @setParams -ErrorAction Stop
        Write-Ansi "${Green}  ✓ Changes applied successfully.${Reset}"

        # Log the changes
        $tenant = $script:ConnectedTenant
        Write-ChangeLog -PolicyName $PolicyName -TenantName $tenant -Changes $Changes

        Write-Ansi "${Gray}  Changes logged to: $LogFile${Reset}"
        Write-Ansi ""
        Wait-KeyPress
        return $true
    } catch {
        $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
        Write-Ansi ""
        Write-Box -Title "Error Applying Changes" -BorderColor $Red -Lines @(
            "${Red}Set-OwaMailboxPolicy failed:${Reset}",
            "",
            "${Yellow}$msg${Reset}",
            "",
            "Some parameters may have been applied before the error.",
            "Check Exchange Admin Center to verify the current state."
        )
        Write-Ansi ""
        Wait-KeyPress
        return $false
    }
}

# ── Inline param editor ───────────────────────────────────────
# Returns the new value (or the original if cancelled / no change)
function Invoke-ParamEdit {
    param([hashtable]$Item, [string]$PolicyName)

    $val  = $Item.Value
    $name = $Item.Name

    # ── Boolean: instant toggle ──────────────────────────────
    if ($val -eq $true)  { return $false }
    if ($val -eq $false) { return $true  }

    # ── Null value: let user choose type ────────────────────
    if ($null -eq $val -or "$val" -eq '') {
        return Invoke-NullParamEdit -ParamName $name -TypeName $Item.TypeName
    }

    # ── Array: edit as comma-separated string ───────────────
    if ($val -is [System.Array]) {
        return Invoke-ArrayParamEdit -ParamName $name -CurrentValue $val
    }

    # ── String / other ───────────────────────────────────────
    return Invoke-TextParamEdit -ParamName $name -CurrentValue ([string]$val)
}

# ── Edit a null param (ask for boolean or text) ───────────────
function Invoke-NullParamEdit {
    param([string]$ParamName, [string]$TypeName)

    # If type is known to be boolean, skip the type-choice prompt
    if ($TypeName -match "Boolean") {
        return Invoke-BoolChoice -ParamName $ParamName
    }

    Clear-Host
    Write-AppHeader -Subtitle "Edit Parameter"
    Write-Ansi ""
    Write-Ansi "  ${Cyan}$ParamName${Reset}  ${Gray}(currently null)${Reset}"
    Write-Ansi ""
    Write-Ansi "  Set as:"
    Write-Ansi "    ${Cyan}[T]${Reset}  True"
    Write-Ansi "    ${Cyan}[F]${Reset}  False"
    Write-Ansi "    ${Cyan}[V]${Reset}  Enter a text value"
    Write-Ansi "    ${Gray}[ESC]${Reset} Cancel"
    Write-Ansi ""
    Write-KeyHints -Hints @("T True", "F False", "V Text value", "ESC Cancel")

    while ($true) {
        $key = Read-Key
        switch ($key.Char) {
            { $_ -eq 't' -or $_ -eq 'T' } { return $true  }
            { $_ -eq 'f' -or $_ -eq 'F' } { return $false }
            { $_ -eq 'v' -or $_ -eq 'V' } { return Invoke-TextParamEdit -ParamName $ParamName -CurrentValue "" }
        }
        if ($key.Code -eq "Escape") { return $null }
    }
}

# ── Boolean choice prompt ─────────────────────────────────────
function Invoke-BoolChoice {
    param([string]$ParamName)
    Clear-Host
    Write-AppHeader -Subtitle "Edit Parameter"
    Write-Ansi ""
    Write-Ansi "  ${Cyan}$ParamName${Reset}"
    Write-Ansi ""
    Write-Ansi "  ${Cyan}[T]${Reset} True   ${Cyan}[F]${Reset} False   ${Gray}[ESC]${Reset} Cancel"
    Write-Ansi ""
    Write-KeyHints -Hints @("T True", "F False", "ESC Cancel")

    while ($true) {
        $key = Read-Key
        switch ($key.Char) {
            { $_ -eq 't' -or $_ -eq 'T' } { return $true  }
            { $_ -eq 'f' -or $_ -eq 'F' } { return $false }
        }
        if ($key.Code -eq "Escape") { return $null }
    }
}

# ── Text param editor ─────────────────────────────────────────
function Invoke-TextParamEdit {
    param([string]$ParamName, [string]$CurrentValue)

    Clear-Host
    Write-AppHeader -Subtitle "Edit Parameter"
    Write-Ansi ""
    Write-Ansi "  ${Cyan}$ParamName${Reset}"
    Write-Ansi "  ${Gray}Current: ${Yellow}$(if($CurrentValue){"'$CurrentValue'"}else{'(empty)'})${Reset}"
    Write-Ansi ""
    Write-Ansi "  ${Gray}Type new value and press ENTER. ESC to cancel.${Reset}"
    Write-Ansi "  ${Gray}Leave blank and press ENTER to clear the value.${Reset}"
    Write-Ansi ""
    Write-Ansi "  ${White}New value: ${Reset}" -NoNewline

    $result = Invoke-LineInput -Default $CurrentValue -Color $White
    if ($null -eq $result) { return $null }   # ESC = cancel — caller will skip

    # Return "" to signal "clear this field" (stored as $null in Exchange)
    if ($result.Trim() -eq "") { return "" }
    return $result
}

# ── Array param editor ────────────────────────────────────────
function Invoke-ArrayParamEdit {
    param([string]$ParamName, $CurrentValue)

    $currentStr = if ($CurrentValue) { $CurrentValue -join ", " } else { "" }

    Clear-Host
    Write-AppHeader -Subtitle "Edit Parameter"
    Write-Ansi ""
    Write-Ansi "  ${Cyan}$ParamName${Reset}  ${Gray}(multi-value — separate with commas)${Reset}"
    Write-Ansi "  ${Gray}Current: ${Yellow}$(if($currentStr){"'$currentStr'"}else{'(empty)'})${Reset}"
    Write-Ansi ""
    Write-Ansi "  ${Gray}Enter comma-separated values. Leave blank to clear.${Reset}"
    Write-Ansi ""
    Write-Ansi "  ${White}New value: ${Reset}" -NoNewline

    $result = Invoke-LineInput -Default $currentStr -Color $White
    if ($null -eq $result) { return $null }
    if ($result.Trim() -eq "") { return @() }

    # Split by comma and trim each entry
    $parts = $result -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    return $parts
}

# ── Normalize a param value to a stable string for comparison ─
# Treats $null and "" as identical (both = "empty")
function Normalize-ParamValue {
    param($Value)
    if ($null -eq $Value)          { return "__null__" }
    if ($Value -eq $true)          { return "True" }
    if ($Value -eq $false)         { return "False" }
    if ($Value -eq "")             { return "__null__" }
    if ($Value -is [System.Array]) { return ($Value | Sort-Object) -join "," }
    return [string]$Value
}

# ── Format a value for the diff/confirm screen ────────────────
function Format-ChangeValue {
    param($Value)
    if ($null -eq $Value)           { return "${Yellow}(empty)${Reset}" }
    if ($Value -eq $true)           { return "${Green}True${Reset}" }
    if ($Value -eq $false)          { return "${Red}False${Reset}" }
    if ($Value -is [System.Array])  { return "${White}$($Value -join ', ')${Reset}" }
    $s = [string]$Value
    if ($s -eq '')                  { return "${Yellow}(empty)${Reset}" }
    return "${White}$s${Reset}"
}
