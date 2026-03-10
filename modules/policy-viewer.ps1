# ============================================================
# modules/policy-viewer.ps1 — OWA Policy Viewer (read-only)
# ============================================================
# Displays all parameters of an OWA Mailbox Policy in a
# scrollable 4-section layout:
#
#   ✅ TRUE   — boolean params set to True
#   ❌ FALSE  — boolean params set to False
#   ⬜ NULL   — null or empty params
#   📝 VALUES — non-boolean params with actual content
#
# Pressing ENTER on a parameter opens the editor focused on it.
# ============================================================

# ── Fetch policy params from Exchange Online ──────────────────
# Returns @{ Params = [ordered]@{...}; TypeMap = @{...} }
# TypeMap stores the .NET type name for each parameter
function Get-PolicyParams {
    param([string]$PolicyName)

    try {
        $policy    = Get-OwaMailboxPolicy -Identity $PolicyName -ErrorAction Stop
        $params    = [System.Collections.Specialized.OrderedDictionary]::new()
        $typeMap   = @{}

        # Internal/readonly properties we skip — they are identity fields,
        # not settable via Set-OwaMailboxPolicy
        $skipProps = @(
            'RunspaceId','PSComputerName','PSShowComputerName','PSSourceJobInstanceId',
            'Identity','Id','Name','DistinguishedName','Guid','ObjectCategory',
            'ObjectClass','ObjectState','OrganizationId','OriginatingServer',
            'ExchangeVersion','IsValid','WhenChanged','WhenCreatedUTC','WhenChangedUTC',
            'WhenCreated','IsDefault','AdminDisplayName','ExchangeObjectId',
            'ImmutableId','IsDefaultPublicFolderMailbox'
        )

        foreach ($prop in $policy.PSObject.Properties | Sort-Object Name) {
            if ($prop.Name -in $skipProps) { continue }
            $params[$prop.Name]  = $prop.Value
            $typeMap[$prop.Name] = $prop.TypeNameOfValue
        }

        return @{ Params = $params; TypeMap = $typeMap; Policy = $policy }
    } catch {
        throw "Failed to fetch policy '$PolicyName': $($_.Exception.Message.Split([Environment]::NewLine)[0])"
    }
}

# ── Build the flat display list used by viewer and editor ─────
# Returns an array of items:
#   @{ Type="header"|"param"; Section; Name; Value; OriginalValue; IsModified; IsBoolean; TypeName }
function Build-DisplayList {
    param(
        [System.Collections.Specialized.OrderedDictionary]$CurrentValues,
        [hashtable]$OriginalValues,   # Name → original value (may be empty)
        [hashtable]$ModifiedSet,      # Set of param names that were edited
        [hashtable]$TypeMap           # Name → type name string
    )

    $true_params  = [System.Collections.Generic.List[hashtable]]::new()
    $false_params = [System.Collections.Generic.List[hashtable]]::new()
    $null_params  = [System.Collections.Generic.List[hashtable]]::new()
    $value_params = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($key in $CurrentValues.Keys) {
        $val        = $CurrentValues[$key]
        $typeName   = if ($TypeMap.ContainsKey($key)) { $TypeMap[$key] } else { "" }
        $isModified = $ModifiedSet.ContainsKey($key)
        $origVal    = if ($isModified) { $OriginalValues[$key] } else { $val }

        # Classify as boolean if type says so, or if value is actual bool
        $isBoolean = ($val -is [bool]) -or
                     ($typeName -match 'Boolean') -or
                     ($typeName -match 'Microsoft\.Exchange.*bool' )

        $entry = @{
            Type          = "param"
            Name          = $key
            Value         = $val
            OriginalValue = $origVal
            IsModified    = $isModified
            IsBoolean     = $isBoolean
            TypeName      = $typeName
        }

        if     ($val -eq $true)                            { $true_params.Add($entry)  }
        elseif ($val -eq $false)                           { $false_params.Add($entry) }
        elseif ($null -eq $val -or [string]$val -eq '')    { $null_params.Add($entry)  }
        else                                               { $value_params.Add($entry) }
    }

    $list = [System.Collections.Generic.List[hashtable]]::new()

    if ($true_params.Count -gt 0) {
        $list.Add(@{ Type = "header"; Section = "TRUE";   Label = " TRUE " })
        foreach ($e in $true_params  | Sort-Object { $_.Name }) { $list.Add($e) }
    }
    if ($false_params.Count -gt 0) {
        $list.Add(@{ Type = "header"; Section = "FALSE";  Label = " FALSE " })
        foreach ($e in $false_params | Sort-Object { $_.Name }) { $list.Add($e) }
    }
    if ($null_params.Count -gt 0) {
        $list.Add(@{ Type = "header"; Section = "NOTSET"; Label = " NOT SET " })
        foreach ($e in $null_params  | Sort-Object { $_.Name }) { $list.Add($e) }
    }
    if ($value_params.Count -gt 0) {
        $list.Add(@{ Type = "header"; Section = "VALUES"; Label = " VALUES " })
        foreach ($e in $value_params | Sort-Object { $_.Name }) { $list.Add($e) }
    }

    return $list.ToArray()
}

# ── Format a parameter value for display ─────────────────────
function Format-ParamValue {
    param($Value, [string]$Section, [switch]$IsModified)

    if ($null -eq $Value -or [string]$Value -eq '') {
        return "${Yellow}(null)${Reset}"
    }
    if ($Value -eq $true) {
        return "${Green}True${Reset}"
    }
    if ($Value -eq $false) {
        return "${Red}False${Reset}"
    }
    if ($Value -is [System.Array]) {
        $joined = $Value -join ", "
        return "${White}$joined${Reset}"
    }
    # Strings and other types
    $s = [string]$Value
    return "${White}$s${Reset}"
}

# ── Render the section header row ────────────────────────────
function Render-SectionHeader {
    param([string]$Section, [int]$Width)

    $icon  = switch ($Section) {
        "TRUE"   { "${Green}✅${Reset}" }
        "FALSE"  { "${Red}❌${Reset}" }
        "NOTSET" { "${Yellow}⬜${Reset}" }
        "VALUES" { "${Blue}📝${Reset}" }
        default  { " " }
    }
    $label = switch ($Section) {
        "TRUE"   { "${Green}${Bold} TRUE ${Reset}" }
        "FALSE"  { "${Red}${Bold} FALSE ${Reset}" }
        "NOTSET" { "${Yellow}${Bold} NOT SET ${Reset}" }
        "VALUES" { "${Blue}${Bold} VALUES ${Reset}" }
        default  { "  $Section  " }
    }

    $cleanLabel = [regex]::Replace($label, '\x1b\[[0-9;]*m', '')
    $lineLen    = $Width - $cleanLabel.Length - 4  # 4 = 2 dashes + 2 spaces
    if ($lineLen -lt 2) { $lineLen = 2 }
    $line = "${Gray}──${Reset} $icon $label ${Gray}" + ('─' * $lineLen) + "${Reset}"
    return $line
}

# ── Render one parameter row ──────────────────────────────────
function Render-ParamRow {
    param(
        [hashtable]$Item,
        [bool]$IsSelected,
        [int]$Width,
        [switch]$IsViewer   # Viewer = no edit hints
    )
    $nameW = 42  # Fixed column width for param names

    # Modified marker
    $marker = if ($Item.IsModified) { "${Yellow}*${Reset}" } else { " " }

    # Name column
    $name = $Item.Name
    if ($name.Length -gt $nameW - 1) { $name = $name.Substring(0, $nameW - 4) + "..." }
    $namePad = $name.PadRight($nameW)

    # Value column
    $valStr = Format-ParamValue -Value $Item.Value -Section $Item.Section

    # Modified: show original → new
    $diffStr = ""
    if ($Item.IsModified) {
        $origStr = Format-ParamValue -Value $Item.OriginalValue -Section "ORIG"
        $diffStr = "  ${Gray}(was: $origStr${Gray})${Reset}"
    }

    $row = " $marker ${White}$namePad${Reset}  $valStr$diffStr"

    if ($IsSelected) {
        # Strip ANSI for reverse-video rendering (cursor highlight)
        return "${Rev}${Cyan}${Bold} $marker $namePad${Reset}${Rev}  $valStr${Reset}"
    }
    return $row
}

# ╔══════════════════════════════════════════════════════════════╗
# ║               4-COLUMN LAYOUT HELPERS                        ║
# ╚══════════════════════════════════════════════════════════════╝

# ── Build 4-column data structure from current values ─────────
# Returns PSCustomObject with .True .False .NotSet .Values arrays
# Each entry: @{ Type; Name; Value; OriginalValue; IsModified; IsBoolean; TypeName }
function Build-ColItems {
    param(
        [System.Collections.Specialized.OrderedDictionary]$CurrentValues,
        [hashtable]$OriginalValues,
        [hashtable]$ModifiedSet,
        [hashtable]$TypeMap
    )

    $c0 = [System.Collections.Generic.List[hashtable]]::new()
    $c1 = [System.Collections.Generic.List[hashtable]]::new()
    $c2 = [System.Collections.Generic.List[hashtable]]::new()
    $c3 = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($key in $CurrentValues.Keys) {
        $val        = $CurrentValues[$key]
        $typeName   = if ($TypeMap.ContainsKey($key)) { $TypeMap[$key] } else { "" }
        $isModified = $ModifiedSet.ContainsKey($key)
        $origVal    = if ($isModified) { $OriginalValues[$key] } else { $val }
        $isBoolean  = ($val -is [bool]) -or ($typeName -match 'Boolean')

        $entry = @{
            Type          = "param"
            Name          = $key
            Value         = $val
            OriginalValue = $origVal
            IsModified    = $isModified
            IsBoolean     = $isBoolean
            TypeName      = $typeName
        }

        if     ($val -eq $true)                          { $c0.Add($entry) }
        elseif ($val -eq $false)                         { $c1.Add($entry) }
        elseif ($null -eq $val -or [string]$val -eq '')  { $c2.Add($entry) }
        else                                             { $c3.Add($entry) }
    }

    return [PSCustomObject]@{
        True   = @($c0 | Sort-Object { $_.Name })
        False  = @($c1 | Sort-Object { $_.Name })
        NotSet = @($c2 | Sort-Object { $_.Name })
        Values = @($c3 | Sort-Object { $_.Name })
    }
}

# ── Sync per-column scroll offset ─────────────────────────────
function Sync-ColView {
    param([int]$Cursor, [int]$Offset, [int]$VHeight)
    if ($Cursor -lt $Offset)            { return $Cursor }
    if ($Cursor -ge $Offset + $VHeight) { return $Cursor - $VHeight + 1 }
    return $Offset
}

# ── Find a param by name across all 4 columns ─────────────────
# Returns @{ Col = colIdx; Idx = itemIdx } or $null
function Find-ParamInCols {
    param([PSCustomObject]$ColData, [string[]]$ColProps, [string]$Name)
    for ($c = 0; $c -lt 4; $c++) {
        $col = $ColData.($ColProps[$c])
        for ($i = 0; $i -lt $col.Count; $i++) {
            if ($col[$i].Name -eq $Name) { return @{ Col = $c; Idx = $i } }
        }
    }
    return $null
}

# ── Clamp column cursors to valid bounds after rebuild ─────────
function Clamp-ColCursors {
    param([PSCustomObject]$ColData, [string[]]$ColProps, [int[]]$Cursors)
    for ($c = 0; $c -lt 4; $c++) {
        $cnt = $ColData.($ColProps[$c]).Count
        if ($cnt -gt 0 -and $Cursors[$c] -ge $cnt) { $Cursors[$c] = $cnt - 1 }
    }
}

# ── Render one column header cell ──────────────────────────────
# Returns a string of exactly $Width visual characters
function Render-ColHeader {
    param([int]$ColIdx, [int]$ItemCount, [bool]$IsActive, [int]$Width)

    $icons   = @("✅", "❌", "⬜", "📝")
    $labels  = @(" TRUE", " FALSE", " NOT SET", " VALUES")
    $colors  = @($Green, $Red, $Yellow, $Blue)

    $icon      = $icons[$ColIdx]
    $label     = $labels[$ColIdx]
    $color     = $colors[$ColIdx]
    $countStr  = " ($ItemCount)"
    $plainText = "$icon$label$countStr"

    # Emoji are double-width in terminal (+1 visual col vs .NET Length)
    $visLen = $plainText.Length + 1

    if ($IsActive) {
        $raw = "${Rev}${color}${Bold}${plainText}${Reset}"
    } else {
        $raw = "${color}${Bold}${plainText}${Reset}"
    }

    $pad = [Math]::Max(0, $Width - $visLen)
    return $raw + (' ' * $pad)
}

# ── Render one column item cell ────────────────────────────────
# Returns a string of exactly $Width visual characters
function Render-ColCell {
    param(
        [hashtable]$Item,
        [bool]$IsSelected,
        [int]$ColIdx,   # 0=TRUE, 1=FALSE, 2=NOTSET, 3=VALUES
        [int]$Width
    )

    # Layout: " M text" = 1 space + 1 marker + 1 space + text
    $textMaxW = [Math]::Max(1, $Width - 3)

    # Build plain display text
    $displayText = if ($ColIdx -eq 3) {
        $valStr = if ($Item.Value -is [System.Array]) {
            $Item.Value -join ", "
        } else { [string]$Item.Value }
        "$($Item.Name)=$valStr"
    } else {
        $Item.Name
    }

    # Truncate plain text to fit
    if ($displayText.Length -gt $textMaxW) {
        $displayText = $displayText.Substring(0, [Math]::Max(0, $textMaxW - 1)) + [char]0x2026
    }

    $marker = if ($Item.IsModified) { "*" } else { " " }
    $visLen = 3 + $displayText.Length   # " M " + text
    $pad    = [Math]::Max(0, $Width - $visLen)

    if ($IsSelected) {
        return "${Rev}${Cyan}${Bold} $marker $displayText${Reset}" + (' ' * $pad)
    }

    $nameColor = switch ($ColIdx) {
        0 { $Green  }
        1 { $Red    }
        2 { $Yellow }
        3 { $White  }
    }
    $markerStr = if ($Item.IsModified) { "${Yellow}*${Reset}" } else { " " }
    return " $markerStr ${nameColor}${displayText}${Reset}" + (' ' * $pad)
}

# ── Main viewer screen (4-column layout) ──────────────────────
# Shows the policy in read-only mode. Returns "back" when done.
function Show-PolicyViewer {
    param([string]$PolicyName)

    if (-not (Assert-Connected)) { return "back" }

    Clear-Host
    Write-AppHeader -Subtitle "Loading policy..."
    Write-Ansi "${Gray}  Fetching $PolicyName from Exchange Online...${Reset}"

    try {
        $result  = Get-PolicyParams -PolicyName $PolicyName
        $params  = $result.Params
        $typeMap = $result.TypeMap
    } catch {
        Write-Ansi "${Red}  Error: $_${Reset}"
        Wait-KeyPress
        return "back"
    }

    $colProps   = @("True", "False", "NotSet", "Values")
    $colData    = Build-ColItems -CurrentValues $params -OriginalValues @{} -ModifiedSet @{} -TypeMap $typeMap
    $colOffsets = @(0, 0, 0, 0)
    $colCursors = @(0, 0, 0, 0)
    $activeCol  = 0
    for ($c = 0; $c -lt 4; $c++) {
        if ($colData.($colProps[$c]).Count -gt 0) { $activeCol = $c; break }
    }

    while ($true) {
        $ts         = Get-TermSize
        $w          = $ts.W
        $viewHeight = [Math]::Max(1, $ts.H - 8)  # header(4) + colhdr(1) + divider(1) + footer(2)
        $colW       = [Math]::Max(8, [int](($w - 3) / 4))
        $col3W      = [Math]::Max(8, $w - ($colW * 3) - 3)
        $widths     = @($colW, $colW, $colW, $col3W)

        # Sync offsets for all columns
        for ($c = 0; $c -lt 4; $c++) {
            $colOffsets[$c] = Sync-ColView -Cursor $colCursors[$c] -Offset $colOffsets[$c] -VHeight $viewHeight
        }

        # [Console]::CursorVisible = $false
        # Clear-Host
        [Console]::CursorVisible = $false
        [Console]::SetCursorPosition(0, 0)
        Write-AppHeader -Subtitle "Viewer  ›  $PolicyName"

        # Column header row
        $hLine = ""
        for ($c = 0; $c -lt 4; $c++) {
            $cnt    = $colData.($colProps[$c]).Count
            $hLine += Render-ColHeader -ColIdx $c -ItemCount $cnt -IsActive ($c -eq $activeCol) -Width $widths[$c]
            if ($c -lt 3) { $hLine += "${Gray}│${Reset}" }
        }
        Write-Ansi $hLine

        # Divider row
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

        Write-KeyHints -Hints @("Tab/→← Switch col", "↑↓ Navigate", "PgUp/Dn Jump 5", "Home/End First/Last", "E/ENTER Edit", "B/Q Back")
        [Console]::CursorVisible = $true

        $key  = Read-Key
        $code = $key.Code

        if ($code -eq "Tab" -or $code -eq "Right") {
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
        elseif ($code -eq "Enter" -or $code -eq "e" -or $code -eq "E") {
            $cnt        = $colData.($colProps[$activeCol]).Count
            $focusParam = if ($cnt -gt 0) { $colData.($colProps[$activeCol])[$colCursors[$activeCol]].Name } else { $null }
            Show-PolicyEditor -PolicyName $PolicyName -FocusParam $focusParam
            # Reload after editor
            try {
                $result     = Get-PolicyParams -PolicyName $PolicyName
                $colData    = Build-ColItems -CurrentValues $result.Params -OriginalValues @{} -ModifiedSet @{} -TypeMap $result.TypeMap
                $colOffsets = @(0, 0, 0, 0)
                $colCursors = @(0, 0, 0, 0)
            } catch { }
        }
        elseif ($code -eq "b" -or $code -eq "B" -or $code -eq "Q" -or $code -eq "q" -or $code -eq "Escape") {
            [Console]::CursorVisible = $true
            return "back"
        }
    }
}

# ── Navigation helpers ────────────────────────────────────────

# Move cursor down, skipping section headers
function Get-NextParamIdx {
    param([array]$Items, [int]$CurrentIdx)
    $idx = $CurrentIdx + 1
    while ($idx -lt $Items.Count) {
        if ($Items[$idx].Type -eq "param") { return $idx }
        $idx++
    }
    return $CurrentIdx
}

# Move cursor up, skipping section headers
function Get-PrevParamIdx {
    param([array]$Items, [int]$CurrentIdx)
    $idx = $CurrentIdx - 1
    while ($idx -ge 0) {
        if ($Items[$idx].Type -eq "param") { return $idx }
        $idx--
    }
    return $CurrentIdx
}

# First selectable param index
function Find-FirstParam {
    param([array]$Items)
    for ($i = 0; $i -lt $Items.Count; $i++) {
        if ($Items[$i].Type -eq "param") { return $i }
    }
    return 0
}

# Last selectable param index
function Find-LastParam {
    param([array]$Items)
    for ($i = $Items.Count - 1; $i -ge 0; $i--) {
        if ($Items[$i].Type -eq "param") { return $i }
    }
    return 0
}

# Find a param by name; return its flat index
function Find-ParamByName {
    param([array]$Items, [string]$Name)
    for ($i = 0; $i -lt $Items.Count; $i++) {
        if ($Items[$i].Type -eq "param" -and $Items[$i].Name -eq $Name) { return $i }
    }
    return (Find-FirstParam $Items)
}

# Adjust viewOffset so that selectedIdx is visible
function Sync-ViewOffset {
    param([int]$SelIdx, [int]$VOffset, [int]$VHeight)
    if ($SelIdx -lt $VOffset)                  { return $SelIdx }
    if ($SelIdx -ge $VOffset + $VHeight)       { return $SelIdx - $VHeight + 1 }
    return $VOffset
}

# When page changes, snap selection to first visible param
function Sync-SelectionToView {
    param([array]$Items, [int]$VOffset, [int]$VHeight, [int]$CurrentSel)
    # If current selection is visible, keep it
    if ($CurrentSel -ge $VOffset -and $CurrentSel -lt $VOffset + $VHeight) {
        return $CurrentSel
    }
    # Otherwise find first param in the new viewport
    for ($i = $VOffset; $i -lt [Math]::Min($VOffset + $VHeight, $Items.Count); $i++) {
        if ($Items[$i].Type -eq "param") { return $i }
    }
    return $CurrentSel
}
