# Write-Host "DEBUG ENTERING server-init.ps1"
# ============================================================
# server-init.ps1 — Shared constants and TUI stubs for Pode
# ============================================================
# Loaded first via Use-PodeScript so every Pode runspace has:
#   - App-level path constants the modules read at call time
#   - Empty ANSI color vars (modules reference them; we don't render)
#   - No-op stubs for every TUI function modules might call in
#     catch/warning blocks — prevents "command not found" errors
#
# $PSScriptRoot here is the project root because Use-PodeScript
# dot-sources with the full path, which sets PSScriptRoot correctly.
# ============================================================

# ── Path constants ────────────────────────────────────────────
# $global:LogFile    = Join-Path $PSScriptRoot 'exchange-owa-manager.log'
# $global:ConfigFile = Join-Path $PSScriptRoot 'config.json'

# $root = Get-PodeState -Name 'ServerRoot'
# $global:LogFile    = Join-Path $using:serverRoot 'exchange-owa-manager.log'
# $global:ConfigFile = Join-Path $using:serverRoot 'config.json'

# $root = Get-PodeState -Name 'ServerRoot'
# Write-Host "DEBUG server-init root=[$root]"
# $global:LogFile    = Join-Path $root 'exchange-owa-manager.log'
# $global:ConfigFile = Join-Path $root 'config.json'

$root = $PSScriptRoot   # This is the project root when dot-sourced via Use-PodeScript
Write-Host "DEBUG server-init root=[$root]"

# $global:LogFile    = Join-Path $root 'exchange-owa-manager.log'
# $global:ConfigFile = Join-Path $root 'config.json'

function global:Get-AppPaths {
    $root = Get-PodeState -Name 'ServerRoot'
    return @{
        ConfigFile = Join-Path $root 'config.json'
        LogFile    = Join-Path $root 'exchange-owa-manager.log'
    }
}

# $root = Split-Path -Parent $MyInvocation.MyCommand.Path
# Write-Host "DEBUG server-init: MyInvocation.MyCommand.Path=[$($MyInvocation.MyCommand.Path)] PSScriptRoot=[$PSScriptRoot]"

# $global:LogFile    = Join-Path $root 'exchange-owa-manager.log'
# $global:ConfigFile = Join-Path $root 'config.json'

# ── ANSI color vars — empty strings, not real escape codes ───
# Modules embed these in strings passed to Write-Ansi, which
# strips ANSI anyway. Setting to empty avoids $null substitution.
$global:ESC    = [char]27
$global:Reset  = ''
$global:Bold   = ''
$global:Dim    = ''
$global:Rev    = ''
$global:Green  = ''
$global:Red    = ''
$global:Yellow = ''
$global:Blue   = ''
$global:Cyan   = ''
$global:White  = ''
$global:Gray   = ''

# ── Write-Ansi stub ───────────────────────────────────────────
# Modules call this in catch blocks for warnings. In server context
# we strip ANSI codes and write to the console for visibility.
function global:Write-Ansi {
    param([string]$Text, [switch]$NoNewline)
    $clean = [regex]::Replace($Text, '\x1b\[[0-9;]*m', '')
    if ($NoNewline) { Write-Host $clean -NoNewline } else { Write-Host $clean }
}

# ── TUI function stubs ────────────────────────────────────────
# These are referenced by module functions that we never call from
# routes (Show-PolicyEditor, Invoke-CreatePolicy, etc.). They must
# exist to prevent "command not found" if any indirect call path
# reaches them — all are no-ops in server context.
function global:Get-TermSize        { return @{ W = 80; H = 24 } }
function global:Read-Key            { return @{ Code = ''; Char = [char]0; VK = 0; Ctrl = $false; Alt = $false } }
function global:Wait-KeyPress       { param([string]$Message) }
function global:Confirm-Action      { param([string]$Question, [string]$Default = 'N'); return $false }
function global:Invoke-LineInput    { param([string]$Prompt, [string]$Default, [string]$Color); return $null }
function global:Write-AppHeader     { param([string]$Subtitle) }
function global:Write-Box           { param([string]$Title, [string[]]$Lines, [string]$BorderColor, [int]$MinWidth) }
function global:Write-Rule          { param([string]$Label, [string]$Color, [int]$Width) }
function global:Write-KeyHints      { param([string[]]$Hints) }
function global:Show-Menu           { param([string]$Title, [array]$Items, [string[]]$Hints); return $null }
function global:Format-Cell         { param([string]$Text, [int]$Width); return $Text }
function global:Get-ConnectionStatusLine { return '' }

# ── Param value serializer ────────────────────────────────────
# Convert a raw Exchange param value to a JSON-safe primitive.
# Exchange returns MultiValuedProperty, enums, and other .NET types
# that ConvertTo-Json chokes on or misrepresents.
function global:ConvertTo-SafeParamValue {
    param($Value)
    if ($null -eq $Value)    { return $null }
    if ($Value -is [bool])   { return $Value }
    if ($Value -is [int])    { return $Value }
    if ($Value -is [long])   { return $Value }
    if ($Value -is [string]) { return $Value }

    # Arrays and Exchange MultiValuedProperty — flatten to string array
    $type = $Value.GetType()
    if ($type.IsArray -or ($Value -is [System.Collections.IEnumerable])) {
        return @($Value | ForEach-Object { if ($null -eq $_) { $null } else { $_.ToString() } })
    }

    # Enum, Exchange object, anything else — string representation
    return $Value.ToString()
}

# [Uri]::UnescapeDataString is always available without loading System.Web
function global:Decode-RouteParam {
    param([string]$Value)
    return [Uri]::UnescapeDataString($Value)
}
