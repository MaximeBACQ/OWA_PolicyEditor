# ============================================================
# modules/logger.ps1 — Change Logger
# ============================================================
# Appends structured change records to exchange-owa-manager.log
# in the script root. Only called AFTER the user confirms with
# [Y] on the confirmation screen. Never overwrites existing log.
#
# Log format per save operation:
#   [DD/MM/YYYY HH:mm] TENANT: <tenant> | POLICY: <policy>
#     ParameterName   : OldValue  →  NewValue
# ============================================================

# Write a single change block to the log file.
# $Changes is an array of @{ Name; OldValue; NewValue }
function Write-ChangeLog {
    param(
        [string]$PolicyName,
        [string]$TenantName,
        [array]$Changes       # each: @{ Name; OldValue; NewValue }
    )

    if ($Changes.Count -eq 0) { return }

    try {
        $timestamp = (Get-Date).ToString("dd/MM/yyyy HH:mm")
        $lines     = [System.Collections.Generic.List[string]]::new()

        $lines.Add("[${timestamp}] TENANT: ${TenantName} | POLICY: ${PolicyName}")

        # Compute column width for alignment
        $maxNameLen = ($Changes | Measure-Object { $_.Name.Length } -Maximum).Maximum
        if ($maxNameLen -lt 20) { $maxNameLen = 20 }

        foreach ($c in $Changes) {
            $oldDisplay = Format-LogValue $c.OldValue
            $newDisplay = Format-LogValue $c.NewValue
            $namePad    = $c.Name.PadRight($maxNameLen + 1)
            $lines.Add("  ${namePad}: ${oldDisplay}  ->  ${newDisplay}")
        }

        $lines.Add("")  # Blank line between blocks

        Add-Content -Path $LogFile -Value $lines -Encoding UTF8
    } catch {
        Write-Ansi "${Yellow}  Warning: Could not write to log file — $($_.Exception.Message)${Reset}"
    }
}

# Format a value for log display
function Format-LogValue {
    param($Value)
    if ($null -eq $Value -or $Value -eq '')     { return "(empty)" }
    if ($Value -is [bool])                       { return $Value.ToString() }
    if ($Value -is [System.Array])               { return $Value -join ", " }
    return [string]$Value
}
