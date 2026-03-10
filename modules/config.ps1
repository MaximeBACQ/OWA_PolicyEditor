# ============================================================
# modules/config.ps1 — Local Configuration (config.json)
# ============================================================
# Reads and writes a JSON config file in the script root.
# Stores: last used UPN, UI preferences.
# Loaded on startup; updated after each successful connection.
# ============================================================

# In-memory config object (populated by Initialize-Config)
$script:Config = $null

# Config getter
function Get-DefaultConfig {
    return [PSCustomObject]@{
        LastUPN     = ""
        Preferences = [PSCustomObject]@{
            ColorEnabled = $true
        }
    }
}

# Load config.json from disk; create with defaults if missing
function Initialize-Config {
    $configFile = (Get-AppPaths).ConfigFile

    # if (Test-Path $global:ConfigFile) {
    if (Test-Path $configFile) {
        try {
            $raw = Get-Content $global:ConfigFile -Raw -Encoding UTF8
            $script:Config = $raw | ConvertFrom-Json
            # Ensure all expected keys exist (forward-compatibility)
            if ($null -eq $script:Config.LastUPN)     { $script:Config | Add-Member -NotePropertyName 'LastUPN'     -NotePropertyValue '' -Force }
            if ($null -eq $script:Config.Preferences)  { $script:Config | Add-Member -NotePropertyName 'Preferences'  -NotePropertyValue ([PSCustomObject]@{ ColorEnabled = $true }) -Force }
        } catch {
            # Corrupted config — reset to defaults
            $script:Config = Get-DefaultConfig
            Save-Config
        }
    } else {
        $script:Config = Get-DefaultConfig
        Save-Config
    }
}

# Persist current config to disk
function Save-Config {
    try {
        $configFile = (Get-AppPaths).ConfigFile
        $script:Config | ConvertTo-Json -Depth 5 | Set-Content $configFile -Encoding UTF8
        # $script:Config | ConvertTo-Json -Depth 5 | Set-Content $global:ConfigFile -Encoding UTF8
    } catch {
        # Non-fatal: config save failure doesn't stop operation
        Write-Ansi "${Yellow}  Warning: Could not save config — $($_.Exception.Message)${Reset}"
    }
}

# Convenience: get the last UPN used (empty string if none)
function Get-LastUPN {
    if ($null -eq $script:Config) { return "" }
    return [string]$script:Config.LastUPN
}

# Convenience: persist the UPN for next launch
function Set-LastUPN {
    param([string]$UPN)
    if ($null -eq $script:Config) { Initialize-Config }
    $script:Config.LastUPN = $UPN
    Save-Config
}
