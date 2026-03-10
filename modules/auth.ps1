# ============================================================
# modules/auth.ps1 — Exchange Online Authentication
# ============================================================
# Handles connect / disconnect / status for Exchange Online.
#
# Auth flow:
#   1. User enters UPN (pre-filled from config.json)
#   2. Connect-ExchangeOnline triggers Microsoft OAuth browser
#      window (handles MFA automatically)
#   3. After success, tenant name is read and displayed
#   4. UPN is saved to config for next launch
#
# Dependency: ExchangeOnlineManagement PowerShell module
#   Install-Module ExchangeOnlineManagement -Scope CurrentUser
# ============================================================

# Cached connection info after successful login
$script:ConnectedUPN    = ""
$script:ConnectedTenant = ""

# ── Check if an active Exchange Online session exists ─────────
function Test-ExchangeConnected {
    try {
        # Get-ConnectionInformation is available in EXO v3+
        $conn = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if ($conn -and $conn.State -eq 'Connected') {
            return $true
        }
        return $false
    } catch {
        return $false
    }
}

# ── Return a one-line status string for the header ────────────
function Get-ConnectionStatusLine {
    if (Test-ExchangeConnected) {
        return " ${Green}● Connected${Reset}  ${Gray}$script:ConnectedUPN${Reset}  ${Gray}|${Reset}  ${Blue}$script:ConnectedTenant${Reset}"
    } else {
        return " ${Red}● Not connected${Reset}  ${Gray}Run Connect from Main Menu or press ENTER${Reset}"
    }
}

# ── Interactive auth screen ───────────────────────────────────
# Shows a UPN prompt, attempts connection, returns "OK" or "QUIT"
function Show-AuthScreen {
    while ($true) {
        Clear-Host
        Write-AppHeader -Subtitle "Sign In"

        Write-Ansi ""
        Write-Ansi "${Blue}  Exchange Online Authentication${Reset}"
        Write-Ansi "${Gray}  A browser window will open for Microsoft OAuth / MFA.${Reset}"
        Write-Ansi ""

        # Check that the module is installed
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Box -Title "Module Not Found" -BorderColor $Red -Lines @(
                "${Red}ExchangeOnlineManagement module is not installed.${Reset}",
                "",
                "Run the following command and restart the script:",
                "${Cyan}  Install-Module ExchangeOnlineManagement -Scope CurrentUser${Reset}"
            )
            Write-Ansi ""
            Wait-KeyPress "Press any key to exit..."
            return "QUIT"
        }

        # Pre-fill last UPN from config
        $lastUPN = Get-LastUPN
        Write-Ansi "${White}  Enter your admin UPN:${Reset}"
        Write-Ansi "" -NoNewline
        Write-Ansi "  " -NoNewline
        $upn = Invoke-LineInput -Prompt "" -Default $lastUPN -Color $White

        if ($null -eq $upn -or $upn.Trim() -eq "") {
            Write-Ansi ""
            Write-Ansi "${Yellow}  No UPN entered. Press Q to quit or any key to retry.${Reset}"
            $key = Read-Key
            if ($key.Char -eq 'q' -or $key.Char -eq 'Q') { return "QUIT" }
            continue
        }

        $upn = $upn.Trim()

        Write-Ansi ""
        Write-Ansi "${Gray}  Connecting as ${White}$upn${Gray}...${Reset}"
        Write-Ansi "${Gray}  (A browser window will open for authentication)${Reset}"
        Write-Ansi ""

        try {
            # Import module if not already loaded
            if (-not (Get-Module ExchangeOnlineManagement)) {
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
            }

            Connect-ExchangeOnline -UserPrincipalName $upn -ShowBanner:$false -ErrorAction Stop

            # Fetch tenant info
            $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
            $tenant    = if ($orgConfig) { $orgConfig.Name } else { "Unknown" }

            $script:ConnectedUPN    = $upn
            $script:ConnectedTenant = $tenant
            Set-LastUPN $upn

            Write-Ansi "${Green}  Connected successfully!${Reset}"
            Write-Ansi "${Blue}  Tenant : ${White}$tenant${Reset}"
            Write-Ansi "${Blue}  Account: ${White}$upn${Reset}"
            Write-Ansi ""
            Wait-KeyPress "Press any key to continue..."
            return "OK"

        } catch {
            $msg = $_.Exception.Message
            Write-Ansi ""
            Write-Box -Title "Connection Failed" -BorderColor $Red -Lines @(
                "${Red}Authentication error:${Reset}",
                "",
                # Trim message to avoid raw stack traces
                "${Yellow}$($msg.Split([Environment]::NewLine)[0])${Reset}"
            )
            Write-Ansi ""
            Write-Ansi "${Gray}  Press Q to quit or any key to retry.${Reset}"
            $key = Read-Key
            if ($key.Char -eq 'q' -or $key.Char -eq 'Q') { return "QUIT" }
        }
    }
}

# ── Disconnect from Exchange Online ───────────────────────────
function Disconnect-FromExchange {
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    } catch { }
    $script:ConnectedUPN    = ""
    $script:ConnectedTenant = ""
    Write-Ansi "${Green}  Disconnected.${Reset}"
    Start-Sleep -Milliseconds 800
}

# ── Guard: redirect to auth if not connected ─────────────────
# Call this at the top of any function that requires a connection.
# Returns $true if connected, $false if user quit from auth.
function Assert-Connected {
    if (Test-ExchangeConnected) { return $true }
    $result = Show-AuthScreen
    return ($result -eq "OK")
}
