# ============================================================
# server.ps1 — Pode web server entry point
# ============================================================
# Route definitions only. All Exchange logic lives in /modules/.
# Run with: pwsh -File server.ps1
# Then open: http://localhost:8080
# ============================================================

Import-Module Pode -ErrorAction Stop

$serverRoot = $PSScriptRoot

# $global:serverRoot = $PSScriptRoot

Start-PodeServer -Threads 1 -Verbose{ # Deactivate threading (ps works bad with it)

    Add-PodeEndpoint -Address localhost -Port 8080 -Protocol Http
    Set-PodeState -Name 'ServerRoot' -Value $serverRoot #this line creates a separate thread-safe Hashtable outside of PS's runspace

    New-PodeLoggingMethod -Terminal | Enable-PodeErrorLogging

    Add-PodeStaticRoute -Path '/' -Source "$serverRoot/public" -Defaults @('index.html')

    # Add-PodeStaticRoute -Path '/' -Source "$global:serverRoot/public" -Defaults @('index.html')

    # Add-PodeStaticRoute -Path '/public' -Source "$global:serverRoot/public"

    # Add-PodeRoute -Method Get -Path '/' -ScriptBlock {
    #     Move-PodeResponseUrl -Path './public/index.html'
    # }
    # server-init.ps1 must load first: defines $global:LogFile, $global:ConfigFile,
    # ANSI stubs, Write-Ansi, TUI no-ops, and ConvertTo-SafeParamValue.
    Use-PodeScript -Path "$serverRoot/server-init.ps1"
    # Use-PodeScript -Path "$global:serverRoot/server-init.ps1"

    # Business logic modules — order matters: viewer defines Get-PolicyParams
    # which editor and manager both depend on.
    Use-PodeScript -Path "$serverRoot/modules/config.ps1"
    Use-PodeScript -Path "$serverRoot/modules/logger.ps1"
    Use-PodeScript -Path "$serverRoot/modules/auth.ps1"
    Use-PodeScript -Path "$serverRoot/modules/policy-viewer.ps1"
    Use-PodeScript -Path "$serverRoot/modules/policy-editor.ps1"
    Use-PodeScript -Path "$serverRoot/modules/policy-manager.ps1"

    # use global var because $using only works on distant machines in pode scripts
    # Use-PodeScript -Path "$global:serverRoot/modules/config.ps1"
    # Use-PodeScript -Path "$global:serverRoot/modules/logger.ps1"
    # Use-PodeScript -Path "$global:serverRoot/modules/auth.ps1"
    # Use-PodeScript -Path "$global:serverRoot/modules/policy-viewer.ps1"
    # Use-PodeScript -Path "$global:serverRoot/modules/policy-editor.ps1"
    # Use-PodeScript -Path "$global:serverRoot/modules/policy-manager.ps1"

    # ── One-time lazy init ────────────────────────────────────
    # config.ps1 uses $script:Config which is null in the worker runspace
    # until Initialize-Config runs there. We lazy-init on the first
    # request rather than needing a separate Pode startup hook.
    Add-PodeMiddleware -Name 'AppInit' -ScriptBlock {
        if ($null -eq $script:Config) {
            Write-Host "DEBUG serverRoot=[$global:serverRoot] ConfigFile=[$global:ConfigFile] LogFile=[$global:LogFile]"
            Initialize-Config
        }
        return $true
    }

    # ── Auth guard ────────────────────────────────────────────
    # Every /api/* route except connect/status/disconnect requires
    # an active Exchange Online session.
    Add-PodeMiddleware -Name 'ExchangeAuth' -Route '/api/*' -ScriptBlock {
        $openRoutes = @('/api/connect', '/api/status', '/api/disconnect')
        if ($WebEvent.Path -in $openRoutes) { return $true }

        if (-not (Test-ExchangeConnected)) {
            Set-PodeResponseStatus -Code 401
            Write-PodeJsonResponse -Value @{ error = 'Not connected to Exchange Online' }
            return $false
        }
        return $true
    }

    # ==========================================================
    #  AUTH ROUTES
    # ==========================================================

    # GET /api/status — polled by the UI on load to restore session state
    Add-PodeRoute -Method Get -Path '/api/status' -ScriptBlock {
        try {
            $connected = Test-ExchangeConnected
            $info      = Get-PodeState -Name 'ConnectionInfo'
            Write-PodeJsonResponse -Value @{
                connected = $connected
                upn       = if ($connected -and $info) { $info.UPN    } else { '' }
                tenant    = if ($connected -and $info) { $info.Tenant } else { '' }
            }
        } catch {
            Set-PodeResponseStatus -Code 500
            Write-PodeJsonResponse -Value @{ error = $_.Exception.Message.Split([Environment]::NewLine)[0] }
        }
    }

    # POST /api/connect
    # Body: { upn: "admin@tenant.onmicrosoft.com" }
    # Triggers the OAuth browser popup — route blocks until MFA completes.
    Add-PodeRoute -Method Post -Path '/api/connect' -ScriptBlock {
        try {
            $upn = [string]$WebEvent.Data.upn
            if (-not $upn -or $upn.Trim() -eq '') {
                Set-PodeResponseStatus -Code 400
                Write-PodeJsonResponse -Value @{ error = 'UPN is required' }
                return
            }
            $upn = $upn.Trim()

            if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
                Set-PodeResponseStatus -Code 500
                Write-PodeJsonResponse -Value @{ error = 'ExchangeOnlineManagement module not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser' }
                return
            }

            if (-not (Get-Module ExchangeOnlineManagement)) {
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
            }

            Connect-ExchangeOnline -UserPrincipalName $upn -ShowBanner:$false -ErrorAction Stop

            $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
            $tenant    = if ($orgConfig) { $orgConfig.Name } else { 'Unknown' }

            # Persist for status queries and for Write-ChangeLog calls
            Set-PodeState -Name 'ConnectionInfo' -Value @{ UPN = $upn; Tenant = $tenant }

            # Keep module-level script vars consistent — Write-ChangeLog reads $script:ConnectedTenant
            $script:ConnectedUPN    = $upn
            $script:ConnectedTenant = $tenant

            Set-LastUPN $upn

            Write-PodeJsonResponse -Value @{ connected = $true; upn = $upn; tenant = $tenant }
        } catch {
            Set-PodeResponseStatus -Code 500
            Write-PodeJsonResponse -Value @{ error = $_.Exception.Message.Split([Environment]::NewLine)[0] }
        }
    }

    # POST /api/disconnect
    Add-PodeRoute -Method Post -Path '/api/disconnect' -ScriptBlock {
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Set-PodeState -Name 'ConnectionInfo' -Value @{ UPN = ''; Tenant = '' }
            $script:ConnectedUPN    = ''
            $script:ConnectedTenant = ''
            Write-PodeJsonResponse -Value @{ disconnected = $true }
        } catch {
            Set-PodeResponseStatus -Code 500
            Write-PodeJsonResponse -Value @{ error = $_.Exception.Message.Split([Environment]::NewLine)[0] }
        }
    }

    # ==========================================================
    #  POLICY ROUTES
    # ==========================================================

    # GET /api/policies
    Add-PodeRoute -Method Get -Path '/api/policies' -ScriptBlock {
        try {
            $raw    = Get-AllPoliciesWithCounts
            $result = @($raw | ForEach-Object {
                @{
                    name      = $_.Name
                    shortName = $_.ShortName
                    isDefault = [bool]$_.IsDefault
                    userCount = [int]$_.UserCount
                }
            })
            Write-PodeJsonResponse -Value $result
        } catch {
            Set-PodeResponseStatus -Code 500
            Write-PodeJsonResponse -Value @{ error = $_.Exception.Message.Split([Environment]::NewLine)[0] }
        }
    }

    # GET /api/policies/:name/params
    Add-PodeRoute -Method Get -Path '/api/policies/:name/params' -ScriptBlock {
        try {
            $name   = Decode-RouteParam $WebEvent.Parameters['name']
            $result = Get-PolicyParams -PolicyName $name

            # Convert ordered dict of Exchange objects to JSON-safe flat hashtable
            $safeParams = [ordered]@{}
            foreach ($key in $result.Params.Keys) {
                $safeParams[$key] = ConvertTo-SafeParamValue $result.Params[$key]
            }

            Write-PodeJsonResponse -Value @{
                params  = $safeParams
                typeMap = $result.TypeMap
            }
        } catch {
            Set-PodeResponseStatus -Code 500
            Write-PodeJsonResponse -Value @{ error = $_.Exception.Message.Split([Environment]::NewLine)[0] }
        }
    }

    # PATCH /api/policies/:name
    # Body: { changes: [{ name, oldValue, newValue }] }
    # Note: we cannot use Apply-PolicyChanges (the module wrapper) because it
    # calls Wait-KeyPress which would block the route handler indefinitely.
    # Instead we replicate only the Exchange call + Write-ChangeLog.
    Add-PodeRoute -Method Patch -Path '/api/policies/:name' -ScriptBlock {
        try {
            $policyName = Decode-RouteParam $WebEvent.Parameters['name']
            $changes    = @($WebEvent.Data.changes)

            if ($changes.Count -eq 0) {
                Write-PodeJsonResponse -Value @{ success = $true; applied = 0 }
                return
            }

            # Fetch TypeMap so we can cast values to the types Set-OwaMailboxPolicy expects
            $typeMap   = (Get-PolicyParams -PolicyName $policyName).TypeMap
            $setParams = @{ Identity = $policyName }
            $logChanges = @()

            foreach ($c in $changes) {
                $paramName = [string]$c.name
                $newVal    = $c.newValue
                $typeName  = if ($typeMap.ContainsKey($paramName)) { $typeMap[$paramName] } else { '' }

                # Cast to correct .NET type — Exchange rejects wrong types silently or with errors
                if ($typeName -match 'Boolean' -or $newVal -is [bool]) {
                    $newVal = [bool]$newVal
                } elseif ($typeName -match 'Int' -and "$newVal" -match '^\d+$') {
                    $newVal = [int]$newVal
                } elseif ($null -eq $newVal -or "$newVal" -eq '') {
                    $newVal = $null
                } elseif ($newVal -is [System.Collections.IEnumerable] -and -not ($newVal -is [string])) {
                    $newVal = @($newVal | ForEach-Object { [string]$_ })
                }

                $setParams[$paramName] = $newVal
                $logChanges += @{ Name = $paramName; OldValue = $c.oldValue; NewValue = $newVal }
            }

            Set-OwaMailboxPolicy @setParams -ErrorAction Stop

            $info   = Get-PodeState -Name 'ConnectionInfo'
            $tenant = if ($info) { $info.Tenant } else { 'Unknown' }
            Write-ChangeLog -PolicyName $policyName -TenantName $tenant -Changes $logChanges

            Write-PodeJsonResponse -Value @{ success = $true; applied = $changes.Count }
        } catch {
            Set-PodeResponseStatus -Code 500
            Write-PodeJsonResponse -Value @{ error = $_.Exception.Message.Split([Environment]::NewLine)[0] }
        }
    }

    # POST /api/policies — create new policy
    # Body: { name: "PolicyName" }
    Add-PodeRoute -Method Post -Path '/api/policies' -ScriptBlock {
        try {
            $name = [string]$WebEvent.Data.name
            if (-not $name -or $name.Trim() -eq '') {
                Set-PodeResponseStatus -Code 400
                Write-PodeJsonResponse -Value @{ error = 'Policy name is required' }
                return
            }
            $name = $name.Trim()
            New-OwaMailboxPolicy -Name $name -ErrorAction Stop | Out-Null
            Write-PodeJsonResponse -Value @{ created = $true; name = $name }
        } catch {
            Set-PodeResponseStatus -Code 500
            Write-PodeJsonResponse -Value @{ error = $_.Exception.Message.Split([Environment]::NewLine)[0] }
        }
    }

    # DELETE /api/policies/:name
    Add-PodeRoute -Method Delete -Path '/api/policies/:name' -ScriptBlock {
        try {
            $name = Decode-RouteParam $WebEvent.Parameters['name']
            Remove-OwaMailboxPolicy -Identity $name -Confirm:$false -ErrorAction Stop
            Write-PodeJsonResponse -Value @{ deleted = $true; name = $name }
        } catch {
            Set-PodeResponseStatus -Code 500
            Write-PodeJsonResponse -Value @{ error = $_.Exception.Message.Split([Environment]::NewLine)[0] }
        }
    }

    # ==========================================================
    #  USER ROUTES
    # ==========================================================

    # GET /api/users
    Add-PodeRoute -Method Get -Path '/api/users' -ScriptBlock {
        try {
            $mailboxes = @(Get-CasMailbox -ResultSize Unlimited -ErrorAction Stop |
                          Select-Object DisplayName, PrimarySmtpAddress, OwaMailboxPolicy |
                          Sort-Object DisplayName)

            $result = @($mailboxes | ForEach-Object {
                # Strip "OrgName\PolicyName" prefix that Exchange includes in the identity
                $policy = [string]$_.OwaMailboxPolicy -replace '^.*\\', ''
                @{
                    displayName        = [string]$_.DisplayName
                    primarySmtpAddress = [string]$_.PrimarySmtpAddress
                    owaMailboxPolicy   = $policy
                }
            })
            Write-PodeJsonResponse -Value $result
        } catch {
            Set-PodeResponseStatus -Code 500
            Write-PodeJsonResponse -Value @{ error = $_.Exception.Message.Split([Environment]::NewLine)[0] }
        }
    }

    # POST /api/users/:upn/policy — assign policy to a user
    # Body: { policyName: "PolicyName" }
    Add-PodeRoute -Method Post -Path '/api/users/:upn/policy' -ScriptBlock {
        try {
            $upn        = Decode-RouteParam $WebEvent.Parameters['upn']
            $policyName = [string]$WebEvent.Data.policyName
            if (-not $policyName) {
                Set-PodeResponseStatus -Code 400
                Write-PodeJsonResponse -Value @{ error = 'policyName is required' }
                return
            }
            Set-CasMailbox -Identity $upn -OwaMailboxPolicy $policyName -ErrorAction Stop
            Write-PodeJsonResponse -Value @{ assigned = $true; upn = $upn; policy = $policyName }
        } catch {
            Set-PodeResponseStatus -Code 500
            Write-PodeJsonResponse -Value @{ error = $_.Exception.Message.Split([Environment]::NewLine)[0] }
        }
    }

    # DELETE /api/users/:upn/policy — reset user to org default
    Add-PodeRoute -Method Delete -Path '/api/users/:upn/policy' -ScriptBlock {
        try {
            $upn = Decode-RouteParam $WebEvent.Parameters['upn']

            # Same logic as Invoke-ResetUserToDefault: find the IsDefault policy
            $defaultPolicy = Get-OwaMailboxPolicy -ErrorAction Stop |
                             Where-Object { $_.IsDefault -eq $true } |
                             Select-Object -First 1

            if (-not $defaultPolicy) {
                Set-PodeResponseStatus -Code 500
                Write-PodeJsonResponse -Value @{ error = 'Could not determine the default OWA policy' }
                return
            }

            Set-CasMailbox -Identity $upn -OwaMailboxPolicy $defaultPolicy.Name -ErrorAction Stop
            Write-PodeJsonResponse -Value @{ reset = $true; upn = $upn; defaultPolicy = $defaultPolicy.Name }
        } catch {
            Set-PodeResponseStatus -Code 500
            Write-PodeJsonResponse -Value @{ error = $_.Exception.Message.Split([Environment]::NewLine)[0] }
        }
    }

    # ==========================================================
    #  LOG ROUTE
    # ==========================================================

    # GET /api/log
    Add-PodeRoute -Method Get -Path '/api/log' -ScriptBlock {
        try {
            if (-not (Test-Path $global:LogFile)) {
                Write-PodeJsonResponse -Value @{ lines = @() }
                return
            }
            $lines = @(Get-Content $global:LogFile -Encoding UTF8)
            Write-PodeJsonResponse -Value @{ lines = $lines }
        } catch {
            Set-PodeResponseStatus -Code 500
            Write-PodeJsonResponse -Value @{ error = $_.Exception.Message.Split([Environment]::NewLine)[0] }
        }
    }
}
