# ============================================================
# modules/gpo-placeholder.ps1 — GPO / Registry Module (STUB)
# ============================================================
#
# PURPOSE (future implementation):
#   Apply registry keys to block personal account addition in
#   classic Outlook Win32 via Group Policy Objects (GPO) or
#   direct registry writes.
#
#   Target hive : HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Options
#   Registry keys to set:
#     DisableOffice365SimplifiedAccountCreation  DWORD = 1
#       Prevents the simplified OAuth account-add flow used by
#       personal Microsoft accounts in Outlook 2016+/365.
#
#     DisableIMAP  DWORD = 1
#       Disables IMAP account addition in Outlook.
#
#     DisablePOP3  DWORD = 1
#       Disables POP3 account addition in Outlook.
#
#   These keys mirror Administrative Template (ADMX) policies
#   from the Microsoft 365 Apps ADMX pack. GPO deployment is
#   the recommended method in domain environments; direct
#   registry writes can be used for standalone/test machines.
#
# TODO:
#   - List: enumerate current registry values on local machine
#   - Apply: write registry keys (optional: push via Invoke-GPUpdate)
#   - Export: generate a .reg file for manual import
#   - GPO: optionally link/create a GPO via RSAT GroupPolicy module
# ============================================================

function Show-GPOPlaceholder {
    Clear-Host
    Write-AppHeader -Subtitle "GPO / Registry Settings"
    Write-Ansi ""

    Write-Box -Title "Coming Soon" -BorderColor $Blue -Lines @(
        "",
        "  ${Bold}${Blue}GPO / Registry Hardening Module${Reset}",
        "",
        "  This module will allow you to block personal account",
        "  addition in classic Outlook Win32 by applying registry",
        "  keys that mirror Group Policy settings.",
        "",
        "  ${Gray}Target:${Reset}",
        "  ${Cyan}HKCU\\Software\\Policies\\Microsoft\\Office\\16.0\\Outlook\\Options${Reset}",
        "",
        "  ${Gray}Planned registry keys:${Reset}",
        "  ${White}DisableOffice365SimplifiedAccountCreation${Reset}  ${Yellow}DWORD = 1${Reset}",
        "  ${White}DisableIMAP${Reset}                               ${Yellow}DWORD = 1${Reset}",
        "  ${White}DisablePOP3${Reset}                               ${Yellow}DWORD = 1${Reset}",
        "",
        "  ${Gray}Planned features:${Reset}",
        "  ${Gray}  •${Reset} View current registry state",
        "  ${Gray}  •${Reset} Apply keys locally or via GPO",
        "  ${Gray}  •${Reset} Export as .reg file",
        "  ${Gray}  •${Reset} GPO link via RSAT GroupPolicy module",
        ""
    )

    Write-Ansi ""
    Write-KeyHints -Hints @("Any key to go back")
    $null = Read-Key
}
