<#
.SYNOPSIS
  Enable "Reload/Open in Internet Explorer mode" button in Microsoft Edge (machine policy).

.NOTES
  - Run this script elevated (as Administrator).
  - This writes Group Policy registry keys under HKLM so they apply to all users on the device.
  - Edge must be restarted for policy to take effect. The script will stop running Edge processes (msedge).
#>

# Ensure we are running elevated
If (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

$policyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"

# Create policy key if missing
If (-not (Test-Path -Path $policyPath)) {
    New-Item -Path $policyPath -Force | Out-Null
}

# 1) Enable IE integration (IEMode)
# InternetExplorerIntegrationLevel:
#   0 = None, 1 = IEMode, 2 = NeedIE (IE11)
New-ItemProperty -Path $policyPath -Name "InternetExplorerIntegrationLevel" -PropertyType DWord -Value 1 -Force | Out-Null

# 2) Allow users to reload unconfigured sites into IE mode
# InternetExplorerIntegrationReloadInIEModeAllowed: 0 = Don't allow, 1 = Allow
New-ItemProperty -Path $policyPath -Name "InternetExplorerIntegrationReloadInIEModeAllowed" -PropertyType DWord -Value 1 -Force | Out-Null

# 3) Show/pin the "Reload in Internet Explorer mode" toolbar button
# InternetExplorerModeToolbarButtonEnabled: 1 = show/pin, 0 = don't
New-ItemProperty -Path $policyPath -Name "InternetExplorerModeToolbarButtonEnabled" -PropertyType DWord -Value 1 -Force | Out-Null

# Optional: If you want to enforce that IE mode is available for a site list (not required here),
# you would set InternetExplorerIntegrationSiteList to an XML path/URL. (We are not adding sites per your request.)

Write-Host "`nPolicies written to $policyPath :"
Get-ItemProperty -Path $policyPath | Select-Object InternetExplorerIntegrationLevel, InternetExplorerIntegrationReloadInIEModeAllowed, InternetExplorerModeToolbarButtonEnabled | Format-List

# Attempt to gracefully stop Edge so policy takes effect on next start
$edgeProcs = Get-Process -Name msedge -ErrorAction SilentlyContinue
if ($edgeProcs) {
    Write-Host "`nStopping running Microsoft Edge processes so changes take effect..."
    $edgeProcs | Stop-Process -Force
    Write-Host "Edge processes stopped."
} else {
    Write-Host "`nNo running msedge processes found."
}

Write-Host "`nDone. Please start Microsoft Edge and check:"
Write-Host "  - edge://settings/defaultBrowser  -> 'Allow sites to be reloaded in Internet Explorer mode' should be enabled (policy enforced)."
Write-Host "  - The IE-mode toolbar button should be visible (or available to show) in the toolbar / appearance settings."
