# Whitelist Visual Studio Code in Windows Firewall and Windows Defender
# Run this script as Administrator

# Self-elevate if not running as administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Not running as administrator. Attempting to relaunch with elevated privileges..."
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $MyInvocation.MyCommand.Definition + '"'
    $psi.Verb = 'runas'
    try {
        [System.Diagnostics.Process]::Start($psi) | Out-Null
    } catch {
        Write-Host "Elevation cancelled or failed. Exiting."
        exit 1
    }
    exit 0
}

# Path to Code.exe (user-specified)
$codeExe = "C:\Users\kylem\AppData\Local\Programs\Microsoft VS Code\Code.exe"

function Show-ErrorAndPause($msg) {
    Write-Host "[ERROR] $msg" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Check for admin rights
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Show-ErrorAndPause "Script must be run as Administrator. Right-click PowerShell and choose 'Run as administrator.'"
}

if (-not (Test-Path $codeExe)) {
    Show-ErrorAndPause "Could not find Code.exe at $codeExe. Please check the path or reinstall Visual Studio Code."
}

Write-Host "Visual Studio Code found at: $codeExe"

# Add VS Code to Windows Firewall (allow inbound and outbound)
Write-Host "Adding Visual Studio Code to Windows Firewall..."

try {
    # Remove existing rules by DisplayName
    $existingRules = Get-NetFirewallRule -DisplayName "Allow Visual Studio Code" -ErrorAction SilentlyContinue
    if ($existingRules) {
        Remove-NetFirewallRule -DisplayName "Allow Visual Studio Code"
    }
    # Add new rules
    New-NetFirewallRule -DisplayName "Allow Visual Studio Code" -Direction Inbound -Program $codeExe -Action Allow -Profile Any -Enabled True | Out-Null
    New-NetFirewallRule -DisplayName "Allow Visual Studio Code" -Direction Outbound -Program $codeExe -Action Allow -Profile Any -Enabled True | Out-Null
} catch {
    Show-ErrorAndPause "Failed to add firewall rules. Possible reasons: insufficient privileges, firewall service not running, or syntax error. $_"
}

Write-Host "VS Code whitelisted in Windows Firewall."

# Add exclusion to Windows Defender
Write-Host "Adding exclusion for Visual Studio Code in Windows Defender..."
try {
    Add-MpPreference -ExclusionProcess $codeExe
} catch {
    Show-ErrorAndPause "Failed to add Defender exclusion. Possible reasons: insufficient privileges, Windows Defender not installed, or syntax error. $_"
}

Write-Host "[SUCCESS] Visual Studio Code is now whitelisted in Windows Firewall and excluded from Windows Defender scans." -ForegroundColor Green

# Optionally reset network stack
Write-Host "\n---"
Write-Host "Resetting network stack to resolve potential agent communication issues..."
try {
    ipconfig /flushdns | Out-Null
    netsh int ip reset | Out-Null
    Write-Host "Network stack reset complete. It is recommended to restart your computer for changes to take full effect." -ForegroundColor Yellow
} catch {
    Write-Host "[WARNING] Failed to reset network stack. You may need to run the following commands manually as Administrator:" -ForegroundColor Yellow
    Write-Host "    ipconfig /flushdns"
    Write-Host "    netsh int ip reset"
}

Read-Host "Press Enter to exit" # Keeps window open
