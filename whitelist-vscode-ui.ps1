Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Minimal, professional Windows Forms UI for VS Code whitelisting
$form = New-Object System.Windows.Forms.Form
$form.Text = 'VS Code Whitelist Utility'
$form.Size = New-Object System.Drawing.Size(540,380)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = 'Run Script'
$runButton.Location = New-Object System.Drawing.Point(20,20)
$runButton.Size = New-Object System.Drawing.Size(100,35)

$statusBox = New-Object System.Windows.Forms.RichTextBox
$statusBox.Location = New-Object System.Drawing.Point(20,70)
$statusBox.Size = New-Object System.Drawing.Size(490,250)
$statusBox.ReadOnly = $true
$statusBox.BackColor = 'White'
$statusBox.Font = New-Object System.Drawing.Font('Consolas',10)

$form.Controls.Add($runButton)
$form.Controls.Add($statusBox)

function Write-Log {
    param($msg, $color)
    $statusBox.SelectionStart = $statusBox.TextLength
    $statusBox.SelectionColor = [System.Drawing.Color]::$color
    $statusBox.AppendText((Get-Date -f 'HH:mm:ss')+': '+$msg+"`n")
    $statusBox.SelectionColor = $statusBox.ForeColor
    $statusBox.ScrollToCaret()
}

function Show-Restart {
    [System.Windows.Forms.MessageBox]::Show('Network stack reset complete. Please restart your computer.','Restart Recommended','OK','Information')
}

$runButton.Add_Click({
    $runButton.Enabled = $false
    $statusBox.Clear()
    try {
        Write-Log 'Starting VS Code whitelisting...' 'Black'
        # Detect VS Code
        $codeExe = "$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe"
        if (-not (Test-Path $codeExe)) {
            Write-Log 'VS Code not found at default location. Prompting for path...' 'Orange'
            $ofd = New-Object System.Windows.Forms.OpenFileDialog
            $ofd.Title = 'Locate Code.exe'
            $ofd.Filter = 'Code.exe|Code.exe'
            if ($ofd.ShowDialog() -eq 'OK') {
                $codeExe = $ofd.FileName
            } else {
                Write-Log 'User cancelled. Exiting.' 'Red'
                return
            }
        }
        Write-Log "VS Code found at: $codeExe" 'Green'

        # Firewall rules
        try {
            $existing = Get-NetFirewallRule -DisplayName 'Allow Visual Studio Code' -ErrorAction SilentlyContinue
            if ($existing) { Remove-NetFirewallRule -DisplayName 'Allow Visual Studio Code' }
            New-NetFirewallRule -DisplayName 'Allow Visual Studio Code' -Direction Inbound -Program $codeExe -Action Allow -Profile Any -Enabled True | Out-Null
            New-NetFirewallRule -DisplayName 'Allow Visual Studio Code' -Direction Outbound -Program $codeExe -Action Allow -Profile Any -Enabled True | Out-Null
            Write-Log 'Firewall rules set.' 'Green'
        } catch {
            Write-Log "Failed to set firewall rules: $_" 'Red'
        }

        # Defender exclusion
        try {
            if (Get-Command Add-MpPreference -ErrorAction SilentlyContinue) {
                Add-MpPreference -ExclusionProcess $codeExe
                Write-Log 'Defender exclusion added.' 'Green'
            } else {
                Write-Log 'Windows Defender module not found. Skipping Defender exclusion.' 'Orange'
            }
        } catch {
            Write-Log "Failed to add Defender exclusion: $_" 'Red'
        }

        # Network stack reset
        try {
            ipconfig /flushdns | Out-Null
            netsh int ip reset | Out-Null
            Write-Log 'Network stack reset.' 'Orange'
            Show-Restart
        } catch {
            Write-Log 'Failed to reset network stack. Run manually if needed.' 'Red'
        }

        Write-Log 'All done.' 'Green'
    } catch {
        Write-Log "Unexpected error: $_" 'Red'
    } finally {
        $runButton.Enabled = $true
    }
})

[void]$form.ShowDialog()
