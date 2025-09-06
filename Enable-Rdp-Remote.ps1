param (
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ComputerNames
)

# Help message
if ($ComputerNames -contains "/h" -or $ComputerNames -contains "/help") {
    Write-Host ""
    Write-Host "USAGE" -ForegroundColor Cyan
    Write-Host "This command enables Remote Desktop on remote hosts" -ForegroundColor White
    Write-Host "Requires 1 to 5 hostnames as arguments" -ForegroundColor White
    Write-Host "Example" -ForegroundColor White
    Write-Host ".\Enable-RDP.ps1 tlab-r11pc1 tlab-r11pc2" -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

# Validate number of hostnames
if ($ComputerNames.Count -lt 1 -or $ComputerNames.Count -gt 5) {
    Write-Host "Please provide between 1 and 5 hostnames. Use /h for help" -ForegroundColor Red
    exit 1
}

foreach ($computer in $ComputerNames) {
    Write-Host ""
    Write-Host "Configuring Remote Desktop on $computer..." -ForegroundColor Cyan

    try {
        Invoke-Command -ComputerName $computer -ScriptBlock {
            # Enable Remote Desktop
            Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name 'fDenyTSConnections' -Value 0

            # Enable firewall rule for Remote Desktop
            Enable-NetFirewallRule -DisplayGroup "Remote Desktop"

            Write-Host "Remote Desktop enabled on $env:COMPUTERNAME" -ForegroundColor Green
        } -ErrorAction Stop
    } catch {
        Write-Host "Failed to configure Remote Desktop on $computer ${_}" -ForegroundColor Red
    }
}