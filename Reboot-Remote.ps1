param (
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ComputerNames
)

# Help message
if ($ComputerNames -contains "/h" -or $ComputerNames -contains "/help") {
    Write-Host ""
    Write-Host "USAGE" -ForegroundColor Cyan
    Write-Host "This command reboots remote hosts" -ForegroundColor White
    Write-Host "Requires 1 to 5 hostnames as arguments" -ForegroundColor White
    Write-Host "Example" -ForegroundColor White
    Write-Host ".\Reboot-Remote.ps1 tlab-r11pc1 tlab-r11pc2" -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

# Validate number of hostnames
if ($ComputerNames.Count -lt 1 -or $ComputerNames.Count -gt 5) {
    Write-Host "Please provide between 1 and 5 hostnames. Use /h for help" -ForegroundColor Red
    exit 1
}

# Display warning
Write-Host ""
Write-Host "WARNING: Are you sure you want to reboot these machines remotely!" -ForegroundColor White -BackgroundColor Red
Write-Host "Hostnames: $($ComputerNames -join ', ')" -ForegroundColor Yellow
$confirmation = Read-Host "Type Y to continue or N to cancel"

if ($confirmation -notin @("Y", "y")) {
    Write-Host "Operation cancelled" -ForegroundColor Magenta
    exit 0
}

# Reboot logic with timeout
foreach ($computer in $ComputerNames) {
    Write-Host ""
    Write-Host "Attempting to restart $computer" -ForegroundColor Cyan

    $job = Start-Job -ScriptBlock {
        param($comp)
        Restart-Computer -ComputerName $comp -Force -ErrorAction Stop
    } -ArgumentList $computer

    $completed = Wait-Job -Job $job -Timeout 5

    if ($completed) {
        try {
            Receive-Job -Job $job -ErrorAction Stop
            Write-Host "Restart triggered on $computer" -ForegroundColor Green
        } catch {
            Write-Host "Failed to restart $computer ${_}" -ForegroundColor Red
        }
    } else {
        Write-Host "Timeout $computer may be offline or unresponsive" -ForegroundColor DarkYellow
    }

    Remove-Job -Job $job -Force
}