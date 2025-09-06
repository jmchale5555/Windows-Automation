param (
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ComputerNames
)

function Show-Help {
    Write-Host "`nUSAGE:" -ForegroundColor Cyan
    Write-Host "  This command updates group policy on remote hosts." -ForegroundColor White
    Write-Host "  Requires 1 to 5 hostnames as arguments." -ForegroundColor White
    Write-Host "  Example:" -ForegroundColor White
    Write-Host "    .\Update-GP-remote.ps1 hostname1 hostname2`n" -ForegroundColor Yellow
}

# Show help if /h or /help is passed, or if no arguments are provided
if (!$ComputerNames -or $ComputerNames -contains "/h" -or $ComputerNames -contains "/help") {
    Show-Help
    exit 0
}

# Validate number of hostnames
if ($ComputerNames.Count -lt 1 -or $ComputerNames.Count -gt 5) {
    Write-Error "Please provide between 1 and 5 hostnames. Use /h for help."
    exit 1
}

foreach ($computer in $ComputerNames) {
    Write-Host "Invoking gpupdate /force on $computer..." -ForegroundColor Cyan

    $scriptblock = {
        gpupdate /force
    }

    try {
        Invoke-Command -ComputerName $computer -ScriptBlock $scriptblock -ErrorAction Stop
        Write-Host "Successfully updated group policy on $computer." -ForegroundColor Green
    } catch {
        Write-Host "Failed to update group policy on $computer ${_}" -ForegroundColor Red
    }
}
