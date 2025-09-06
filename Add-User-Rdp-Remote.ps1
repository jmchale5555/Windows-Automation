Write-Host ""
Write-Host "This script will add the given user to the Remote Desktop Users group on the specified host" -ForegroundColor Cyan
Write-Host ""

# Prompt for input
$userName = Read-Host "Enter the username (e.g. jsmith or DOMAIN\jsmith)"
$computerName = Read-Host "Enter the target computer name"

# Normalize domain prefix if missing
if ($userName -notmatch '\\') {
    $domain = $env:USERDOMAIN
    $userName = "$domain\$userName"
}

Write-Host ""
Write-Host "Attempting to add $userName to Remote Desktop Users on $computerName..." -ForegroundColor Yellow

# Start job with timeout
$job = Start-Job -ScriptBlock {
    param($comp, $user)
    Invoke-Command -ComputerName $comp -ScriptBlock {
        param($uName)
        Add-LocalGroupMember -Group "Remote Desktop Users" -Member $uName
    } -ArgumentList $user -ErrorAction Stop
} -ArgumentList $computerName, $userName

$completed = Wait-Job -Job $job -Timeout 5

if ($completed) {
    try {
        Receive-Job -Job $job -ErrorAction Stop
        Write-Host "Success: ${userName} was added to Remote Desktop Users on $computerName" -ForegroundColor Green
    } catch {
        Write-Host "Error: Failed to add user on ${computerName} - ${_}" -ForegroundColor Red
    }
} else {
    Write-Host "Timeout: ${computerName} may be offline or unreachable" -ForegroundColor DarkYellow
}

Remove-Job -Job $job -Force