Write-Host ""
Write-Host "This script inspects the subfolders of a user's profile directory on a remote machine" -ForegroundColor Cyan
Write-Host ""

# Prompt for input
$userInput = Read-Host "Enter the username (e.g. jsmith or DOMAIN\\jsmith)"
$computerName = Read-Host "Enter the target computer name"

# Normalize domain prefix if missing
if ($userInput -notmatch '\\') {
    $domain = $env:USERDOMAIN
    $userInput = "$domain\$userInput"
}

# Extract just the username for folder path
$usernameOnly = $userInput.Split('\')[-1]
$profilePath = "C:\Users\$usernameOnly"

Write-Host ""
Write-Host "Inspecting folder: $profilePath on $computerName..." -ForegroundColor Yellow

try {
$results = Invoke-Command -ComputerName $computerName -ScriptBlock {
    param($targetPath)

    if (-Not (Test-Path $targetPath)) {
        Write-Output "Profile path not found: $targetPath"
        return
    }

    $subfolders = Get-ChildItem -Path $targetPath -Directory -ErrorAction Stop

    foreach ($folder in $subfolders) {
        try {
            $items = Get-ChildItem -Path $folder.FullName -Recurse -Force -ErrorAction Stop |
                    Where-Object { $_.PSIsContainer -eq $false -and $_.Length -ne $null }

            $sizeBytes = ($items | Measure-Object -Property Length -Sum).Sum
            if ($sizeBytes -eq $null) { $sizeBytes = 0 }

            if ($sizeBytes -gt 1GB) {
                $sizeFormatted = "{0:N2} GB" -f ($sizeBytes / 1GB)
            } else {
                $sizeFormatted = "{0:N2} MB" -f ($sizeBytes / 1MB)
            }

            Write-Output "$($folder.Name): $sizeFormatted"
        } catch {
            Write-Output "$($folder.Name): [Error accessing folder]"
        }
    }
} -ArgumentList $profilePath -ErrorAction Stop
    foreach ($line in $results) {
        Write-Host $line -ForegroundColor Green
    }
} catch {
Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
}
