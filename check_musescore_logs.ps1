# Check MuseScore 4 logs for plugin errors
$logPath = "$env:LOCALAPPDATA\MuseScore\MuseScore4\logs"

if (Test-Path $logPath) {
    Write-Host "Checking MuseScore 4 logs in: $logPath" -ForegroundColor Green
    Write-Host ""
    
    $latestLog = Get-ChildItem $logPath -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    
    if ($latestLog) {
        Write-Host "Latest log file: $($latestLog.Name)" -ForegroundColor Yellow
        Write-Host "Last modified: $($latestLog.LastWriteTime)" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "=== Last 30 lines of log (looking for plugin errors) ===" -ForegroundColor Cyan
        Write-Host ""
        
        $content = Get-Content $latestLog.FullName -Tail 30
        $content | ForEach-Object {
            if ($_ -match "plugin|Plugin|NoteNameAbove|ERROR|error") {
                Write-Host $_ -ForegroundColor Red
            } else {
                Write-Host $_
            }
        }
    } else {
        Write-Host "No log files found in $logPath" -ForegroundColor Red
    }
} else {
    Write-Host "Log directory not found: $logPath" -ForegroundColor Red
}

Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

