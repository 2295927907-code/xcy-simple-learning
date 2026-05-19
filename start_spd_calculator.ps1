$ErrorActionPreference = "Stop"

$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $appDir "spectral_spd_gui.py"
$logPath = Join-Path $appDir "SPD_startup_log.txt"
$codexPython = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

function Write-LaunchLog {
    param([string]$Message)
    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -LiteralPath $logPath -Encoding UTF8 -Value "[$time] $Message"
}

try {
    if (Test-Path -LiteralPath $logPath) {
        Remove-Item -LiteralPath $logPath -Force
    }

    Write-LaunchLog "App dir: $appDir"
    Write-LaunchLog "Script: $scriptPath"

    if (-not (Test-Path -LiteralPath $scriptPath)) {
        throw "Cannot find spectral_spd_gui.py"
    }

    $pythonPath = $null
    if (Test-Path -LiteralPath $codexPython) {
        $pythonPath = $codexPython
        Write-LaunchLog "Python: bundled runtime"
    } else {
        $cmd = Get-Command python -ErrorAction SilentlyContinue
        if ($cmd) {
            $pythonPath = $cmd.Source
            Write-LaunchLog "Python: system runtime"
        }
    }

    if (-not $pythonPath) {
        throw "Python was not found."
    }

    $env:PYTHONUTF8 = "1"
    $argsForPython = @($scriptPath)
    if ($env:SPD_SELF_TEST -eq "1") {
        $argsForPython += "--self-test"
        Write-LaunchLog "Mode: self-test"
    } elseif ($env:SPD_GUI_SMOKE -eq "1") {
        Write-LaunchLog "Mode: GUI smoke"
    } else {
        Write-LaunchLog "Mode: GUI"
    }

    & $pythonPath @argsForPython 2>&1 | ForEach-Object {
        Write-LaunchLog $_.ToString()
    }

    $exitCode = $LASTEXITCODE
    if ($null -ne $exitCode -and $exitCode -ne 0) {
        throw "Exit code: $exitCode"
    }
} catch {
    Write-LaunchLog "FAILED: $($_.Exception.Message)"
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show(
        "Spectral SPD Calculator failed to start.`n`nPlease check SPD_startup_log.txt in this folder.",
        "Startup failed",
        "OK",
        "Error"
    ) | Out-Null
    exit 1
}
