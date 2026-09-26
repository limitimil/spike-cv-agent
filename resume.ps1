# One entry point from any directory. Python environment stays beside the repo.
$ErrorActionPreference = 'Stop'
$resumePython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $resumePython)) {
    Write-Error 'Run setup first: Python 3.10+ -m venv .venv, then .venv\Scripts\python -m pip install -r requirements.txt. See README.md.'
    exit 1
}
Push-Location $PSScriptRoot
try {
    if ($args.Count -eq 0) { & $resumePython -m src start }
    else { & $resumePython -m src @args }
    $resumeExitCode = $LASTEXITCODE
} finally { Pop-Location }
exit $resumeExitCode
