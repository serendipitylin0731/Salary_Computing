param(
    [switch]$SkipAppBuild
)

$ErrorActionPreference = "Stop"

$projectRoot = [IO.Path]::GetFullPath((Split-Path -Parent $MyInvocation.MyCommand.Path))
Set-Location -LiteralPath $projectRoot

if (-not $SkipAppBuild) {
    & (Join-Path $projectRoot "build_exe.ps1")
    if ($LASTEXITCODE -ne 0) {
        throw "Application build failed with exit code $LASTEXITCODE"
    }
}

$compilerCandidates = @(
    (Get-Command iscc.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue),
    (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"),
    (Join-Path $env:ProgramFiles "Inno Setup 6\ISCC.exe"),
    (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe")
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

$compiler = $compilerCandidates | Select-Object -First 1
if (-not $compiler) {
    throw "Inno Setup 6 was not found. Install JRSoftware.InnoSetup and retry."
}

$appExe = Join-Path $projectRoot "dist\SalaryComputing\SalaryComputing.exe"
if (-not (Test-Path -LiteralPath $appExe)) {
    throw "Missing application executable: $appExe"
}

& $compiler (Join-Path $projectRoot "installer.iss")
if ($LASTEXITCODE -ne 0) {
    throw "Installer build failed with exit code $LASTEXITCODE"
}

$installer = Join-Path $projectRoot "dist\SalaryComputing-Setup.exe"
if (-not (Test-Path -LiteralPath $installer)) {
    throw "Installer output was not created: $installer"
}

Write-Host "Installer build complete: $installer"
