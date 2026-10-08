$ErrorActionPreference = "Stop"

$projectRoot = [IO.Path]::GetFullPath((Split-Path -Parent $MyInvocation.MyCommand.Path))
Set-Location -LiteralPath $projectRoot

function Remove-GeneratedDirectory([string]$leafName) {
    $target = [IO.Path]::GetFullPath((Join-Path $projectRoot $leafName))
    $parent = [IO.Path]::GetDirectoryName($target)
    if ($parent -ne $projectRoot -or $leafName -notin @("build", "dist")) {
        throw "Refusing to remove an unexpected directory: $target"
    }
    if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
}

function Remove-PythonCaches {
    foreach ($relativePath in @("__pycache__", "src\__pycache__", "tests\__pycache__")) {
        $target = [IO.Path]::GetFullPath((Join-Path $projectRoot $relativePath))
        if (-not $target.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar)) {
            throw "Refusing to remove a cache outside the project: $target"
        }
        if (Test-Path -LiteralPath $target) {
            Remove-Item -LiteralPath $target -Recurse -Force
        }
    }
}

function Remove-RootSpecFiles {
    Get-ChildItem -LiteralPath $projectRoot -Filter "*.spec" -File | ForEach-Object {
        if ($_.DirectoryName -ne $projectRoot) {
            throw "Refusing to remove a spec outside the project: $($_.FullName)"
        }
        Remove-Item -LiteralPath $_.FullName -Force
    }
}

foreach ($requiredPath in @("main.py", "requirements.txt", "salary.json", "USER_GUIDE.md")) {
    if (-not (Test-Path -LiteralPath (Join-Path $projectRoot $requiredPath))) {
        throw "Missing required file: $requiredPath"
    }
}

python -c "import openpyxl, PyInstaller" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Installing build dependencies..."
    python -m pip install -r requirements.txt pyinstaller
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to install build dependencies."
    }
}

Write-Host "Running tests..."
python -m unittest discover -s tests
if ($LASTEXITCODE -ne 0) {
    throw "Tests failed; build cancelled."
}

Write-Host "Cleaning old build output..."
Remove-GeneratedDirectory "build"
Remove-GeneratedDirectory "dist"
Remove-RootSpecFiles
Remove-PythonCaches

$appName = "SalaryComputing"
Write-Host "Building $appName.exe..."
python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onedir `
    --specpath build `
    --name $appName `
    main.py

if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE"
}

$outputRoot = Join-Path $projectRoot "dist\$appName"

Copy-Item -LiteralPath (Join-Path $projectRoot "salary.json") -Destination $outputRoot -Force
Copy-Item -LiteralPath (Join-Path $projectRoot "USER_GUIDE.md") `
    -Destination (Join-Path $outputRoot "README.md") `
    -Force

foreach ($forbiddenPath in @("example", "docs\module.xlsx")) {
    if (Test-Path -LiteralPath (Join-Path $outputRoot $forbiddenPath)) {
        throw "Forbidden package content found: $forbiddenPath"
    }
}
$packagedWorkbooks = @(
    Get-ChildItem -LiteralPath $outputRoot -Recurse -File |
        Where-Object { $_.Extension.ToLowerInvariant() -in @(".xlsx", ".xlsm", ".xls") }
)
if ($packagedWorkbooks.Count -gt 0) {
    throw "Workbook files must not be included in the package."
}

Remove-GeneratedDirectory "build"
Remove-RootSpecFiles
Remove-PythonCaches

Write-Host "Build complete: $(Join-Path $outputRoot ($appName + '.exe'))"
