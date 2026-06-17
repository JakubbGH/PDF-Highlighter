param(
    [string]$NodePath,
    [string]$PythonPath
)

$ErrorActionPreference = "Stop"
$ProjectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$PlaceholderVbaHash = "0CED1464B3677E98F5E3A8C5D80135E18DC98DCA39299F1A8CFD2A00999FBF9F"

function Resolve-Tool {
    param(
        [string]$ProvidedPath,
        [string]$CommandName,
        [string[]]$FallbackPaths
    )

    if (-not [string]::IsNullOrWhiteSpace($ProvidedPath)) {
        if (-not (Test-Path -LiteralPath $ProvidedPath)) {
            throw "$CommandName was provided but does not exist: $ProvidedPath"
        }
        return [System.IO.Path]::GetFullPath($ProvidedPath)
    }

    foreach ($fallback in $FallbackPaths) {
        if (Test-Path -LiteralPath $fallback) {
            return [System.IO.Path]::GetFullPath($fallback)
        }
    }

    $command = Get-Command $CommandName -ErrorAction SilentlyContinue
    if ($null -ne $command) {
        if ($command.Source -like "*\WindowsApps\python.exe") {
            throw "The Windows Store Python alias was found, but a real Python runtime was not. Pass -PythonPath or install Python."
        }
        return $command.Source
    }

    throw "$CommandName was not found. Pass -${CommandName}Path or install $CommandName."
}

function Invoke-Check {
    param(
        [string]$Name,
        [scriptblock]$Script
    )

    Write-Host "== $Name"
    & $Script
    Write-Host "OK: $Name"
}

function Assert-File {
    param([string]$RelativePath)

    $path = Join-Path $ProjectRoot $RelativePath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Required publish file missing: $RelativePath"
    }
}

$nodeFallback = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
$pythonFallback = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

$NodePath = Resolve-Tool $NodePath "node" @($nodeFallback)
$PythonPath = Resolve-Tool $PythonPath "python" @($pythonFallback)

Push-Location $ProjectRoot
try {
    Invoke-Check "publish files exist" {
        @(
            ".nojekyll",
            "index.html",
            "styles.css",
            "app.js",
            "assets\sample-floor.svg",
            "vendor\pdf.min.js",
            "vendor\pdf.worker.min.js",
            "vendor\README.md",
            "vendor\excel\vbaProject.bin",
            "vendor\excel\README.md",
            "vendor\excel\ThisWorkbookCode.bas",
            "vendor\excel\RefreshZoneColours.bas",
            "vendor\excel\ProgressSheetChange.bas",
            "tools\extract_vba_project.py",
            "tools\install_excel_macro_template.ps1",
            "tools\run_checks.ps1",
            "tools\smoke_test.js",
            "tools\browser_smoke_test.js",
            "README.md"
        ) | ForEach-Object { Assert-File $_ }
    }

    Invoke-Check "JavaScript syntax" {
        & $NodePath --check app.js
        & $NodePath --check tools\smoke_test.js
        & $NodePath --check tools\browser_smoke_test.js
    }

    Invoke-Check "PowerShell helper syntax" {
        $null = [scriptblock]::Create((Get-Content -LiteralPath "tools\install_excel_macro_template.ps1" -Raw))
        $null = [scriptblock]::Create((Get-Content -LiteralPath "tools\run_checks.ps1" -Raw))
    }

    Invoke-Check "Python extractor starts" {
        & $PythonPath tools\extract_vba_project.py --help | Out-Null
    }

    Invoke-Check "app smoke test" {
        & $NodePath tools\smoke_test.js
    }

    $vbaHash = (Get-FileHash -Algorithm SHA256 -LiteralPath "vendor\excel\vbaProject.bin").Hash
    if ($vbaHash -eq $PlaceholderVbaHash) {
        Write-Warning "vendor\excel\vbaProject.bin is still the placeholder sample macro. XLSM exports can still be installed with a local macro template in the browser, but repo-default exports will be snapshot-only until tools\install_excel_macro_template.ps1 is run on a machine with Excel."
    } else {
        Write-Host "OK: compiled VBA project is not the placeholder sample."
    }

    Write-Host "All local checks passed."
} finally {
    Pop-Location
}
