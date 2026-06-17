param(
    [string]$SourceCodePath,
    [string]$TemplatePath,
    [string]$OutputPath,
    [switch]$NoBackup,
    [switch]$KeepTemplate
)

$ErrorActionPreference = "Stop"

function Resolve-ProjectPath {
    param(
        [string]$PathValue,
        [string]$DefaultRelativePath
    )

    $value = if ([string]::IsNullOrWhiteSpace($PathValue)) { $DefaultRelativePath } else { $PathValue }
    if ([System.IO.Path]::IsPathRooted($value)) {
        return [System.IO.Path]::GetFullPath($value)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $ProjectRoot $value))
}

function New-BackupPath {
    param([string]$PathValue)

    $candidate = "$PathValue.bak"
    $counter = 1
    while (Test-Path -LiteralPath $candidate) {
        $candidate = "$PathValue.bak$counter"
        $counter += 1
    }
    return $candidate
}

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

$ProjectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$SourceCodePath = Resolve-ProjectPath $SourceCodePath "vendor\excel\ThisWorkbookCode.bas"
$TemplatePath = Resolve-ProjectPath $TemplatePath "tmp\floor-plan-macro-template.xlsm"
$OutputPath = Resolve-ProjectPath $OutputPath "vendor\excel\vbaProject.bin"

if (-not (Test-Path -LiteralPath $SourceCodePath)) {
    throw "Macro source file not found: $SourceCodePath"
}

$templateDirectory = Split-Path -Parent $TemplatePath
$outputDirectory = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Force -Path $templateDirectory | Out-Null
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

$excel = $null
$workbook = $null
$worksheets = $null
$vbProject = $null
$thisWorkbookComponent = $null

try {
    Write-Host "Starting Excel..."
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $excel.Visible = $false

    $workbook = $excel.Workbooks.Add()

    try {
        $vbProject = $workbook.VBProject
        $thisWorkbookComponent = $vbProject.VBComponents.Item("ThisWorkbook")
    } catch {
        throw "Excel blocked access to the VBA project. Enable 'Trust access to the VBA project object model' in Excel Trust Center, then run this script again. Original error: $($_.Exception.Message)"
    }

    $codeModule = $thisWorkbookComponent.CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }

    $macroCode = Get-Content -LiteralPath $SourceCodePath -Raw
    $codeModule.AddFromString($macroCode)

    $worksheets = $workbook.Worksheets
    while ($worksheets.Count -lt 2) {
        [void]$worksheets.Add()
    }
    $worksheets.Item(1).Name = "Plan"
    $worksheets.Item(2).Name = "Progress"

    if (Test-Path -LiteralPath $TemplatePath) {
        Remove-Item -LiteralPath $TemplatePath -Force
    }

    Write-Host "Saving macro-enabled template..."
    $xlOpenXMLWorkbookMacroEnabled = 52
    $workbook.SaveAs($TemplatePath, $xlOpenXMLWorkbookMacroEnabled)
    $workbook.Close($false)
    $workbook = $null
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
    }
    if ($null -ne $excel) {
        $excel.Quit()
    }

    Release-ComObject $thisWorkbookComponent
    Release-ComObject $vbProject
    Release-ComObject $worksheets
    Release-ComObject $workbook
    Release-ComObject $excel
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = $null
$entryStream = $null
$outputStream = $null

try {
    $archive = [System.IO.Compression.ZipFile]::OpenRead($TemplatePath)
    $entry = $archive.Entries | Where-Object { $_.FullName -ieq "xl/vbaProject.bin" } | Select-Object -First 1
    if ($null -eq $entry) {
        throw "The generated template does not contain xl/vbaProject.bin."
    }
    if ($entry.Length -lt 1024) {
        throw "The generated vbaProject.bin is unexpectedly small."
    }

    if ((Test-Path -LiteralPath $OutputPath) -and -not $NoBackup) {
        $backupPath = New-BackupPath $OutputPath
        Copy-Item -LiteralPath $OutputPath -Destination $backupPath -Force
        Write-Host "Backed up existing VBA project to $backupPath"
    }

    $entryStream = $entry.Open()
    $outputStream = [System.IO.File]::Create($OutputPath)
    $entryStream.CopyTo($outputStream)
    Write-Host "Installed compiled VBA project to $OutputPath"
} finally {
    if ($null -ne $outputStream) { $outputStream.Dispose() }
    if ($null -ne $entryStream) { $entryStream.Dispose() }
    if ($null -ne $archive) { $archive.Dispose() }
}

if (-not $KeepTemplate) {
    Remove-Item -LiteralPath $TemplatePath -Force
} else {
    Write-Host "Kept macro template at $TemplatePath"
}

Write-Host "Done. Refresh the website and export a new XLSM."
