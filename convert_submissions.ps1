param(
    [Parameter(Mandatory = $true)]
    [string]$InputDir,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string]$Template = "",
    [string]$Pattern = "*.zip",
    [string]$Report = "",
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"

if (-not $Template) {
    $Template = Join-Path (Split-Path -Parent $PSCommandPath) "template.docx"
}

$pythonArgs = @(
    (Join-Path (Split-Path -Parent $PSCommandPath) "batch_convert_archives.py"),
    "--input-dir", $InputDir,
    "--output-dir", $OutputDir,
    "--template", $Template,
    "--pattern", $Pattern
)

if ($Report) {
    $pythonArgs += @("--report", $Report)
}

if ($Overwrite) {
    $pythonArgs += "--overwrite"
}

python @pythonArgs
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
