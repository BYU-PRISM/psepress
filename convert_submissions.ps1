param(
    [Parameter(Mandatory = $true)]
    [string]$InputDir,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string]$Template = "",
    [string]$Pattern = "*.zip",
    [string]$Report = "",
    [string]$Conference = "escape-37-2027",
    [string]$ConferenceName = "",
    [string]$ConferenceLocation = "",
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
    "--pattern", $Pattern,
    "--conference", $Conference
)

if ($Report) {
    $pythonArgs += @("--report", $Report)
}

if ($ConferenceName) {
    $pythonArgs += @("--conference-name", $ConferenceName)
}

if ($ConferenceLocation) {
    $pythonArgs += @("--conference-location", $ConferenceLocation)
}

if ($Overwrite) {
    $pythonArgs += "--overwrite"
}

python @pythonArgs
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
