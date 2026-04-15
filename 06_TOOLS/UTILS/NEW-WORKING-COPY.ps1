[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceFile,

    [Parameter(Mandatory = $false)]
    [string]$ProjectRoot = "C:\01. GitHub\BeSmart15dModel"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $SourceFile)) {
    throw "SourceFile no existe: $SourceFile"
}

$srcItem = Get-Item -LiteralPath $SourceFile
$destDir = Join-Path $ProjectRoot '01_RAW\WORKING_COPY'
if (-not (Test-Path -LiteralPath $destDir)) {
    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
}

$stamp = Get-Date -Format 'yyyyMMdd.HHmmss'
$destFile = Join-Path $destDir ("{0}.{1}{2}" -f $srcItem.BaseName, $stamp, $srcItem.Extension)

Copy-Item -LiteralPath $SourceFile -Destination $destFile -Force
Write-Host "WORKING COPY creada:"
Write-Host $destFile
