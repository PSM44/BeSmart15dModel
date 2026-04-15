[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Root = "C:\01. GitHub\BeSmart15dModel",

    [Parameter(Mandatory = $false)]
    [ValidateSet("Index","Core","All")]
    [string]$Mode = "All"
)

$ErrorActionPreference = "Stop"

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-RelativePathSafe {
    param(
        [Parameter(Mandatory = $true)][string]$BasePath,
        [Parameter(Mandatory = $true)][string]$TargetPath
    )
    try {
        $baseUri   = [System.Uri]((Resolve-Path -LiteralPath $BasePath).Path.TrimEnd('\') + '\')
        $targetUri = [System.Uri]((Resolve-Path -LiteralPath $TargetPath).Path)
        $relUri    = $baseUri.MakeRelativeUri($targetUri)
        return [System.Uri]::UnescapeDataString($relUri.ToString()).Replace('/', '\')
    }
    catch {
        return $TargetPath
    }
}

if (-not (Test-Path -LiteralPath $Root)) {
    throw "Root no existe: $Root"
}

$OutputDir = Join-Path $Root '07_RADAR\OUTPUTS'
Ensure-Directory -Path $OutputDir

$Stamp = (Get-Date).ToString('yyyyMMdd.HHmmss')
$IndexPath = Join-Path $OutputDir ("Radar.Index.{0}.txt" -f $Stamp)
$CorePath  = Join-Path $OutputDir ("Radar.Core.{0}.txt"  -f $Stamp)

$allFiles = Get-ChildItem -LiteralPath $Root -Recurse -File -Force |
    Where-Object {
        $_.FullName -notmatch '\\07_RADAR\\OUTPUTS\\'
    } |
    Sort-Object FullName

if ($Mode -in @('Index','All')) {
    $indexLines = New-Object System.Collections.Generic.List[string]
    $indexLines.Add("============================================================")
    $indexLines.Add("RADAR.INDEX")
    $indexLines.Add("============================================================")
    $indexLines.Add("")
    $indexLines.Add("ROOT: $Root")
    $indexLines.Add("GENERATED: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))")
    $indexLines.Add("TOTAL_FILES: $($allFiles.Count)")
    $indexLines.Add("")
    $indexLines.Add("RELATIVE_PATH`tEXT`tSIZE_BYTES`tLAST_WRITE")
    foreach ($f in $allFiles) {
        $rel = Get-RelativePathSafe -BasePath $Root -TargetPath $f.FullName
        $line = "{0}`t{1}`t{2}`t{3}" -f $rel, $f.Extension, $f.Length, $f.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
        $indexLines.Add($line)
    }
    Set-Content -LiteralPath $IndexPath -Value $indexLines -Encoding UTF8
}

if ($Mode -in @('Core','All')) {
    $coreLines = New-Object System.Collections.Generic.List[string]
    $coreLines.Add("============================================================")
    $coreLines.Add("RADAR.CORE")
    $coreLines.Add("============================================================")
    $coreLines.Add("")
    $coreLines.Add("ROOT: $Root")
    $coreLines.Add("GENERATED: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))")
    $coreLines.Add("")

    $includeExt = @('.txt', '.ps1', '.psm1', '.json', '.csv', '.md', '.yaml', '.yml', '.xml')
    $coreFiles = $allFiles | Where-Object { $includeExt -contains $_.Extension.ToLowerInvariant() }

    foreach ($f in $coreFiles) {
        $rel = Get-RelativePathSafe -BasePath $Root -TargetPath $f.FullName
        $coreLines.Add("------------------------------------------------------------")
        $coreLines.Add("FILE: $rel")
        $coreLines.Add("------------------------------------------------------------")
        $coreLines.Add("")
        try {
            $content = Get-Content -LiteralPath $f.FullName -ErrorAction Stop
            if ($null -eq $content -or $content.Count -eq 0) {
                $coreLines.Add("[EMPTY]")
            }
            else {
                foreach ($line in $content) {
                    $coreLines.Add([string]$line)
                }
            }
        }
        catch {
            $coreLines.Add("[UNREADABLE] $($_.Exception.Message)")
        }
        $coreLines.Add("")
    }

    Set-Content -LiteralPath $CorePath -Value $coreLines -Encoding UTF8
}

Write-Host "RADAR generado correctamente."
if (Test-Path -LiteralPath $IndexPath) { Write-Host "INDEX: $IndexPath" }
if (Test-Path -LiteralPath $CorePath)  { Write-Host "CORE : $CorePath"  }
