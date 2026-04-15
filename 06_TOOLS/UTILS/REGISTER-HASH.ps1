[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FilePath,

    [Parameter(Mandatory = $false)]
    [string]$ProjectRoot = "C:\01. GitHub\BeSmart15dModel"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $FilePath)) {
    throw "Archivo no existe: $FilePath"
}

$hash = Get-FileHash -LiteralPath $FilePath -Algorithm SHA256
$item = Get-Item -LiteralPath $FilePath

$csvPath = Join-Path $ProjectRoot '02_AUDIT\HASH\HASH.REGISTRY.csv'
if (-not (Test-Path -LiteralPath $csvPath)) {
    'timestamp_utc,file_name,file_path,size_bytes,last_write_time_utc,sha256,status,notes' | Set-Content -LiteralPath $csvPath -Encoding UTF8
}

$row = [pscustomobject]@{
    timestamp_utc        = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
    file_name            = $item.Name
    file_path            = $item.FullName
    size_bytes           = $item.Length
    last_write_time_utc  = $item.LastWriteTimeUtc.ToString('yyyy-MM-dd HH:mm:ss')
    sha256               = $hash.Hash
    status               = 'REGISTERED'
    notes                = ''
}

$row | Export-Csv -LiteralPath $csvPath -Append -NoTypeInformation -Encoding UTF8
Write-Host "HASH registrado en: $csvPath"
Write-Host "SHA256: $($hash.Hash)"
