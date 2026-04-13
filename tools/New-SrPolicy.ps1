param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$PolicyPepper = "SheetRendererAuth:v1"

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $scriptRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    }
    elseif ($MyInvocation.MyCommand.Path) {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    else {
        (Get-Location).Path
    }

    $OutputPath = Join-Path $scriptRoot "..\SheetRenderer\sr_policy.dat"
}

function Get-Sha256Hex {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $hashBytes = $sha256.ComputeHash($bytes)
        return -join ($hashBytes | ForEach-Object { $_.ToString("x2") })
    }
    finally {
        $sha256.Dispose()
    }
}

function New-PolicyLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserName,

        [string]$ExpireDateText = ""
    )

    $normalizedExpireDate = if ($null -eq $ExpireDateText) { "" } else { $ExpireDateText.Trim() }
    if ($normalizedExpireDate) {
        [void][DateTime]::ParseExact(
            $normalizedExpireDate,
            "yyyy-MM-dd",
            [System.Globalization.CultureInfo]::InvariantCulture
        )
    }

    $payload = "$PolicyPepper|$UserName|$normalizedExpireDate"
    $hash = Get-Sha256Hex -Text $payload
    return "$hash|$normalizedExpireDate"
}

function Compress-ToBase64 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $inputBytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $memoryStream = New-Object System.IO.MemoryStream
    try {
        $gzipStream = New-Object System.IO.Compression.GZipStream($memoryStream, [System.IO.Compression.CompressionMode]::Compress, $true)
        try {
            $gzipStream.Write($inputBytes, 0, $inputBytes.Length)
        }
        finally {
            $gzipStream.Dispose()
        }

        return [Convert]::ToBase64String($memoryStream.ToArray())
    }
    finally {
        $memoryStream.Dispose()
    }
}

$resolvedInputPath = (Resolve-Path $InputPath).Path
$lines = Get-Content $resolvedInputPath -Encoding UTF8

$outputLines = New-Object System.Collections.Generic.List[string]
foreach ($rawLine in $lines) {
    $line = if ($null -eq $rawLine) { "" } else { $rawLine.Trim() }
    if (-not $line -or $line.StartsWith("#")) {
        continue
    }

    $parts = $line.Split(",", 2)
    $userName = $parts[0].Trim()
    $expireDateText = if ($parts.Length -gt 1) { $parts[1].Trim() } else { "" }

    if (-not $userName) {
        continue
    }

    $outputLines.Add((New-PolicyLine -UserName $userName -ExpireDateText $expireDateText))
}

$payloadText = ($outputLines -join [Environment]::NewLine)
$encodedPayload = Compress-ToBase64 -Text $payloadText

$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = Split-Path $resolvedOutputPath -Parent
if (-not (Test-Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

[System.IO.File]::WriteAllText($resolvedOutputPath, $encodedPayload, [System.Text.Encoding]::UTF8)
Write-Host "Created: $resolvedOutputPath"
