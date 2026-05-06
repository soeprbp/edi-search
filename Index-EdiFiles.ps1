# Index-EdiFiles.ps1
# Standalone script to index EDI files to JSON with resumable progress
# Supports incremental saves and cancellation
param(
    [string]$SharePath = "\\svwpefs\WelchEncoreShare\WELCHPKG\EDI",
    [string]$OutputPath = "",
    [int]$MaxContentSize = 50000,
    [int]$BatchSize = 500,
    [int]$SaveIntervalMs = 10000,
    [switch]$ForceRebuild
)

$ErrorActionPreference = "Continue"

if ([string]::IsNullOrEmpty($OutputPath)) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $OutputPath = Join-Path $scriptDir "data\edi-index.json"
}

$dbDir = Split-Path -Parent $OutputPath
if (-not (Test-Path $dbDir)) {
    New-Item -ItemType Directory -Path $dbDir -Force | Out-Null | Out-Null
}

Write-Host "EDI File Indexer" -ForegroundColor Cyan
Write-Host "Source: $SharePath" -ForegroundColor Gray
Write-Host "Output: $OutputPath" -ForegroundColor Gray

if (-not (Test-Path $SharePath)) {
    Write-Error "Share path not found: $SharePath"
    exit 1
}

$startTime = Get-Date

Write-Host "Scanning share..." -ForegroundColor Cyan
$allFiles = @(Get-ChildItem -LiteralPath $SharePath -Recurse -File -ErrorAction SilentlyContinue)
Write-Host "Found $($allFiles.Count) files" -ForegroundColor Gray

$existingMap = @{}
$index = @{
    SharePath = $SharePath
    GeneratedUtc = $startTime.ToUniversalTime().ToString("o")
    FileCount = 0
    Items = @()
}

if (-not $ForceRebuild -and (Test-Path $OutputPath)) {
    try {
        $existing = Get-Content $OutputPath -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
        if ($existing.Items) {
            foreach ($item in $existing.Items) {
                $existingMap[$item.FullPath] = $item
            }
            $index = $existing
            Write-Host "Loaded existing index with $($index.Items.Count) files" -ForegroundColor Gray
        }
    } catch {
        Write-Warning "Could not load existing index, will rebuild: $_"
    }
}

$newFiles = @()
$unchangedCount = 0
$fileCount = $allFiles.Count

Write-Host "Checking for changes..." -ForegroundColor Cyan
for ($i = 0; $i -lt $fileCount; $i++) {
    $file = $allFiles[$i]
    $hash = "$($file.Length)-$($file.LastWriteTimeUtc.ToString('o'))"
    
    if ($existingMap.ContainsKey($file.FullName)) {
        $existing = $existingMap[$file.FullName]
        if ($existing.ContentHash -eq $hash) {
            $unchangedCount++
            continue
        }
    }
    
    $newFiles += $file
    
    if ($i % 1000 -eq 0 -and $i -gt 0) {
        Write-Host "  Scanned $i / $fileCount files, $($newFiles.Count) need processing..." -ForegroundColor Gray
    }
}

Write-Host "Unchanged: $unchangedCount" -ForegroundColor Gray
Write-Host "Need to process: $($newFiles.Count) files" -ForegroundColor Yellow

if ($newFiles.Count -eq 0) {
    Write-Host "Index is up to date!" -ForegroundColor Green
    $index.FileCount = $allFiles.Count
    $index.GeneratedUtc = (Get-Date).ToUniversalTime().ToString("o")
    $index | ConvertTo-Json -Depth 5 | Set-Content $OutputPath -Encoding UTF8
    Write-Host "Index saved to: $OutputPath" -ForegroundColor Green
    exit 0
}

$throttleLimit = [Math]::Max(4, [Environment]::ProcessorCount)
$isPS7 = $PSVersionTable.PSVersion.Major -ge 7

Write-Host "Indexing $($newFiles.Count) files..." -ForegroundColor Cyan

$allItems = @($index.Items)
$allItemsMap = @{}
foreach ($item in $allItems) {
    $allItemsMap[$item.FullPath] = $item
}

$processedCount = 0
$lastSaveTime = Get-Date

foreach ($file in $newFiles) {
    $content = ""
    try {
        $content = Get-Content -LiteralPath $file.FullName -Raw -TotalCount $MaxContentSize -ErrorAction SilentlyContinue
    } catch {
        $content = ""
    }
    
    $relativePath = if ($file.FullName.StartsWith($SharePath, [System.StringComparison]::OrdinalIgnoreCase)) {
        $file.FullName.Substring($SharePath.Length).TrimStart("\")
    } else {
        $file.FullName
    }
    
    $hash = "$($file.Length)-$($file.LastWriteTimeUtc.ToString('o'))"
    
    $item = @{
        Name = $file.Name
        FullPath = $file.FullName
        RelativePath = $relativePath
        Directory = $file.DirectoryName
        Extension = $file.Extension
        Length = [int64]$file.Length
        LastWriteUtc = $file.LastWriteTimeUtc.ToString("o")
        Content = $content
        ContentHash = $hash
    }
    
    $allItemsMap[$file.FullName] = $item
    $processedCount++
    
    if ($processedCount % 100 -eq 0) {
        $elapsed = (Get-Date) - $startTime
        $rate = if ($elapsed.TotalSeconds -gt 0) { [int]($processedCount / $elapsed.TotalSeconds) } else { 0 }
        Write-Host "  Processed $processedCount / $($newFiles.Count) ($rate files/sec)" -ForegroundColor Gray
    }
    
    $now = Get-Date
    if (($now - $lastSaveTime).TotalMilliseconds -gt $SaveIntervalMs) {
        $allItems = @($allItemsMap.Values)
        $index.Items = $allItems
        $index.FileCount = $allItems.Count
        $index.GeneratedUtc = $now.ToUniversalTime().ToString("o")
        
        $tempPath = $OutputPath + ".partial"
        $index | ConvertTo-Json -Depth 5 | Set-Content $tempPath -Encoding UTF8
        Write-Host "  Checkpoint saved ($($allItems.Count) total files)" -ForegroundColor Gray
        $lastSaveTime = $now
    }
}

$finalItems = @($allItemsMap.Values)
$index.Items = $finalItems
$index.FileCount = $finalItems.Count
$index.GeneratedUtc = (Get-Date).ToUniversalTime().ToString("o")

$totalElapsed = (Get-Date) - $startTime
$rate = if ($totalElapsed.TotalSeconds -gt 0) { [int]($processedCount / $totalElapsed.TotalSeconds) } else { 0 }

Write-Host "Indexed $($newFiles.Count) new/updated files" -ForegroundColor Green
Write-Host "Total files: $($index.FileCount)" -ForegroundColor Gray
Write-Host "Time: $([math]::Round($totalElapsed.TotalSeconds, 1))s ($rate files/sec)" -ForegroundColor Gray

if (Test-Path $OutputPath) {
    Remove-Item $OutputPath -Force -ErrorAction SilentlyContinue
}
if (Test-Path ($OutputPath + ".partial")) {
    Remove-Item ($OutputPath + ".partial") -Force -ErrorAction SilentlyContinue
}

$index | ConvertTo-Json -Depth 5 | Set-Content $OutputPath -Encoding UTF8
Write-Host "Index saved to: $OutputPath" -ForegroundColor Green