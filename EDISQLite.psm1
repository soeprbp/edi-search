# EDISQLite.psm1
# SQLite-based EDI file index for on-demand searching
# Uses native sqlite3.dll from iTunes

$script:SQLiteDbPath = $null
$script:Connection = $null
$script:SQLiteAssembly = $null

function Initialize-EDIDatabase {
    param(
        [Parameter(Mandatory=$true)]
        [string]$DatabasePath,
        
        [switch]$ForceNew
    )
    
    $script:SQLiteDbPath = $DatabasePath
    $dbDir = Split-Path -Parent $DatabasePath
    if (-not (Test-Path $dbDir)) {
        New-Item -ItemType Directory -Path $dbDir -Force | Out-Null
    }
    
    if ($ForceNew -and (Test-Path $DatabasePath)) {
        Remove-Item $DatabasePath -Force
    }
    
    $nativeDll = "C:\Program Files\iTunes\SQLite3.dll"
    if (-not (Test-Path $nativeDll)) {
        $nativeDll = "C:\Program Files\Fortinet\FortiClient\sqlite3.dll"
    }
    
    $dllPaths = @(
        "C:\Program Files\iTunes\SQLite3.dll",
        "C:\Program Files\Fortinet\FortiClient\sqlite3.dll"
    )
    
    foreach ($dllPath in $dllPaths) {
        if (Test-Path $dllPath) {
            $env:Path += ";" + (Split-Path $dllPath)
            Write-Host "Using SQLite from: $dllPath" -ForegroundColor Gray
            break
        }
    }
    
    $nugetDll = "C:\Users\soperbp\.nuget\packages\microsoft.data.sqlite.core\10.0.0\lib\netstandard2.0\Microsoft.Data.Sqlite.dll"
    if (Test-Path $nugetDll) {
        try {
            Add-Type -Path $nugetDll -ErrorAction Stop
            $script:SQLiteAssembly = $true
        }
        catch {
            Write-Warning "Could not load Microsoft.Data.Sqlite: $_"
        }
    }
    
    if ($script:SQLiteAssembly) {
        $script:Connection = New-Object Microsoft.Data.Sqlite.SqliteConnection
        $script:Connection.ConnectionString = "Data Source=$DatabasePath"
        $script:Connection.Open()
        
        $createTableCmd = $script:Connection.CreateCommand()
        $createTableCmd.CommandText = @"
CREATE TABLE IF NOT EXISTS edi_files (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    full_path TEXT NOT NULL UNIQUE,
    relative_path TEXT NOT NULL,
    directory TEXT NOT NULL,
    extension TEXT,
    length INTEGER,
    last_write_utc TEXT,
    content TEXT,
    content_hash TEXT,
    indexed_utc TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_name ON edi_files(name);
CREATE INDEX IF NOT EXISTS idx_extension ON edi_files(extension);
CREATE INDEX IF NOT EXISTS idx_full_path ON edi_files(full_path);
"@
        $createTableCmd.ExecuteNonQuery() | Out-Null
        $createTableCmd.Dispose()
        
        Write-Host "Database initialized at: $DatabasePath" -ForegroundColor Green
    }
    else {
        Write-Warning "Could not initialize SQLite. Falling back to file-based indexing."
    }
}

function Index-EdiFiles {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePath,
        
        [int]$MaxContentSize = 50000
    )
    
    if ($null -eq $script:Connection) {
        throw "Database not initialized. Call Initialize-EDIDatabase first."
    }
    
    if (-not (Test-Path $SourcePath)) {
        throw "Source path not found: $SourcePath"
    }
    
    Write-Host "Scanning: $SourcePath" -ForegroundColor Cyan
    $allFiles = @(Get-ChildItem -LiteralPath $SourcePath -Recurse -File -ErrorAction SilentlyContinue)
    Write-Host "Found $($allFiles.Count) files" -ForegroundColor Gray
    
    $existingPaths = @{}
    $checkCmd = $script:Connection.CreateCommand()
    $checkCmd.CommandText = "SELECT full_path, content_hash FROM edi_files"
    $reader = $checkCmd.ExecuteReader()
    while ($reader.Read()) {
        $existingPaths[$reader.GetString(0)] = $reader.GetString(1)
    }
    $reader.Close()
    $checkCmd.Dispose()
    
    $filesToProcess = @()
    foreach ($file in $allFiles) {
        $hash = "$($file.Length)-$($file.LastWriteTimeUtc.ToString('o'))"
        if (-not $existingPaths.ContainsKey($file.FullName) -or $existingPaths[$file.FullName] -ne $hash) {
            $filesToProcess += $file
        }
    }
    
    Write-Host "Need to index: $($filesToProcess.Count) files ($($allFiles.Count - $filesToProcess.Count) unchanged)" -ForegroundColor Yellow
    
    $insertCmd = $script:Connection.CreateCommand()
    $insertCmd.CommandText = @"
INSERT OR REPLACE INTO edi_files 
(name, full_path, relative_path, directory, extension, length, last_write_utc, content, content_hash, indexed_utc)
VALUES (@name, @fullPath, @relativePath, @directory, @extension, @length, @lastWriteUtc, @content, @contentHash, @indexedUtc)
"@
    
    $count = 0
    $startTime = Get-Date
    
    foreach ($file in $filesToProcess) {
        $content = ""
        try {
            $content = Get-Content -LiteralPath $file.FullName -Raw -TotalCount $MaxContentSize -ErrorAction SilentlyContinue
        } catch { }
        
        $relativePath = if ($file.FullName.StartsWith($SourcePath, [System.StringComparison]::OrdinalIgnoreCase)) {
            $file.FullName.Substring($SourcePath.Length).TrimStart("\")
        } else {
            $file.FullName
        }
        
        $hash = "$($file.Length)-$($file.LastWriteTimeUtc.ToString('o'))"
        
        $insertCmd.Parameters.Clear()
        $insertCmd.Parameters.AddWithValue("@name", $file.Name) | Out-Null
        $insertCmd.Parameters.AddWithValue("@fullPath", $file.FullName) | Out-Null
        $insertCmd.Parameters.AddWithValue("@relativePath", $relativePath) | Out-Null
        $insertCmd.Parameters.AddWithValue("@directory", $file.DirectoryName) | Out-Null
        $insertCmd.Parameters.AddWithValue("@extension", $file.Extension) | Out-Null
        $insertCmd.Parameters.AddWithValue("@length", $file.Length) | Out-Null
        $insertCmd.Parameters.AddWithValue("@lastWriteUtc", $file.LastWriteTimeUtc.ToString('o')) | Out-Null
        $insertCmd.Parameters.AddWithValue("@content", $content) | Out-Null
        $insertCmd.Parameters.AddWithValue("@contentHash", $hash) | Out-Null
        $insertCmd.Parameters.AddWithValue("@indexedUtc", (Get-Date).ToUniversalTime().ToString('o')) | Out-Null
        
        $insertCmd.ExecuteNonQuery() | Out-Null
        $count++
        
        if ($count % 50 -eq 0) {
            $elapsed = (Get-Date) - $startTime
            $rate = if ($elapsed.TotalSeconds -gt 0) { [int]($count / $elapsed.TotalSeconds) } else { 0 }
            Write-Host "  Indexed $count / $($filesToProcess.Count) ($rate files/sec)..." -ForegroundColor Gray
        }
    }
    
    $insertCmd.Dispose()
    
    $totalElapsed = (Get-Date) - $startTime
    $avgRate = if ($totalElapsed.TotalSeconds -gt 0) { [int]($count / $totalElapsed.TotalSeconds) } else { 0 }
    Write-Host "Indexed $count files in $([math]::Round($totalElapsed.TotalSeconds, 1))s ($avgRate files/sec)" -ForegroundColor Green
}

function Search-EdiDatabase {
    param(
        [string]$Query = "",
        [string]$Mode = "all",
        [int]$Limit = 100
    )
    
    if ($null -eq $script:Connection) {
        throw "Database not initialized. Call Initialize-EDIDatabase first."
    }
    
    $query = if ($null -eq $Query) { "" } else { $Query.Trim() }
    $searchAll = [string]::IsNullOrEmpty($query)
    
    $cmd = $script:Connection.CreateCommand()
    
    if ($searchAll) {
        $cmd.CommandText = "SELECT id, name, full_path, relative_path, directory, extension, length, last_write_utc FROM edi_files ORDER BY last_write_utc DESC LIMIT $Limit"
    }
    else {
        $likePattern = "%$query%"
        $whereClause = switch ($Mode.ToLower()) {
            "name"    { "WHERE name LIKE @pattern" }
            "path"    { "WHERE relative_path LIKE @pattern OR full_path LIKE @pattern" }
            "content" { "WHERE content LIKE @pattern" }
            default   { "WHERE name LIKE @pattern OR relative_path LIKE @pattern OR content LIKE @pattern" }
        }
        
        $cmd.CommandText = "SELECT id, name, full_path, relative_path, directory, extension, length, last_write_utc FROM edi_files $whereClause ORDER BY last_write_utc DESC LIMIT $Limit"
        $param = $cmd.CreateParameter()
        $param.ParameterName = "@pattern"
        $param.Value = $likePattern
        $cmd.Parameters.Add($param) | Out-Null
    }
    
    $results = @()
    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        $results += [PSCustomObject]@{
            Id = $reader.GetInt64(0)
            Name = $reader.GetString(1)
            FullPath = $reader.GetString(2)
            RelativePath = $reader.GetString(3)
            Directory = $reader.GetString(4)
            Extension = if ($reader.IsDBNull(5)) { "" } else { $reader.GetString(5) }
            Length = $reader.GetInt64(6)
            LastWriteUtc = $reader.GetString(7)
        }
    }
    $reader.Close()
    $cmd.Dispose()
    
    return $results
}

function Get-EdiFileContent {
    param(
        [Parameter(Mandatory=$true)]
        [int]$FileId
    )
    
    if ($null -eq $script:Connection) {
        throw "Database not initialized. Call Initialize-EDIDatabase first."
    }
    
    $cmd = $script:Connection.CreateCommand()
    $cmd.CommandText = "SELECT content, full_path FROM edi_files WHERE id = @id"
    $param = $cmd.CreateParameter()
    $param.ParameterName = "@id"
    $param.Value = $FileId
    $cmd.Parameters.Add($param) | Out-Null
    
    $reader = $cmd.ExecuteReader()
    $result = $null
    if ($reader.Read()) {
        $result = @{
            Content = if ($reader.IsDBNull(0)) { "" } else { $reader.GetString(0) }
            FullPath = $reader.GetString(1)
        }
    }
    $reader.Close()
    $cmd.Dispose()
    
    return $result
}

function Get-EDIDatabaseStats {
    if ($null -eq $script:Connection) {
        return @{ FileCount = 0; DatabasePath = $null }
    }
    
    $cmd = $script:Connection.CreateCommand()
    $cmd.CommandText = "SELECT COUNT(*) FROM edi_files"
    $count = $cmd.ExecuteScalar()
    $cmd.Dispose()
    
    return @{
        FileCount = $count
        DatabasePath = $script:SQLiteDbPath
    }
}

function Close-EDIDatabase {
    if ($null -ne $script:Connection) {
        $script:Connection.Close()
        $script:Connection.Dispose()
        $script:Connection = $null
    }
}

Export-ModuleMember -Function Initialize-EDIDatabase, Index-EdiFiles, Search-EdiDatabase, Get-EdiFileContent, Get-EDIDatabaseStats, Close-EDIDatabase