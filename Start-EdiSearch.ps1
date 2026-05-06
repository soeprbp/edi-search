param(
    [string]$SharePath = "\\svwpefs\WelchEncoreShare\WELCHPKG\EDI",
    [string]$LogPath = "\\svwpefs\WelchEncoreShare\WELCHPKG\EDI\Logs\EDI_log.txt",
    [string]$IndexPath = "",
    [string]$Prefix = "",
    [int]$Port = 0,
    [int]$MaxResults = 200,
    [int]$ThrottleLimit = [Math]::Max(2, [Environment]::ProcessorCount),
    [int]$MaxContentSize = 50000,
    [switch]$Rebuild,
    [switch]$NoIndex,
    [switch]$NoLogIndex
)

# Fix for when run as compiled exe - determine script location
if ([string]::IsNullOrEmpty($PSScriptRoot)) {
    $scriptRoot = $null
    try {
        $procPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if ($procPath) { $scriptRoot = [System.IO.Path]::GetDirectoryName($procPath) }
    } catch {}
    if ([string]::IsNullOrEmpty($scriptRoot)) {
        try {
            $asmLoc = [System.Reflection.Assembly]::GetExecutingAssembly().Location
            if ($asmLoc -and $asmLoc -notmatch "System.Reflection") {
                $scriptRoot = [System.IO.Path]::GetDirectoryName($asmLoc)
            }
        } catch {}
    }
    if ([string]::IsNullOrEmpty($scriptRoot)) {
        try {
            $parent = Split-Path -Parent $MyInvocation.MyCommand.Path
            if ($parent) { $scriptRoot = $parent }
        } catch {}
    }
    # Look for data folder in current or parent dir
    if ([string]::IsNullOrEmpty($IndexPath)) {
        $testPath = Join-Path $scriptRoot "data\edi-index.json"
        if (Test-Path $testPath) { $IndexPath = $testPath }
        else {
            $parent = Split-Path -Parent $scriptRoot
            $testPath = Join-Path $parent "data\edi-index.json"
            if (Test-Path $testPath) { $IndexPath = $testPath }
            else {
                $testPath = Join-Path $scriptRoot "data\edi-index.clixml"
                if (Test-Path $testPath) { $IndexPath = $testPath }
                else {
                    $parent = Split-Path -Parent $scriptRoot
                    $testPath = Join-Path $parent "data\edi-index.clixml"
                    if (Test-Path $testPath) { $IndexPath = $testPath }
                    else {
                        $IndexPath = Join-Path $scriptRoot "data\edi-index.json"
                    }
                }
            }
        }
    }
}
else {
    $IndexPath = Join-Path $PSScriptRoot "data\edi-index.json"
}

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

function Ensure-ParentDirectory {
    param([string]$Path)

    $parent = Split-Path -Path $Path -Parent
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
}

function Escape-Html {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Trim-Text {
    param(
        [AllowNull()][string]$Text,
        [int]$MaxLength = 220
    )

    if ([string]::IsNullOrEmpty($Text)) { return "" }
    if ($Text.Length -le $MaxLength) { return $Text }
    return $Text.Substring(0, $MaxLength) + "..."
}

function Get-Snippet {
    param(
        [AllowNull()][string]$Content,
        [AllowNull()][string]$Query
    )

    if ([string]::IsNullOrWhiteSpace($Content)) { return "" }
    if ([string]::IsNullOrWhiteSpace($Query)) { return (Trim-Text -Text ($Content -replace "\s+", " ")) }

    $index = $Content.IndexOf($Query, [System.StringComparison]::OrdinalIgnoreCase)
    if ($index -lt 0) {
        return (Trim-Text -Text ($Content -replace "\s+", " "))
    }

    $start = [Math]::Max(0, $index - 90)
    $length = [Math]::Min(220, $Content.Length - $start)
    $snippet = $Content.Substring($start, $length) -replace "\s+", " "
    if ($start -gt 0) { $snippet = "..." + $snippet }
    if (($start + $length) -lt $Content.Length) { $snippet += "..." }
    return $snippet
}

function Get-SafeRelativePath {
    param(
        [string]$Root,
        [string]$Path
    )

    if ($Path.StartsWith($Root, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $Path.Substring($Root.Length).TrimStart("\")
    }

    return $Path
}

function Import-ExistingIndexItems {
    param([string]$OutputPath)

    if (-not (Test-Path -LiteralPath $OutputPath)) {
        return @{}
    }

    $existingIndex = Import-Clixml -LiteralPath $OutputPath
    $map = @{}
    foreach ($item in @($existingIndex.Items)) {
        $map[$item.FullPath] = $item
    }

    return $map
}

function Test-ShareAccessible {
    param([string]$Path)

    try {
        return Test-Path -LiteralPath $Path
    }
    catch {
        return $false
    }
}

function New-IndexedItem {
    param(
        [System.IO.FileInfo]$File,
        [string]$RootPath,
        [int]$MaxContentBytes = 50000
    )

    $content = ""
    try {
        $content = Get-Content -LiteralPath $File.FullName -Raw -TotalCount $MaxContentBytes -ErrorAction SilentlyContinue
    } catch {
        $content = ""
    }

    $relativePath = if ($File.FullName.StartsWith($RootPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        $File.FullName.Substring($RootPath.Length).TrimStart("\")
    } else {
        $File.FullName
    }

    return [pscustomobject]@{
        Name         = $File.Name
        FullPath     = $File.FullName
        RelativePath = $relativePath
        Directory    = $File.DirectoryName
        Extension    = $File.Extension
        Length       = [int64]$File.Length
        LastWriteUtc = $File.LastWriteTimeUtc.ToString("o")
        Content      = $content
    }
}

function Sync-EdiIndex {
    param(
        [string]$SourcePath,
        [string]$OutputPath,
        [int]$Parallelism,
        [int]$MaxContentSize = 50000,
        [switch]$ForceRebuild
    )

    if (-not (Test-ShareAccessible -Path $SourcePath)) {
        throw "Share path not found or not accessible: $SourcePath"
    }

    Ensure-ParentDirectory -Path $OutputPath

    $existingMap = if ($ForceRebuild) { @{} } else { Import-ExistingIndexItems -OutputPath $OutputPath }
    $allFiles = @(Get-ChildItem -LiteralPath $SourcePath -Recurse -File -ErrorAction SilentlyContinue)

    $unchangedItems = New-Object System.Collections.Generic.List[object]
    $workItems = New-Object System.Collections.Generic.List[System.IO.FileInfo]

    foreach ($file in $allFiles) {
        $existing = $existingMap[$file.FullName]
        $currentLastWriteUtc = $file.LastWriteTimeUtc.ToString("o")

        if ($existing -and $existing.Length -eq [int64]$file.Length -and $existing.LastWriteUtc -eq $currentLastWriteUtc) {
            $unchangedItems.Add($existing)
        } else {
            $workItems.Add($file)
        }
    }

    Write-Host ("Files on share: {0}" -f @($allFiles).Count)
    Write-Host ("Reused from index: {0}" -f $unchangedItems.Count)
    Write-Host ("New or changed: {0}" -f $workItems.Count)
    if ($workItems.Count -gt 0 -and $workItems.Count -le 20) {
        Write-Host ("  Changed files:") -ForegroundColor Gray
        foreach ($f in $workItems) { Write-Host ("    {0}" -f $f.Name) -ForegroundColor Gray }
    }

    $indexedWorkItems = @()
    if ($workItems.Count -gt 0) {
        Write-Host "Indexing $($workItems.Count) files..." -ForegroundColor Yellow
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            $indexedWorkItems = $workItems | ForEach-Object -Parallel {
                $content = ""
                try {
                    $content = Get-Content -LiteralPath $_.FullName -Raw -TotalCount $using:MaxContentSize -ErrorAction SilentlyContinue
                } catch {
                    $content = ""
                }

                $relativePath = if ($_.FullName.StartsWith($using:SourcePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $_.FullName.Substring($using:SourcePath.Length).TrimStart("\")
                } else {
                    $_.FullName
                }

                [pscustomobject]@{
                    Name         = $_.Name
                    FullPath     = $_.FullName
                    RelativePath = $relativePath
                    Directory    = $_.DirectoryName
                    Extension    = $_.Extension
                    Length       = [int64]$_.Length
                    LastWriteUtc = $_.LastWriteTimeUtc.ToString("o")
                    Content      = $content
                }
            } -ThrottleLimit $Parallelism
        }
        else {
            foreach ($workItem in $workItems) {
                $indexedWorkItems += New-IndexedItem -File $workItem -RootPath $SourcePath -MaxContentBytes $MaxContentSize
            }
        }
    }

    $combinedItems = @($unchangedItems.ToArray()) + @($indexedWorkItems)
    $sortedItems = @($combinedItems | Sort-Object FullPath)

    Write-Host "Saving index to disk..." -ForegroundColor Yellow
    $payload = [pscustomobject]@{
        GeneratedUtc = (Get-Date).ToUniversalTime().ToString("o")
        SharePath    = $SourcePath
        FileCount    = @($sortedItems).Count
        Items        = $sortedItems
    }

    $payload | ConvertTo-Json -Depth 5 | Set-Content -Path $OutputPath -Encoding UTF8
    return $payload
}

function Load-EdiIndex {
    param(
        [string]$SourcePath,
        [string]$OutputPath,
        [int]$Parallelism,
        [int]$MaxContentSize = 50000,
        [switch]$ForceRebuild
    )

    try {
        return Sync-EdiIndex -SourcePath $SourcePath -OutputPath $OutputPath -Parallelism $Parallelism -MaxContentSize $MaxContentSize -ForceRebuild:$ForceRebuild
    }
    catch {
        if (-not $ForceRebuild -and (Test-Path -LiteralPath $OutputPath)) {
            Write-Warning ("Could not refresh the EDI share right now. Using the last local index instead. Reason: {0}" -f $_.Exception.Message)
            try {
                return Get-Content -LiteralPath $OutputPath -Raw | ConvertFrom-Json
            } catch {
                return Import-Clixml -LiteralPath $OutputPath
            }
        }

        throw
    }
}

function Search-EdiIndex {
    param(
        [pscustomobject]$Index,
        [string]$Query,
        [string]$Mode,
        [int]$Limit
    )

    $query = if ($null -eq $Query) { "" } else { $Query.Trim() }
    $mode = if ([string]::IsNullOrWhiteSpace($Mode)) { "all" } else { $Mode.ToLowerInvariant() }
    
    $items = if ($Index.Items) { $Index.Items } else { @() }
    if ($items.Count -eq 0) { return @() }

    $searchAll = [string]::IsNullOrEmpty($query)
    $isPowerShell7OrHigher = $PSVersionTable.PSVersion.Major -ge 7

    if ($isPowerShell7OrHigher) {
        $allResults = $items | ForEach-Object -Parallel {
            $item = $_
            $queryVar = $using:query
            $modeVar = $using:mode
            $searchAllVar = $using:searchAll

            $nameHit = $false
            $pathHit = $false
            $contentHit = $false

            if ($searchAllVar) {
                $nameHit = $true
            } else {
                if ($modeVar -in @("all", "name")) {
                    if ($item.Name.IndexOf($queryVar, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $nameHit = $true }
                }
                if ($modeVar -in @("all", "path")) {
                    if ($item.RelativePath.IndexOf($queryVar, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $pathHit = $true }
                }
                if ($modeVar -in @("all", "content") -and -not [string]::IsNullOrEmpty($item.Content)) {
                    if ($item.Content.IndexOf($queryVar, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $contentHit = $true }
                }
            }

            if ($nameHit -or $pathHit -or $contentHit) {
                $matchAreas = @()
                if ($nameHit) { $matchAreas += "name" }
                if ($pathHit) { $matchAreas += "path" }
                if ($contentHit) { $matchAreas += "content" }

                [pscustomobject]@{
                    Name         = $item.Name
                    RelativePath = $item.RelativePath
                    FullPath     = $item.FullPath
                    Extension    = $item.Extension
                    Length       = $item.Length
                    LastWriteUtc = $item.LastWriteUtc
                    MatchAreas   = ($matchAreas -join ", ")
                    Content      = $item.Content
                    QueryForSnippet = $queryVar
                }
            }
        } -ThrottleLimit ([Math]::Max(8, [Environment]::ProcessorCount * 2))

        $results = foreach ($r in $allResults) {
            [pscustomobject]@{
                Name         = $r.Name
                RelativePath = $r.RelativePath
                FullPath     = $r.FullPath
                Extension    = $r.Extension
                Length       = $r.Length
                LastWriteUtc = $r.LastWriteUtc
                MatchAreas   = $r.MatchAreas
                Snippet      = Get-Snippet -Content $r.Content -Query $r.QueryForSnippet
            }
        }

        return $results |
            Sort-Object @{ Expression = "LastWriteUtc"; Descending = $true }, @{ Expression = "Name"; Descending = $false } |
            Select-Object -First $Limit
    }
    else {
        $results = foreach ($item in $items) {
            $nameHit = $false
            $pathHit = $false
            $contentHit = $false

            if ($searchAll) {
                $nameHit = $true
            } else {
                if ($mode -in @("all", "name")) {
                    if ($item.Name.IndexOf($query, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $nameHit = $true }
                }
                if ($mode -in @("all", "path")) {
                    if ($item.RelativePath.IndexOf($query, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $pathHit = $true }
                }
                if ($mode -in @("all", "content") -and -not [string]::IsNullOrEmpty($item.Content)) {
                    if ($item.Content.IndexOf($query, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $contentHit = $true }
                }
            }

            if ($nameHit -or $pathHit -or $contentHit) {
                $matchAreas = @()
                if ($nameHit) { $matchAreas += "name" }
                if ($pathHit) { $matchAreas += "path" }
                if ($contentHit) { $matchAreas += "content" }

                [pscustomobject]@{
                    Name         = $item.Name
                    RelativePath = $item.RelativePath
                    FullPath     = $item.FullPath
                    Extension    = $item.Extension
                    Length       = $item.Length
                    LastWriteUtc = $item.LastWriteUtc
                    MatchAreas   = ($matchAreas -join ", ")
                    Snippet      = Get-Snippet -Content $item.Content -Query $query
                }
            }
        }

        return $results |
            Sort-Object @{ Expression = "LastWriteUtc"; Descending = $true }, @{ Expression = "Name"; Descending = $false } |
            Select-Object -First $Limit
    }
}

function Parse-LogLine {
    param([string]$Line)

    $parsed = @{
        FullLine = $Line
        DateTime = ""
        Date = ""
        Time = ""
        RunNo = ""
        Cust = ""
        ShipTo = ""
        TrxType = ""
        Msg = ""
        PID = ""
        DateTimeValue = [DateTime]::MinValue
    }

    if ([string]::IsNullOrWhiteSpace($Line)) { return $parsed }

    $line = $Line.Trim()

    if ($line -match '^(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}:\d{2}:\d{2}):\d{6}\s+(AM|PM)\s+(.*)$') {
        $parsed.Date = $matches[1]
        $parsed.Time = $matches[2] + " " + $matches[3]
        $parsed.DateTime = "$($matches[1]) $($matches[2]) $($matches[3])"
        $remainder = $matches[4]

        try {
            $dateStr = "$($matches[1]) $($matches[2]) $($matches[3])"
            $parsed.DateTimeValue = [DateTime]::Parse($dateStr)
        } catch {
            $parsed.DateTimeValue = [DateTime]::MinValue
        }

        if ($remainder -match 'Cust:\s*(\S*)') {
            $parsed.Cust = $matches[1]
        }

        if ($remainder -match 'ShipTo:\s*(\S*)') {
            $parsed.ShipTo = $matches[1]
        }

        if ($remainder -match 'TrxType:\s*(\S*)') {
            $parsed.TrxType = $matches[1]
        }

        if ($remainder -match 'Msg:\s*(.*)') {
            $parsed.Msg = $matches[1]
        }

        if ($remainder -match 'PID[_\s]*(\d+)') {
            $parsed.PID = $matches[1]
        }
    }

    return $parsed
}

function Build-LogIndex {
    param(
        [string]$LogPath,
        [int]$ThrottleLimit,
        [int]$DaysToKeep = 30
    )

    $index = @{
        Entries = @()
        TotalLines = 0
        MinDate = $null
        MaxDate = $null
        LogPath = $LogPath
        LastModified = $null
        DaysToKeep = $DaysToKeep
    }

    if ([string]::IsNullOrWhiteSpace($LogPath) -or -not (Test-Path -LiteralPath $LogPath)) {
        return $index
    }

    $fileInfo = Get-Item -LiteralPath $LogPath
    $index.LastModified = $fileInfo.LastWriteTimeUtc.ToString("o")

    $cutoffDate = (Get-Date).AddDays(-$DaysToKeep)
    Write-Host "Indexing last $DaysToKeep days of log: $LogPath" -ForegroundColor Cyan
    Write-Host "Cutoff date: $($cutoffDate.ToString('yyyy-MM-dd'))" -ForegroundColor Gray

    $entries = [System.Collections.Generic.List[object]]::new()
    $lineCount = 0
    $skippedLines = 0

    $reader = New-Object System.IO.StreamReader($LogPath)
    try {
        while ($null -ne ($line = $reader.ReadLine())) {
            $lineCount++
            if ($lineCount % 100000 -eq 0) {
                Write-Host "  Read $lineCount lines, indexed $($entries.Count) entries..." -ForegroundColor Gray
            }

            $parsed = Parse-LogLine -Line $line
            if ([string]::IsNullOrWhiteSpace($parsed.DateTime)) {
                continue
            }

            $parsed.SourceLineNumber = $lineCount

            if ($parsed.DateTimeValue -lt $cutoffDate) {
                $skippedLines++
                if ($entries.Count -gt 0 -and $skippedLines -gt 100000) {
                    break
                }
                continue
            }

            $entries.Add($parsed)
            $skippedLines = 0
        }
    }
    finally {
        $reader.Close()
    }

    $index.TotalLines = $lineCount
    $index.Entries = @($entries | Sort-Object DateTimeValue)

    if ($index.Entries.Count -gt 0) {
        $index.MinDate = $index.Entries[0].DateTime
        $index.MaxDate = $index.Entries[-1].DateTime
    }

    Write-Host "Log index built: $($index.Entries.Count) entries" -ForegroundColor Green
    Write-Host "  Date range: $($index.MinDate) to $($index.MaxDate)" -ForegroundColor Gray

    return $index
}

function Get-FirstIndexOnOrAfterDate {
    param(
        [object[]]$Entries,
        [DateTime]$TargetDate
    )

    $low = 0
    $high = $Entries.Count - 1
    $result = -1

    while ($low -le $high) {
        $mid = [int](($low + $high) / 2)
        if ($Entries[$mid].DateTimeValue -ge $TargetDate) {
            $result = $mid
            $high = $mid - 1
        }
        else {
            $low = $mid + 1
        }
    }

    return $result
}

function Get-LastIndexOnOrBeforeDate {
    param(
        [object[]]$Entries,
        [DateTime]$TargetDate
    )

    $low = 0
    $high = $Entries.Count - 1
    $result = -1

    while ($low -le $high) {
        $mid = [int](($low + $high) / 2)
        if ($Entries[$mid].DateTimeValue -le $TargetDate) {
            $result = $mid
            $low = $mid + 1
        }
        else {
            $high = $mid - 1
        }
    }

    return $result
}

function Search-LogIndex {
    param(
        [object]$LogIndex,
        [string]$Query,
        [int]$Limit = 50,
        [Nullable[DateTime]]$StartDate = $null,
        [Nullable[DateTime]]$EndDate = $null,
        [int]$ContextBefore = 3,
        [int]$ContextAfter = 5
    )

    $results = @()
    $totalMatches = 0
    $hasMore = $false

    if ($null -eq $LogIndex -or $LogIndex.Entries.Count -eq 0) {
        return [pscustomobject]@{
            Results = $results
            Error = $null
            TotalEntries = 0
            TotalMatches = 0
            MinDate = $null
            MaxDate = $null
        }
    }

    try {
        $entries = @($LogIndex.Entries)
        $startIndex = 0
        $endIndex = $entries.Count - 1

        if ($StartDate.HasValue) {
            $startIndex = Get-FirstIndexOnOrAfterDate -Entries $entries -TargetDate $StartDate.Value
            if ($startIndex -lt 0) {
                return [pscustomobject]@{
                    Results = @()
                    Error = $null
                    TotalEntries = $LogIndex.Entries.Count
                    TotalMatches = 0
                    HasMore = $false
                    MinDate = $LogIndex.MinDate
                    MaxDate = $LogIndex.MaxDate
                }
            }
        }

        if ($EndDate.HasValue) {
            $endIndex = Get-LastIndexOnOrBeforeDate -Entries $entries -TargetDate $EndDate.Value
            if ($endIndex -lt 0) {
                return [pscustomobject]@{
                    Results = @()
                    Error = $null
                    TotalEntries = $LogIndex.Entries.Count
                    TotalMatches = 0
                    HasMore = $false
                    MinDate = $LogIndex.MinDate
                    MaxDate = $LogIndex.MaxDate
                }
            }
        }

        if ($startIndex -gt $endIndex) {
            return [pscustomobject]@{
                Results = @()
                Error = $null
                TotalEntries = $LogIndex.Entries.Count
                TotalMatches = 0
                HasMore = $false
                MinDate = $LogIndex.MinDate
                MaxDate = $LogIndex.MaxDate
            }
        }

        $queryText = if ($Query) { $Query.Trim() } else { "" }
        $matchIndices = New-Object System.Collections.Generic.List[int]

        for ($i = $endIndex; $i -ge $startIndex; $i--) {
            $line = [string]$entries[$i].FullLine
            $isMatch = if ([string]::IsNullOrWhiteSpace($queryText)) {
                $true
            }
            else {
                $line.IndexOf($queryText, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            }

            if (-not $isMatch) { continue }

            $totalMatches++
            if ($matchIndices.Count -lt $Limit) {
                $matchIndices.Add($i)
            }
            else {
                $hasMore = $true
                break
            }
        }

        foreach ($entryIndex in $matchIndices) {
            $entry = $entries[$entryIndex]
            $contextStart = [Math]::Max($startIndex, $entryIndex - $ContextBefore)
            $contextEnd = [Math]::Min($endIndex, $entryIndex + $ContextAfter)

            $contextLines = @()
            for ($i = $contextStart; $i -le $contextEnd; $i++) {
                $ctx = $entries[$i]
                $contextLines += @{
                    LineNumber = if ($ctx.SourceLineNumber) { $ctx.SourceLineNumber } else { $i + 1 }
                    Parsed = @{
                        FullLine = $ctx.FullLine
                        DateTime = $ctx.DateTime
                        Cust = $ctx.Cust
                        ShipTo = $ctx.ShipTo
                        TrxType = $ctx.TrxType
                        Msg = $ctx.Msg
                        PID = $ctx.PID
                    }
                    IsMatch = ($i -eq $entryIndex)
                }
            }

            $results += [pscustomobject]@{
                LineNumber = if ($entry.SourceLineNumber) { $entry.SourceLineNumber } else { $entryIndex + 1 }
                Parsed = @{
                    FullLine = $entry.FullLine
                    DateTime = $entry.DateTime
                    Cust = $entry.Cust
                    ShipTo = $entry.ShipTo
                    TrxType = $entry.TrxType
                    Msg = $entry.Msg
                    PID = $entry.PID
                }
                ContextBefore = $ContextBefore
                ContextAfter = $ContextAfter
                ContextLines = $contextLines
            }
        }
    }
    catch {
        return [pscustomobject]@{
            Results = @()
            Error = $_.Exception.Message
            TotalEntries = $LogIndex.Entries.Count
            TotalMatches = 0
            HasMore = $false
            MinDate = $LogIndex.MinDate
            MaxDate = $LogIndex.MaxDate
        }
    }

    return [pscustomobject]@{
        Results = $results
        Error = $null
        TotalEntries = $LogIndex.Entries.Count
        TotalMatches = $totalMatches
        HasMore = $hasMore
        MinDate = $LogIndex.MinDate
        MaxDate = $LogIndex.MaxDate
    }
}

function Get-LogErrorSummary {
    param(
        [object]$LogIndex,
        [int]$DaysBack = 30,
        [int]$MaxErrorsPerCategory = 100
    )

    if ($null -eq $LogIndex -or $LogIndex.Entries.Count -eq 0) {
        return $null
    }

    $cutoff = (Get-Date).AddDays(-$DaysBack)
    $errorCategories = @{
        "Invalid Paper Code" = [System.Collections.Generic.List[string]]::new()
        "Invalid Vendor" = [System.Collections.Generic.List[string]]::new()
        "PO Receipt - Line Not Exist" = [System.Collections.Generic.List[string]]::new()
        "Other Errors" = [System.Collections.Generic.List[string]]::new()
    }

    foreach ($entry in $LogIndex.Entries) {
        if ($entry.DateTimeValue -lt $cutoff) { break }
        if ([string]::IsNullOrWhiteSpace($entry.Msg)) { continue }

        $msg = $entry.Msg
        $po = ""
        if ($msg -match '(?:for|PO/?)\s*([A-Z]?\d+[-\d]*)') {
            $po = " (PO: $($matches[1]))"
        }

        if ($msg -match 'Invalid Paper Code\s*\(([^)]+)\)') {
            $detail = "Code: $($matches[1])$po"
            if ($errorCategories["Invalid Paper Code"].Count -lt $MaxErrorsPerCategory) {
                $errorCategories["Invalid Paper Code"].Add($detail)
            }
        }
        elseif ($msg -match 'Invalid Vendor\s*\(([^)]+)\)') {
            $detail = "Vendor: $($matches[1])$po"
            if ($errorCategories["Invalid Vendor"].Count -lt $MaxErrorsPerCategory) {
                $errorCategories["Invalid Vendor"].Add($detail)
            }
        }
        elseif ($msg -match 'PO Receipt.*not exist\s*\(([^)]+)\)') {
            $detail = "$($matches[1])$po"
            if ($errorCategories["PO Receipt - Line Not Exist"].Count -lt $MaxErrorsPerCategory) {
                $errorCategories["PO Receipt - Line Not Exist"].Add($detail)
            }
        }
        elseif ($msg -match '(?i)ERROR|FAIL|EXCEPTION') {
            $detail = $msg.Substring(0, [Math]::Min(120, $msg.Length))
            if ($errorCategories["Other Errors"].Count -lt $MaxErrorsPerCategory) {
                $errorCategories["Other Errors"].Add($detail)
            }
        }
    }

    $hasErrors = $false
    $summary = @{}
    foreach ($cat in $errorCategories.Keys) {
        if ($errorCategories[$cat].Count -gt 0) {
            $hasErrors = $true
            $summary[$cat] = @($errorCategories[$cat])
        }
    }

    if (-not $hasErrors) {
        return @{ hasErrors = $false; message = "No errors in last $DaysBack days"; categories = @{} }
    }

    return @{
        hasErrors = $true
        daysBack = $DaysBack
        categories = $summary
        totalCategories = $summary.Count
    }
}

function New-LogSearchCacheKey {
    param(
        [string]$Query,
        [string]$StartDate,
        [string]$EndDate,
        [int]$Limit,
        [object]$LogIndex
    )

    $indexVersion = Get-LogIndexVersion -LogIndex $LogIndex

    $queryPart = if ($null -eq $Query) { "" } else { $Query.Trim().ToLowerInvariant() }
    $startPart = if ($null -eq $StartDate) { "" } else { $StartDate }
    $endPart = if ($null -eq $EndDate) { "" } else { $EndDate }

    return "q={0}|s={1}|e={2}|l={3}|v={4}" -f $queryPart, $startPart, $endPart, $Limit, $indexVersion
}

function Get-LogIndexVersion {
    param([object]$LogIndex)

    if ($null -eq $LogIndex) {
        return "none"
    }

    return "{0}|{1}|{2}" -f $LogIndex.LastModified, $LogIndex.Entries.Count, $LogIndex.DaysToKeep
}

function Get-LogSearchCacheEntry {
    param([string]$Key)

    if (-not $script:LogSearchCache.ContainsKey($Key)) {
        return $null
    }

    $entry = $script:LogSearchCache[$Key]
    $ageSeconds = ((Get-Date).ToUniversalTime() - [DateTime]::Parse($entry.CreatedUtc)).TotalSeconds
    if ($ageSeconds -gt $script:LogSearchCacheTtlSeconds) {
        [void]$script:LogSearchCache.Remove($Key)
        [void]$script:LogSearchCacheOrder.Remove($Key)
        return $null
    }

    return $entry.Payload
}

function Set-LogSearchCacheEntry {
    param(
        [string]$Key,
        [hashtable]$Payload,
        [string]$Query,
        [string]$StartDate,
        [string]$EndDate,
        [int]$Limit,
        [object]$LogIndex
    )

    if ($script:LogSearchCache.ContainsKey($Key)) {
        [void]$script:LogSearchCacheOrder.Remove($Key)
    }

    $script:LogSearchCache[$Key] = [pscustomobject]@{
        CreatedUtc = (Get-Date).ToUniversalTime().ToString("o")
        Payload    = $Payload
        Meta       = [pscustomobject]@{
            QueryNormalized = if ($null -eq $Query) { "" } else { $Query.Trim().ToLowerInvariant() }
            StartDate = if ($null -eq $StartDate) { "" } else { $StartDate }
            EndDate = if ($null -eq $EndDate) { "" } else { $EndDate }
            Limit = $Limit
            IndexVersion = Get-LogIndexVersion -LogIndex $LogIndex
        }
    }
    $script:LogSearchCacheOrder.Add($Key)

    while ($script:LogSearchCacheOrder.Count -gt $script:LogSearchCacheMaxEntries) {
        $oldestKey = $script:LogSearchCacheOrder[0]
        $script:LogSearchCacheOrder.RemoveAt(0)
        [void]$script:LogSearchCache.Remove($oldestKey)
    }
}

function Get-LogSearchPrefixReusePayload {
    param(
        [string]$Query,
        [string]$StartDate,
        [string]$EndDate,
        [int]$Limit,
        [object]$LogIndex
    )

    $queryNormalized = if ($null -eq $Query) { "" } else { $Query.Trim().ToLowerInvariant() }
    if ([string]::IsNullOrWhiteSpace($queryNormalized)) {
        return $null
    }

    $startPart = if ($null -eq $StartDate) { "" } else { $StartDate }
    $endPart = if ($null -eq $EndDate) { "" } else { $EndDate }
    $indexVersion = Get-LogIndexVersion -LogIndex $LogIndex

    $bestEntry = $null
    $bestPrefixLength = -1

    foreach ($pair in $script:LogSearchCache.GetEnumerator()) {
        $entry = $pair.Value
        $ageSeconds = ((Get-Date).ToUniversalTime() - [DateTime]::Parse($entry.CreatedUtc)).TotalSeconds
        if ($ageSeconds -gt $script:LogSearchCacheTtlSeconds) {
            continue
        }

        $meta = $entry.Meta
        if ($null -eq $meta) { continue }
        if ($meta.StartDate -ne $startPart) { continue }
        if ($meta.EndDate -ne $endPart) { continue }
        if ([int]$meta.Limit -ne $Limit) { continue }
        if ($meta.IndexVersion -ne $indexVersion) { continue }

        $baseQuery = [string]$meta.QueryNormalized
        if ([string]::IsNullOrWhiteSpace($baseQuery)) { continue }
        if ($baseQuery.Length -ge $queryNormalized.Length) { continue }
        if (-not $queryNormalized.StartsWith($baseQuery, [System.StringComparison]::Ordinal)) { continue }

        $payload = $entry.Payload
        if ($null -eq $payload) { continue }
        if ($payload.hasMore -eq $true) { continue }

        if ($baseQuery.Length -gt $bestPrefixLength) {
            $bestPrefixLength = $baseQuery.Length
            $bestEntry = $entry
        }
    }

    if ($null -eq $bestEntry) {
        return $null
    }

    $basePayload = $bestEntry.Payload
    $baseResults = @($basePayload.results)
    $filteredResults = @()
    foreach ($result in $baseResults) {
        $line = [string]$result.parsed.FullLine
        if ($line.IndexOf($queryNormalized, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $filteredResults += $result
        }
    }

    return @{
        query = $Query
        limit = $Limit
        startDate = $StartDate
        endDate = $EndDate
        elapsedMs = 0
        cacheHit = $true
        cacheSource = "prefix"
        logPath = $basePayload.logPath
        totalEntries = $basePayload.totalEntries
        totalMatches = $filteredResults.Count
        hasMore = $false
        minDate = $basePayload.minDate
        maxDate = $basePayload.maxDate
        error = $null
        results = @($filteredResults)
    }
}

function Write-JsonResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [object]$Payload,
        [int]$StatusCode = 200
    )

    $json = $Payload | ConvertTo-Json -Depth 6
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $Response.StatusCode = $StatusCode
    $Response.ContentType = "application/json; charset=utf-8"
    $Response.ContentLength64 = $bytes.Length
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.OutputStream.Close()
}

function Write-TextResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [string]$Body,
        [string]$ContentType = "text/html; charset=utf-8",
        [int]$StatusCode = 200
    )

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $Response.StatusCode = $StatusCode
    $Response.ContentType = $ContentType
    $Response.ContentLength64 = $bytes.Length
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.OutputStream.Close()
}

function Get-IndexSummary {
    param([pscustomobject]$Index)

    return [pscustomobject]@{
        sharePath    = $Index.SharePath
        generatedUtc = $Index.GeneratedUtc
        fileCount    = $Index.FileCount
    }
}

function Get-CovDropTargetPath {
    param([string]$SourcePath)

    if ([string]::IsNullOrWhiteSpace($SourcePath)) {
        throw "Source path is required."
    }

    if (-not $SourcePath.EndsWith(".bak", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Only .bak files can be copied as .cov."
    }

    $processedSegment = "\Processed\"
    $processedIndex = $SourcePath.IndexOf($processedSegment, [System.StringComparison]::OrdinalIgnoreCase)
    if ($processedIndex -lt 0) {
        throw "Source path must contain a '\Processed\' segment."
    }

    $prefix = $SourcePath.Substring(0, $processedIndex)
    $fileName = Split-Path -Path $SourcePath -Leaf
    $targetPath = Join-Path -Path $prefix -ChildPath $fileName
    return [System.IO.Path]::ChangeExtension($targetPath, ".cov")
}

function Copy-BakToCovDrop {
    param([string]$SourcePath)

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
        throw "Source file not found: $SourcePath"
    }

    $targetPath = Get-CovDropTargetPath -SourcePath $SourcePath
    $targetDirectory = Split-Path -Path $targetPath -Parent

    if (-not (Test-Path -LiteralPath $targetDirectory -PathType Container)) {
        throw "Target folder does not exist: $targetDirectory"
    }

    if (Test-Path -LiteralPath $targetPath -PathType Leaf) {
        throw "Target file already exists: $targetPath"
    }

    Copy-Item -LiteralPath $SourcePath -Destination $targetPath -ErrorAction Stop

    return [pscustomobject]@{
        sourcePath = $SourcePath
        targetPath = $targetPath
    }
}

function Get-PageHtml {
    param([pscustomobject]$Index)

    $template = @'
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>EDI Search</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Nunito+Sans:wght@400;600;700;800&family=Arimo:wght@400;700&display=swap');
    :root {
      --bg: #edf4ee;
      --bg2: #f7fbf8;
      --panel: #ffffff;
      --ink: #181c18;
      --muted: #4f6152;
      --line: #d5e3d7;
      --brand: #52c75a;
      --brand-deep: #2f8f3f;
      --brand-soft: #e7f6e9;
      --danger-soft: #fceae5;
      --danger: #b9442d;
      --shadow: 0 14px 34px rgba(24, 28, 24, 0.09);
      --radius: 16px;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Nunito Sans", "Segoe UI", Tahoma, sans-serif;
      color: var(--ink);
      background:
        radial-gradient(1400px 600px at 90% -10%, rgba(82, 199, 90, 0.18), transparent 65%),
        linear-gradient(180deg, var(--bg2) 0%, var(--bg) 100%);
      min-height: 100vh;
    }
    .wrap {
      max-width: 1240px;
      margin: 0 auto;
      padding: 28px 18px 48px;
    }
    .hero {
      background:
        linear-gradient(145deg, rgba(24, 28, 24, 0.92), rgba(32, 45, 34, 0.93)),
        radial-gradient(circle at top left, rgba(82, 199, 90, 0.35), transparent 50%);
      border: 1px solid rgba(82, 199, 90, 0.32);
      border-radius: 22px;
      padding: 24px;
      box-shadow: var(--shadow);
      color: #f1fff3;
    }
    h1 {
      margin: 0 0 8px;
      font-size: clamp(1.9rem, 4.6vw, 3rem);
      line-height: 1.05;
      letter-spacing: -0.02em;
      font-weight: 800;
    }
    p {
      margin: 0;
      color: rgba(241, 255, 243, 0.83);
      font-size: 1rem;
      max-width: 75ch;
    }
    .meta {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 12px;
      margin-top: 20px;
    }
    .meta-card,
    .search-panel,
    .results-panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
    }
    .meta-card {
      background: rgba(255, 255, 255, 0.05);
      border-color: rgba(82, 199, 90, 0.24);
      padding: 14px 16px;
    }
    .meta-label {
      display: block;
      color: rgba(241, 255, 243, 0.7);
      font-size: 0.74rem;
      text-transform: uppercase;
      letter-spacing: 0.09em;
    }
    .meta-value {
      display: block;
      margin-top: 6px;
      font-size: 0.98rem;
      word-break: break-word;
      font-family: "Arimo", sans-serif;
    }
    .search-panel {
      margin-top: 18px;
      padding: 18px;
    }
    .search-title {
      font-weight: 800;
      letter-spacing: 0.02em;
      margin: 0 0 12px;
      color: #243227;
    }
    form {
      display: grid;
      grid-template-columns: minmax(0, 1.6fr) 210px 120px 140px;
      gap: 10px;
    }
    .log-form {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
    }
    .log-form input[type="text"] {
      flex: 2;
      min-width: 220px;
    }
    .log-form input[type="date"] { width: 146px; }
    .log-form input[type="number"] { width: 76px; }
    .log-form label {
      font-size: 0.84rem;
      color: var(--muted);
      font-weight: 700;
    }
    .hint {
      margin-top: 8px;
      color: var(--muted);
      font-size: 0.81rem;
    }
    input,
    select {
      width: 100%;
      border-radius: 11px;
      border: 1px solid var(--line);
      padding: 11px 13px;
      font: inherit;
      background: #fff;
      transition: border-color 0.16s ease, box-shadow 0.16s ease, transform 0.12s ease;
    }
    button {
      border-radius: 11px;
      border: 1px solid transparent;
      padding: 11px 16px;
      font: inherit;
      width: auto;
    }
    input:focus,
    select:focus {
      outline: none;
      border-color: var(--brand-deep);
      box-shadow: 0 0 0 3px rgba(82, 199, 90, 0.2);
    }
    button {
      cursor: pointer;
      border: none;
      background: linear-gradient(120deg, var(--brand), var(--brand-deep));
      color: #f7fff8;
      font-weight: 800;
      letter-spacing: 0.01em;
    }
    button.secondary {
      background: linear-gradient(120deg, #253427, #364a39);
    }
    #searchForm button { width: 100%; }
    .log-form button {
      width: auto;
      min-width: 116px;
      padding-inline: 18px;
    }
    button:hover { transform: translateY(-1px); }
    .toolbar {
      display: flex;
      gap: 12px;
      align-items: center;
      margin-top: 12px;
      color: var(--muted);
      flex-wrap: wrap;
    }
    .toolbar .secondary {
      min-width: 180px;
    }
    .status-text {
      display: inline-flex;
      align-items: center;
      padding: 6px 10px;
      border-radius: 999px;
      background: var(--brand-soft);
      border: 1px solid #caeccc;
      font-size: 0.88rem;
    }
    .results-panel {
      margin-top: 18px;
      padding: 8px;
    }
    .results-head {
      display: flex;
      justify-content: space-between;
      gap: 16px;
      align-items: center;
      padding: 12px 14px;
      color: var(--muted);
      font-size: 0.9rem;
    }
    .panel-head {
      padding: 0 0 12px;
    }
    .result {
      border-top: 1px solid #e4efe5;
      padding: 16px 14px;
      animation: rise 0.18s ease;
    }
    .result:first-of-type { border-top: 0; }
    .result-header {
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      gap: 12px;
    }
    .result-header h3 {
      margin: 0;
      font-size: 1.03rem;
      flex: 1;
      min-width: 0;
      font-weight: 800;
    }
    .result-header a {
      color: #234228;
      text-decoration: none;
    }
    .result-header a:hover {
      text-decoration: underline;
      text-underline-offset: 2px;
    }
    .result-date {
      font-size: 0.82rem;
      color: var(--muted);
      white-space: nowrap;
    }
    .path {
      margin-top: 4px;
      color: var(--muted);
      word-break: break-word;
      font-size: 0.92rem;
      font-family: "Arimo", sans-serif;
    }
    .snippet {
      margin-top: 10px;
      font-family: Consolas, "Courier New", monospace;
      background: #fbfefb;
      border: 1px solid #e1eee2;
      border-radius: 10px;
      padding: 10px 12px;
      white-space: pre-wrap;
      color: #1f2b21;
    }
    .badges {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-top: 10px;
    }
    .badge {
      border-radius: 999px;
      padding: 4px 10px;
      background: var(--brand-soft);
      color: #25592e;
      border: 1px solid #cceacf;
      font-size: 0.79rem;
      font-weight: 700;
    }
    .badge-primary {
      background: linear-gradient(120deg, var(--brand), var(--brand-deep));
      color: #f7fff8;
      border-color: transparent;
    }
    .badge-success {
      background: #2e4633;
      color: #e9ffe9;
      border-color: #364f3a;
    }
    .result-actions {
      display: flex;
      gap: 10px;
      margin-top: 12px;
      flex-wrap: wrap;
    }
    .action-btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border-radius: 10px;
      border: 1px solid var(--line);
      background: #fff;
      color: #1f2b21;
      padding: 9px 12px;
      cursor: pointer;
      font: inherit;
      font-weight: 700;
    }
    .action-btn:hover {
      border-color: var(--brand);
      color: #1e4e26;
      background: #f7fff7;
    }
    .empty {
      padding: 28px 16px 34px;
      text-align: center;
      color: var(--muted);
    }
    .alert-error {
      color: var(--danger);
      background: var(--danger-soft);
      border: 1px solid #f1d1c8;
      border-radius: 10px;
      padding: 10px 12px;
      font-weight: 700;
    }
    .loading { opacity: 0.7; }
    
    .indexing-overlay {
      display: none;
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(237, 244, 238, 0.92);
      z-index: 1000;
      justify-content: center;
      align-items: center;
      flex-direction: column;
      gap: 24px;
    }
    .indexing-overlay.active {
      display: flex;
    }
    .indexing-card {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 20px;
      padding: 40px 50px;
      text-align: center;
      box-shadow: 0 20px 60px rgba(24, 28, 24, 0.15);
      max-width: 420px;
    }
    .indexing-spinner {
      width: 56px;
      height: 56px;
      border: 4px solid var(--line);
      border-top-color: var(--brand);
      border-radius: 50%;
      animation: spin 1s linear infinite;
      margin: 0 auto 20px;
    }
    @keyframes spin {
      to { transform: rotate(360deg); }
    }
    .indexing-title {
      font-size: 1.4rem;
      font-weight: 800;
      color: var(--ink);
      margin: 0 0 8px;
    }
    .indexing-status {
      color: var(--muted);
      font-size: 1rem;
      margin: 0 0 20px;
    }
    .indexing-progress {
      background: var(--bg);
      border-radius: 10px;
      height: 10px;
      overflow: hidden;
      margin-bottom: 12px;
    }
    .indexing-progress-bar {
      background: linear-gradient(90deg, var(--brand), var(--brand-deep));
      height: 100%;
      width: 0%;
      border-radius: 10px;
      transition: width 0.3s ease;
      animation: shimmer 1.5s ease-in-out infinite;
    }
    @keyframes shimmer {
      0%, 100% { opacity: 1; }
      50% { opacity: 0.7; }
    }
    .indexing-count {
      font-size: 0.9rem;
      color: var(--muted);
    }
    .indexing-log {
      font-family: monospace;
      font-size: 0.8rem;
      color: var(--muted);
      background: var(--bg);
      padding: 10px;
      border-radius: 8px;
      max-width: 100%;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .log-message {
      margin-top: 8px;
      font-weight: 700;
      color: #253d28;
    }
    .log-hit { margin-bottom: 16px; }
    .log-context {
      margin-top: 12px;
      background: #111615;
      border-radius: 10px;
      overflow: hidden;
      border: 1px solid #2f3a31;
      font-family: Consolas, "Courier New", monospace;
      font-size: 0.82rem;
    }
    .log-line {
      display: flex;
      padding: 4px 10px;
      color: #8ea092;
      border-left: 3px solid transparent;
    }
    .log-line.match {
      background: rgba(82, 199, 90, 0.2);
      color: #effff1;
      border-left-color: var(--brand);
    }
    .line-num {
      min-width: 58px;
      color: #627266;
      text-align: right;
      padding-right: 12px;
      user-select: none;
    }
    .log-line.match .line-num,
    .log-line.match .line-prefix {
      color: #88ea91;
      font-weight: 700;
    }
    .line-prefix {
      min-width: 30px;
      color: #627266;
    }
    .line-content {
      white-space: pre-wrap;
      word-break: break-all;
    }
    @keyframes rise {
      from { opacity: 0; transform: translateY(6px); }
      to { opacity: 1; transform: translateY(0); }
    }
    @media (max-width: 980px) {
      form { grid-template-columns: 1fr; }
      .results-head { flex-direction: column; align-items: flex-start; }
      .log-form input[type="date"] { width: 100%; }
    }
  </style>
</head>
<body>
    <div id="indexingOverlay" class="indexing-overlay">
      <div class="indexing-card">
        <div class="indexing-spinner"></div>
        <h2 class="indexing-title">Indexing Files</h2>
        <p class="indexing-status" id="indexingStatus">Scanning share...</p>
        <div class="indexing-progress">
          <div class="indexing-progress-bar" id="indexingProgressBar"></div>
        </div>
        <p class="indexing-count" id="indexingCount">0 files processed</p>
        <div class="indexing-log" id="indexingLog"></div>
      </div>
    </div>

    <div class="wrap">
    <section class="hero">
      <h1>EDI Search</h1>
      <p>Local browser search for filename, path, and file contents. The share remains untouched; this tool only reads from it and stores its own index locally.</p>
      <div class="meta">
        <div class="meta-card">
          <span class="meta-label">Share</span>
          <span class="meta-value">__SHARE_PATH__</span>
        </div>
        <div class="meta-card">
          <span class="meta-label">Indexed Files</span>
          <span class="meta-value" id="fileCount">__FILE_COUNT__</span>
        </div>
        <div class="meta-card">
          <span class="meta-label">Index Built</span>
          <span class="meta-value" id="generatedUtc">__GENERATED_UTC__</span>
        </div>
      </div>
    </section>

    <section class="search-panel">
      <h2 class="search-title">File Search</h2>
      <form id="searchForm">
        <input id="query" name="query" type="text" placeholder="Search PO number, customer, SKU, filename, path..." autocomplete="off">
        <select id="mode" name="mode">
          <option value="all">Name + Path + Contents</option>
          <option value="name">Filename Only</option>
          <option value="path">Path Only</option>
          <option value="content">Contents Only</option>
        </select>
        <input id="limit" name="limit" type="number" min="1" max="500" value="100">
        <button type="submit">Search</button>
      </form>
      <div class="toolbar">
        <button id="reindexBtn" class="secondary" type="button">Refresh Index</button>
        <span id="status" class="status-text">Ready.</span>
      </div>
    </section>

    <section class="results-panel">
      <div class="results-head">
        <strong id="resultCount">No search yet</strong>
        <span>Tip: leave the query empty and search by recent files.</span>
      </div>
      <div id="results" class="empty">Run a search to see matching files.</div>
    </section>

    <section class="search-panel">
      <div class="results-head panel-head">
        <h2 class="search-title">Log Search</h2>
        <span id="logDateRange"></span>
      </div>
      <form id="logSearchForm" class="log-form">
        <input id="logQuery" name="logQuery" type="text" placeholder="Search log entries..." autocomplete="off">
        <label for="logStartDate">From:</label>
        <input id="logStartDate" name="logStartDate" type="date">
        <label for="logEndDate">To:</label>
        <input id="logEndDate" name="logEndDate" type="date">
        <input id="logLimit" name="logLimit" type="number" min="1" max="200" value="50">
        <button type="submit">Search</button>
      </form>
      <div class="hint">
        Leave dates empty to search all entries. Format: YYYY-MM-DD
      </div>
    </section>

    <section class="results-panel" id="logResultsPanel">
      <div class="results-head">
        <strong id="logResultCount">No log search yet</strong>
      </div>
      <div id="logResults" class="empty">Run a log search to see entries.</div>
    </section>

    <section class="search-panel" id="errorSummaryPanel">
      <div class="results-head panel-head">
        <h2 class="search-title">Error Summary (Last 30 Days)</h2>
        <button id="refreshErrorsBtn" class="secondary" type="button">Refresh</button>
      </div>
      <div id="errorSummaryContent">
        <div class="empty">Loading error summary...</div>
      </div>
    </section>
  </div>

  <script>
    const form = document.getElementById('searchForm');
    const resultsEl = document.getElementById('results');
    const resultCountEl = document.getElementById('resultCount');
    const statusEl = document.getElementById('status');
    const generatedUtcEl = document.getElementById('generatedUtc');
    const fileCountEl = document.getElementById('fileCount');
    const reindexBtn = document.getElementById('reindexBtn');
    const indexingOverlay = document.getElementById('indexingOverlay');
    const indexingStatus = document.getElementById('indexingStatus');
    const indexingProgressBar = document.getElementById('indexingProgressBar');
    const indexingCount = document.getElementById('indexingCount');
    const indexingLog = document.getElementById('indexingLog');
    const logSearchForm = document.getElementById('logSearchForm');
    const logResultsEl = document.getElementById('logResults');
    const logResultCountEl = document.getElementById('logResultCount');
    const logDateRangeEl = document.getElementById('logDateRange');

    function escapeHtml(value) {
      return String(value ?? '')
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
    }

    function renderResults(payload) {
      const results = payload.results ?? [];
      resultCountEl.textContent = `${results.length} result(s)`;

      if (!results.length) {
        resultsEl.className = 'empty';
        resultsEl.innerHTML = 'No matches found.';
        return;
      }

      resultsEl.className = '';
      resultsEl.innerHTML = results.map(item => {
        const badges = [
          item.lastWriteUtc ? `<span class="badge">${new Date(item.lastWriteUtc).toLocaleString()}</span>` : '',
          item.extension ? `<span class="badge">${escapeHtml(item.extension || '[none]')}</span>` : '<span class="badge">[no extension]</span>',
          `<span class="badge">${Number(item.length || 0).toLocaleString()} bytes</span>`,
          item.matchAreas ? `<span class="badge">${escapeHtml(item.matchAreas)}</span>` : ''
        ].join('');
        const fileUrl = item.fullPath ? 'file:///' + item.fullPath.replace(/\\/g, '/') : '';
        const actionButton = item.canCopyToCovDrop
          ? `<div class="result-actions"><button class="action-btn" type="button" data-action="copy-cov-drop" data-path="${escapeHtml(item.fullPath)}" data-target="${escapeHtml(item.suggestedTargetPath || '')}">Copy As .cov</button></div>`
          : '';
        const openFolderBtn = item.fullPath
          ? `<button class="action-btn" type="button" data-action="open-folder" data-path="${escapeHtml(item.fullPath)}" title="Open folder in Explorer">Open Folder</button>`
          : '';

        return `
          <article class="result">
            <div class="result-header">
              <h3><a href="${fileUrl}" target="_blank" title="Open file">${escapeHtml(item.name)}</a></h3>
              <span class="result-date">${item.lastWriteUtc ? new Date(item.lastWriteUtc).toLocaleString() : ''}</span>
            </div>
            <div class="path">${escapeHtml(item.relativePath)}</div>
            <div class="path">${escapeHtml(item.fullPath)}</div>
            <div class="badges">${badges}</div>
            <div class="snippet">${escapeHtml(item.snippet || '')}</div>
            <div class="result-actions">${openFolderBtn}${actionButton}</div>
          </article>
        `;
      }).join('');
    }

    async function copyToCovDrop(sourcePath, targetPath) {
      const message = `Copy this .bak file to the parent folder as .cov?\n\nSource:\n${sourcePath}\n\nTarget:\n${targetPath}`;
      if (!window.confirm(message)) {
        return;
      }

      statusEl.textContent = 'Copying file...';

      try {
        const response = await fetch('/api/copy-to-cov', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ sourcePath })
        });
        const payload = await response.json();

        if (!response.ok || payload.error) {
          throw new Error(payload.error || 'Copy failed.');
        }

        statusEl.textContent = `Copied to ${payload.targetPath}`;
      } catch (error) {
        statusEl.textContent = error.message;
      }
    }

    async function runSearch() {
      const query = document.getElementById('query').value.trim();
      const mode = document.getElementById('mode').value;
      const limit = document.getElementById('limit').value || '100';
      const url = `/api/search?query=${encodeURIComponent(query)}&mode=${encodeURIComponent(mode)}&limit=${encodeURIComponent(limit)}`;

      statusEl.textContent = 'Searching...';
      resultsEl.classList.add('loading');

      try {
        const response = await fetch(url);
        const payload = await response.json();
        
        if (payload.searchTimeout) {
          resultsEl.className = 'empty';
          if (payload.isPartial) {
            resultsEl.innerHTML = '<div class="alert-error">Search timed out (30s). Try narrowing your search or reducing the limit.</div>';
            statusEl.textContent = 'Search timed out - partial results may be available.';
          } else {
            resultsEl.innerHTML = '<div class="alert-error">Search timed out. Try narrowing your search or reducing the limit.</div>';
            statusEl.textContent = 'Search timed out after 30 seconds.';
          }
        } else {
          renderResults(payload);
          const partialNote = payload.isPartial ? ' (partial - timed out)' : '';
          statusEl.textContent = `Search finished in ${payload.elapsedMs} ms${partialNote}.`;
        }
      } catch (error) {
        resultsEl.className = 'empty';
        resultsEl.textContent = error.message;
        statusEl.textContent = 'Search failed.';
      } finally {
        resultsEl.classList.remove('loading');
      }
    }

    async function rebuildIndex() {
      console.log('rebuildIndex called');
      statusEl.textContent = 'Indexing started...';
      reindexBtn.disabled = true;
      
      indexingOverlay.classList.add('active');
      indexingStatus.textContent = 'Starting indexer...';
      indexingProgressBar.style.width = '0%';
      indexingCount.textContent = '0 files processed';
      indexingLog.textContent = '';

      try {
        console.log('Calling /api/reindex...');
        const response = await fetch('/api/reindex', { method: 'POST' });
        console.log('Response status:', response.status);
        const payload = await response.json();
        console.log('Payload:', payload);
        
        if (payload.status === 'started') {
          statusEl.textContent = 'Indexing in progress...';
          indexingStatus.textContent = 'Indexing files...';
          pollIndexStatus();
        } else {
          indexingOverlay.classList.remove('active');
          if (payload.error) {
            statusEl.textContent = 'Error: ' + payload.error;
            indexingStatus.textContent = 'Error: ' + payload.error;
          } else {
            statusEl.textContent = 'Index ready.';
            indexingStatus.textContent = 'Index ready.';
          }
          reindexBtn.disabled = false;
        }
      } catch (error) {
        console.error('rebuildIndex error:', error);
        indexingOverlay.classList.remove('active');
        statusEl.textContent = 'Refresh failed: ' + error.message;
        reindexBtn.disabled = false;
      }
    }

    async function pollIndexStatus() {
      try {
        const response = await fetch('/api/index-status');
        const status = await response.json();
        
        if (status.indexingInProgress) {
          const count = status.currentFileCount || 0;
          const totalFiles = status.estimatedTotal || 42500;
          const progress = Math.min(95, Math.round((count / totalFiles) * 100));
          const logLines = status.lastLogLines || [];
          const lastLine = logLines[logLines.length - 1] || '';
          
          indexingStatus.textContent = 'Indexing files...';
          indexingProgressBar.style.width = progress + '%';
          indexingCount.textContent = count.toLocaleString() + ' files processed';
          indexingLog.textContent = lastLine;
          statusEl.textContent = `Indexing... ${count} files`;
          
          setTimeout(pollIndexStatus, 2000);
        } else {
          indexingProgressBar.style.width = '100%';
          indexingCount.textContent = (status.currentFileCount || 0).toLocaleString() + ' files indexed';
          indexingStatus.textContent = 'Index complete!';
          
          if (status.needsReload) {
            setTimeout(() => {
              indexingOverlay.classList.remove('active');
              location.reload();
            }, 1000);
          } else {
            setTimeout(() => {
              indexingOverlay.classList.remove('active');
              statusEl.textContent = 'Index ready. ' + (status.currentFileCount || 0) + ' files.';
              reindexBtn.disabled = false;
            }, 1500);
          }
        }
      } catch (error) {
        indexingLog.textContent = 'Status check failed';
        setTimeout(pollIndexStatus, 3000);
      }
    }

    function renderLogResults(payload) {
      const results = payload.results ?? [];

      if (payload.error) {
        logResultCountEl.textContent = 'Error';
        logResultsEl.className = 'empty';
        logResultsEl.innerHTML = `<div class="alert-error">${escapeHtml(payload.error)}</div>`;
        logDateRangeEl.textContent = payload.totalEntries > 0 ? `${payload.totalEntries} entries` : '';
        return;
      }

      const totalMatches = payload.totalMatches || results.length;
      const hasMore = payload.hasMore === true;
      if (hasMore) {
        logResultCountEl.textContent = `Showing first ${results.length} matches (more available)`;
      } else {
        logResultCountEl.textContent = `Showing ${results.length} of ${totalMatches} match${totalMatches !== 1 ? 'es' : ''}`;
      }

      let dateRangeText = `${payload.totalEntries || 0} entries`;
      if (payload.minDate && payload.maxDate) {
        dateRangeText += ` | ${payload.minDate} to ${payload.maxDate}`;
      }
      logDateRangeEl.textContent = dateRangeText;

      if (!results.length) {
        logResultsEl.className = 'empty';
        if (payload.query) {
          logResultsEl.innerHTML = 'No matching entries.';
        } else {
          logResultsEl.innerHTML = 'Enter a query to search.';
        }
        return;
      }

      logResultsEl.className = '';
      logResultsEl.innerHTML = results.map(item => {
        const p = item.parsed || {};
        const cust = p.Cust ?? p.cust ?? '';
        const shipTo = p.ShipTo ?? p.shipto ?? '';
        const trxType = p.TrxType ?? p.trxtype ?? '';
        const pid = p.PID ?? p.pid ?? '';
        const dateTime = p.DateTime ?? p.datetime ?? '';
        const message = p.Msg ?? p.msg ?? '';
        const metaBadges = [];
        if (cust) metaBadges.push(`<span class="badge">${escapeHtml(cust)}</span>`);
        if (shipTo && shipTo !== '0') metaBadges.push(`<span class="badge">ShipTo: ${escapeHtml(shipTo)}</span>`);
        if (trxType) metaBadges.push(`<span class="badge">${escapeHtml(trxType)}</span>`);
        if (pid) metaBadges.push(`<span class="badge">PID: ${escapeHtml(pid)}</span>`);

        let contextHtml = '';
        if (item.contextLines && item.contextLines.length > 0) {
          contextHtml = item.contextLines.map(ctx => {
            const lineClass = ctx.isMatch ? 'log-line match' : 'log-line';
            const lineContent = escapeHtml(ctx.parsed?.FullLine ?? ctx.parsed?.fullLine ?? '');
            const lineNum = ctx.lineNumber;
            const prefix = ctx.isMatch ? '>>>' : '   ';
            return `<div class="${lineClass}"><span class="line-num">${lineNum}</span><span class="line-prefix">${prefix}</span><span class="line-content">${lineContent}</span></div>`;
          }).join('');
        }

        return `
          <article class="result log-hit">
            <div class="result-header">
              <span class="badge badge-primary">Line ${item.lineNumber}</span>
              <span class="badge badge-success">${escapeHtml(dateTime)}</span>
            </div>
            <div class="badges">${metaBadges.join('')}</div>
            ${message ? `<div class="log-message">${escapeHtml(message)}</div>` : ''}
            <div class="log-context">${contextHtml}</div>
          </article>
        `;
      }).join('');
    }

    async function runLogSearch() {
      const query = document.getElementById('logQuery').value.trim();
      const limit = document.getElementById('logLimit').value || '50';
      const startDate = document.getElementById('logStartDate').value;
      const endDate = document.getElementById('logEndDate').value;

      let url = `/api/search-log?query=${encodeURIComponent(query)}&limit=${encodeURIComponent(limit)}`;
      if (startDate) url += `&startDate=${encodeURIComponent(startDate)}`;
      if (endDate) url += `&endDate=${encodeURIComponent(endDate)}`;

      statusEl.textContent = 'Searching log...';
      logResultsEl.classList.add('loading');

      try {
        const response = await fetch(url);
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        const payload = await response.json();
        renderLogResults(payload);
        if (payload.cacheHit) {
          statusEl.textContent = payload.cacheSource === 'prefix'
            ? 'Log search returned from prefix cache.'
            : 'Log search returned from cache.';
        } else {
          statusEl.textContent = `Log search finished in ${payload.elapsedMs} ms.`;
        }
      } catch (error) {
        logResultsEl.className = 'empty';
        logResultsEl.innerHTML = `<div class="alert-error">Fetch error: ${escapeHtml(error.message)}</div>`;
        statusEl.textContent = 'Log search failed: ' + error.message;
        console.error('Log search error:', error);
      } finally {
        logResultsEl.classList.remove('loading');
      }
    }

    const refreshErrorsBtn = document.getElementById('refreshErrorsBtn');
    const errorSummaryContent = document.getElementById('errorSummaryContent');

    function renderErrorSummary(data) {
      if (!data || !data.summary) {
        errorSummaryContent.innerHTML = '<div class="empty">Error summary unavailable.</div>';
        return;
      }

      const summary = data.summary;
      if (!summary.hasErrors) {
        errorSummaryContent.innerHTML = `<div class="status-text" style="background: var(--brand-soft); border-color: #caeccc;">No errors found in last ${summary.daysBack || 30} days</div>`;
        return;
      }

      let html = '<div style="display: grid; gap: 16px;">';
      const categoryIcons = {
        'Invalid Paper Code': '📄',
        'Invalid Vendor': '🏢',
        'PO Receipt - Line Not Exist': '📋',
        'Other Errors': '⚠️'
      };

      for (const [category, items] of Object.entries(summary.categories)) {
        const icon = categoryIcons[category] || '🔴';
        html += `<div style="background: #fff; border: 1px solid var(--line); border-radius: 12px; padding: 14px;">
          <h4 style="margin: 0 0 10px; color: var(--danger); display: flex; align-items: center; gap: 8px;">
            <span>${icon}</span> ${category} <span style="font-weight: normal; color: var(--muted); font-size: 0.85rem;">(${items.length} occurrences)</span>
          </h4>
          <div style="display: flex; flex-wrap: wrap; gap: 6px;">`;
        
        for (const item of items.slice(0, 20)) {
          html += `<span class="badge" style="background: var(--danger-soft); color: var(--danger); border-color: #f1d1c8;">${escapeHtml(item)}</span>`;
        }
        if (items.length > 20) {
          html += `<span class="badge" style="background: #f0f0f0; color: #666;">+${items.length - 20} more</span>`;
        }
        html += '</div></div>';
      }
      html += '</div>';
      errorSummaryContent.innerHTML = html;
    }

    async function loadErrorSummary() {
      try {
        const response = await fetch('/api/error-summary');
        const data = await response.json();
        renderErrorSummary(data);
      } catch (error) {
        errorSummaryContent.innerHTML = `<div class="alert-error">Failed to load error summary: ${escapeHtml(error.message)}</div>`;
      }
    }

    refreshErrorsBtn.addEventListener('click', loadErrorSummary);
    loadErrorSummary();

    form.addEventListener('submit', event => {
      event.preventDefault();
      runSearch();
    });

    reindexBtn.addEventListener('click', () => {
      console.log('Refresh button clicked');
      statusEl.textContent = 'Indexing started...';
      fetch('/api/reindex', { method: 'POST' })
        .then(r => r.json())
        .then(d => {
          console.log('reindex response:', d);
          statusEl.textContent = d.message || 'Indexing started';
          if (d.status === 'started') {
            indexingOverlay.classList.add('active');
            pollIndexStatus();
          }
        })
        .catch(e => {
          console.error('reindex error:', e);
          statusEl.textContent = 'Error: ' + e.message;
        });
    });

    logSearchForm.addEventListener('submit', event => {
      event.preventDefault();
      runLogSearch();
    });

    resultsEl.addEventListener('click', event => {
      const copyBtn = event.target.closest('[data-action="copy-cov-drop"]');
      if (copyBtn) {
        copyToCovDrop(copyBtn.dataset.path, copyBtn.dataset.target || '');
        return;
      }
      
      const folderBtn = event.target.closest('[data-action="open-folder"]');
      if (folderBtn) {
        event.preventDefault();
        const folderPath = folderBtn.dataset.path;
        // Call the API to open the folder via PowerShell/Explorer
        fetch('/api/open-folder?path=' + encodeURIComponent(folderPath))
          .then(r => r.json())
          .then(data => {
            if (data.error) {
              alert('Failed to open folder: ' + data.error);
            }
          })
          .catch(err => {
            alert('Error: ' + err.message);
          });
      }
    });
  </script>
</body>
</html>
'@

    return $template.
        Replace("__SHARE_PATH__", (Escape-Html $Index.SharePath)).
        Replace("__FILE_COUNT__", [string]$Index.FileCount).
        Replace("__GENERATED_UTC__", (Escape-Html $Index.GeneratedUtc))
}

$script:IndexingInProgress = $false
$script:IndexNeedsReload = $false

$script:Index = $null
$script:LogIndex = $null
$script:LogErrorSummary = $null
$script:LogSearchCache = @{}
$script:LogSearchCacheOrder = New-Object System.Collections.Generic.List[string]
$script:LogSearchCacheTtlSeconds = 180
$script:LogSearchCacheMaxEntries = 120
$listener = $null
$errorLogPath = Join-Path (Split-Path -Parent $IndexPath) "edi-search-error.log"
$script:ShutdownEvent = New-Object System.Threading.ManualResetEvent($false)
$script:ExitEvent = $null

$script:ExitEvent = Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -Action {
    $script:ShutdownEvent.Set() | Out-Null
}

try {
    Write-Host "Initializing EDI Search..." -ForegroundColor Cyan
    Write-Host "SharePath: $SharePath"
    Write-Host "IndexPath: $IndexPath"
    Write-Host "NoIndex: $NoIndex"

    if ($NoIndex) {
        Write-Host ""
        Write-Host "Skipping file indexing (NoIndex mode)..." -ForegroundColor Yellow
        $script:Index = @{
            SharePath = $SharePath
            FileCount = 0
            GeneratedUtc = (Get-Date).ToUniversalTime().ToString("o")
            Items = @()
        }
    }
    else {
        Write-Host "Loading EDI index and refreshing changed files..." -ForegroundColor Cyan
        $script:Index = Load-EdiIndex -SourcePath $SharePath -OutputPath $IndexPath -Parallelism $ThrottleLimit -MaxContentSize $MaxContentSize -ForceRebuild:$Rebuild
        Write-Host ("EDI Index ready. Files: {0}" -f $script:Index.FileCount)
    }

    Write-Host ""
    if ($NoLogIndex) {
        Write-Host "Skipping log indexing (NoLogIndex mode)..." -ForegroundColor Yellow
        $script:LogIndex = @{
            Entries = @()
            TotalLines = 0
        }
    }
    else {
        Write-Host "Building log file index..." -ForegroundColor Cyan
        $script:LogIndex = Build-LogIndex -LogPath $LogPath -ThrottleLimit $ThrottleLimit
        Write-Host ("Log Index ready. Entries: {0}" -f $script:LogIndex.Entries.Count)

        Write-Host ""
        Write-Host "Analyzing recent errors..." -ForegroundColor Cyan
        $script:LogErrorSummary = Get-LogErrorSummary -LogIndex $script:LogIndex -DaysBack 30
        if ($script:LogErrorSummary.hasErrors) {
            $totalErrors = ($script:LogErrorSummary.categories.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
            Write-Host ("Found {0} errors across {1} categories" -f $totalErrors, $script:LogErrorSummary.totalCategories) -ForegroundColor Yellow
        } else {
            Write-Host "No errors found in last 30 days." -ForegroundColor Green
        }
    }

    Write-Host ""
    Write-Host "Creating HTTP listener..." -ForegroundColor Cyan
    
    $availablePort = 8787
    for ($i = 0; $i -lt 10; $i++) {
        $testPort = $availablePort + $i
        try {
            $testListener = New-Object System.Net.HttpListener
            $testListener.Prefixes.Add("http://localhost:$testPort/")
            $testListener.Start()
            $testListener.Stop()
            $testListener.Close()
            $availablePort = $testPort
            Write-Host "Using port: $availablePort" -ForegroundColor Gray
            break
        }
        catch {
            Write-Host "Port $testPort in use..." -ForegroundColor Gray
            continue
        }
    }
    
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$availablePort/")
    
    try {
        $listener.Start()
    }
    catch {
        Write-Host "Failed to start listener: $_" -ForegroundColor Red
        Write-Host "Attempting to free port $availablePort..." -ForegroundColor Yellow
        $conn = Get-NetTCPConnection -LocalPort $availablePort -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($conn) {
            Write-Host "Found process $($conn.OwningProcess) on port $availablePort" -ForegroundColor Yellow
            try {
                Stop-Process -Id $conn.OwningProcess -Force -ErrorAction Stop
                Start-Sleep -Seconds 1
                $listener.Start()
                Write-Host "Listener started after cleanup" -ForegroundColor Green
            }
            catch {
                Write-Host "Could not stop process: $_" -ForegroundColor Red
                throw
            }
        }
        else {
            throw
        }
    }
    
    $baseUrl = "http://localhost:$availablePort/"
    Write-Host "EDI Search running at $baseUrl" -ForegroundColor Green
    Write-Host "Press Ctrl+C to stop, or close this window."
    Start-Process -FilePath $baseUrl

    while (-not $script:ShutdownEvent.WaitOne(100)) {
        if (-not $listener.IsListening) { break }
        try {
            $context = $listener.GetContext()
        }
        catch [System.Net.HttpListenerException] {
            if ($_.Exception.NativeErrorCode -eq 995) { break }
            Write-Host "Listener exception: $($_.Exception.Message)" -ForegroundColor Yellow
            continue
        }
        catch {
            if ($script:ShutdownEvent.WaitOne(0)) { break }
            Write-Host "Unexpected exception: $($_.Exception.Message)" -ForegroundColor Red
            continue
        }

        $request = $context.Request
        $response = $context.Response

        try {
            switch ($request.Url.AbsolutePath) {
                "/" {
                    Write-TextResponse -Response $response -Body (Get-PageHtml -Index $script:Index)
                }
                "/api/search" {
                    $query = $request.QueryString["query"]
                    $mode = $request.QueryString["mode"]
                    $limitRaw = $request.QueryString["limit"]
                    $limit = $MaxResults

                    if (-not [string]::IsNullOrWhiteSpace($limitRaw)) {
                        [void][int]::TryParse($limitRaw, [ref]$limit)
                    }

                    if ($limit -lt 1) { $limit = 1 }
                    if ($limit -gt 500) { $limit = 500 }

                    if ($script:IndexNeedsReload -and -not $script:IndexingInProgress) {
                        try {
                            if (Test-Path $IndexPath) {
                                $script:Index = Get-Content $IndexPath -Raw | ConvertFrom-Json
                                $script:IndexNeedsReload = $false
                                Write-Host "Index reloaded: $($script:Index.FileCount) files" -ForegroundColor Green
                            }
                        } catch {
                            Write-Warning "Could not reload index: $_"
                        }
                    }

                    $timeoutMs = 30000
                    $sw = [System.Diagnostics.Stopwatch]::StartNew()
                    
                    $runspace = [PowerShell]::Create()
                    [void]$runspace.AddScript({
                        param($Index, $Query, $Mode, $Limit)
                        try {
                            Search-EdiIndex -Index $Index -Query $Query -Mode $Mode -Limit $Limit
                        } catch {
                            Write-Host "Search error: $_"
                            @()
                        }
                    })
                    [void]$runspace.AddParameter("Index", $script:Index)
                    [void]$runspace.AddParameter("Query", $query)
                    [void]$runspace.AddParameter("Mode", $mode)
                    [void]$runspace.AddParameter("Limit", $limit)
                    
                    $handle = $runspace.BeginInvoke()
                    $waitResult = [System.Threading.WaitHandle]::WaitAny(@($handle.AsyncWaitHandle, $script:ShutdownEvent), $timeoutMs)
                    
                    $results = $null
                    $timedOut = $false
                    $isPartial = $false
                    
                    if ($waitResult -eq [System.Threading.WaitHandle]::WaitTimeout) {
                        $runspace.Stop()
                        $timedOut = $true
                        $results = @()
                        $isPartial = $true
                    } else {
                        $results = $runspace.EndInvoke($handle)
                    }
                    $runspace.Dispose()
                    
                    $sw.Stop()

                    Write-JsonResponse -Response $response -Payload @{
                        query = $query
                        mode = $mode
                        limit = $limit
                        elapsedMs = $sw.ElapsedMilliseconds
                        searchTimeout = $timedOut
                        isPartial = $isPartial
                        summary = Get-IndexSummary -Index $script:Index
                        results = @($results | ForEach-Object {
                            @{
                                name = $_.Name
                                relativePath = $_.RelativePath
                                fullPath = $_.FullPath
                                extension = $_.Extension
                                length = $_.Length
                                lastWriteUtc = $_.LastWriteUtc
                                matchAreas = $_.MatchAreas
                                snippet = $_.Snippet
                                canCopyToCovDrop = $_.Extension -eq ".bak" -and $_.FullPath.IndexOf("\Processed\", [System.StringComparison]::OrdinalIgnoreCase) -ge 0
                                suggestedTargetPath = if ($_.Extension -eq ".bak" -and $_.FullPath.IndexOf("\Processed\", [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { Get-CovDropTargetPath -SourcePath $_.FullPath } else { "" }
                            }
                        })
                    }
                    break
                }
                "/api/search-log" {
                    $query = $request.QueryString["query"]
                    $limitRaw = $request.QueryString["limit"]
                    $startDateStr = $request.QueryString["startDate"]
                    $endDateStr = $request.QueryString["endDate"]
                    $limit = 50

                    if (-not [string]::IsNullOrWhiteSpace($limitRaw)) {
                        [void][int]::TryParse($limitRaw, [ref]$limit)
                    }

                    if ($limit -lt 1) { $limit = 1 }
                    if ($limit -gt 200) { $limit = 200 }

                    $startDate = $null
                    $endDate = $null

                    if (-not [string]::IsNullOrWhiteSpace($startDateStr)) {
                        try { $startDate = [DateTime]::Parse($startDateStr) } catch {}
                    }

                    if (-not [string]::IsNullOrWhiteSpace($endDateStr)) {
                        try { $endDate = [DateTime]::Parse($endDateStr).AddDays(1).AddSeconds(-1) } catch {}
                    }

                    $cacheKey = New-LogSearchCacheKey -Query $query -StartDate $startDateStr -EndDate $endDateStr -Limit $limit -LogIndex $script:LogIndex
                    $cachedPayload = Get-LogSearchCacheEntry -Key $cacheKey
                    if ($null -ne $cachedPayload) {
                        $cachedResponse = @{} + $cachedPayload
                        $cachedResponse.elapsedMs = 0
                        $cachedResponse.cacheHit = $true
                        $cachedResponse.cacheSource = "exact"
                        Write-JsonResponse -Response $response -Payload $cachedResponse
                        break
                    }

                    $prefixPayload = Get-LogSearchPrefixReusePayload -Query $query -StartDate $startDateStr -EndDate $endDateStr -Limit $limit -LogIndex $script:LogIndex
                    if ($null -ne $prefixPayload) {
                        Set-LogSearchCacheEntry -Key $cacheKey -Payload $prefixPayload -Query $query -StartDate $startDateStr -EndDate $endDateStr -Limit $limit -LogIndex $script:LogIndex
                        Write-JsonResponse -Response $response -Payload $prefixPayload
                        break
                    }
                    
                    $sw = [System.Diagnostics.Stopwatch]::StartNew()
                    $logResult = Search-LogIndex -LogIndex $script:LogIndex -Query $query -Limit $limit -StartDate $startDate -EndDate $endDate
                    $sw.Stop()

                    $payload = @{
                        query = $query
                        limit = $limit
                        startDate = $startDateStr
                        endDate = $endDateStr
                        elapsedMs = $sw.ElapsedMilliseconds
                        cacheHit = $false
                        logPath = $LogPath
                        totalEntries = $logResult.TotalEntries
                        totalMatches = $logResult.TotalMatches
                        hasMore = $logResult.HasMore
                        minDate = $logResult.MinDate
                        maxDate = $logResult.MaxDate
                        error = $logResult.Error
                        results = @($logResult.Results | ForEach-Object {
                            @{
                                lineNumber = $_.LineNumber
                                parsed = $_.Parsed
                                contextLines = $_.ContextLines
                            }
                        })
                    }

                    Set-LogSearchCacheEntry -Key $cacheKey -Payload $payload -Query $query -StartDate $startDateStr -EndDate $endDateStr -Limit $limit -LogIndex $script:LogIndex
                    Write-JsonResponse -Response $response -Payload $payload
                    break
                }
                "/api/copy-to-cov" {
                    if ($request.HttpMethod -ne "POST") {
                        Write-JsonResponse -Response $response -Payload @{ error = "Method not allowed." } -StatusCode 405
                        break
                    }

                    $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
                    $body = $reader.ReadToEnd()
                    $reader.Close()

                    $payload = if ([string]::IsNullOrWhiteSpace($body)) { @{} } else { $body | ConvertFrom-Json }
                    $copyResult = Copy-BakToCovDrop -SourcePath $payload.sourcePath

                    Write-JsonResponse -Response $response -Payload @{
                        ok = $true
                        sourcePath = $copyResult.sourcePath
                        targetPath = $copyResult.targetPath
                    }
                    break
                }
                "/api/reindex" {
                    if ($request.HttpMethod -ne "POST") {
                        Write-JsonResponse -Response $response -Payload @{ error = "Method not allowed." } -StatusCode 405
                        break
                    }

                    if ($script:IndexingInProgress) {
                        Write-JsonResponse -Response $response -Payload @{
                            ok = $false
                            error = "Indexing already in progress"
                        } -StatusCode 409
                        break
                    }

                    $script:IndexingInProgress = $true
                    
                    $indexerScript = Join-Path (Split-Path -Parent $PSCommandPath) "Index-EdiFiles.ps1"
                    $logPath = Join-Path (Split-Path -Parent $IndexPath) "indexer.log"
                    
                    Write-Host "Starting reindex job..." -ForegroundColor Cyan
                    $job = Start-Job -ScriptBlock {
                        param($Script, $SharePath, $IndexPath, $MaxContentSize, $LogPath)
                        $output = & $Script -SharePath $SharePath -OutputPath $IndexPath -MaxContentSize $MaxContentSize 2>&1
                        $output | Out-File -FilePath $LogPath -Encoding UTF8
                    } -ArgumentList $indexerScript, $SharePath, $IndexPath, $MaxContentSize, $logPath
                    
                    $script:IndexingJobId = $job.Id
                    
                    Write-JsonResponse -Response $response -Payload @{
                        ok = $true
                        status = "started"
                        message = "Indexing started"
                        jobId = $job.Id
                    }
                    break
                }
                "/api/index-status" {
                    $indexerScript = Join-Path (Split-Path -Parent $PSCommandPath) "Index-EdiFiles.ps1"
                    $logPath = Join-Path (Split-Path -Parent $IndexPath) "indexer.log"
                    $jsonPath = $IndexPath
                    
                    if ($script:IndexingJobId) {
                        $job = Get-Job -Id $script:IndexingJobId -ErrorAction SilentlyContinue
                        if ($job -and $job.State -ne 'Running') {
                            $script:IndexingInProgress = $false
                            $script:IndexNeedsReload = $true
                            $script:IndexingJobId = $null
                        }
                    }
                    
                    $status = @{
                        indexingInProgress = $script:IndexingInProgress
                        currentFileCount = 0
                        estimatedTotal = 42500
                        needsReload = if ($script:IndexNeedsReload) { $true } else { $false }
                    }
                    
                    if (Test-Path $jsonPath) {
                        try {
                            $currentIndex = Get-Content $jsonPath -Raw | ConvertFrom-Json
                            $status.currentFileCount = $currentIndex.FileCount
                            $status.generatedUtc = $currentIndex.GeneratedUtc
                        } catch {}
                    }
                    
                    if (Test-Path $logPath) {
                        $status.lastLogLines = @(Get-Content $logPath -Tail 10)
                    }
                    
                    Write-JsonResponse -Response $response -Payload $status
                    break
                }
                "/api/open-folder" {
                    $folderPath = $request.QueryString["path"]
                    if ([string]::IsNullOrWhiteSpace($folderPath)) {
                        Write-JsonResponse -Response $response -Payload @{ error = "Path parameter required." } -StatusCode 400
                        break
                    }

                    try {
                        Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$folderPath`""
                        Write-JsonResponse -Response $response -Payload @{ ok = $true }
                    } catch {
                        Write-JsonResponse -Response $response -Payload @{ error = $_.Exception.Message } -StatusCode 500
                    }
                }
                "/api/error-summary" {
                    Write-JsonResponse -Response $response -Payload @{
                        ok = $true
                        summary = $script:LogErrorSummary
                    }
                    break
                }
                default {
                    Write-TextResponse -Response $response -Body "Not found." -ContentType "text/plain; charset=utf-8" -StatusCode 404
                }
            }
        }
        catch {
            try {
                if ($null -ne $response) {
                    Write-JsonResponse -Response $response -Payload @{ error = $_.Exception.Message } -StatusCode 500
                }
            } catch {}
        }
    }
}
catch {
    Ensure-ParentDirectory -Path $errorLogPath

    $errorLines = @(
        ("Timestamp: {0}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
        ("Message: {0}" -f $_.Exception.Message)
        ("ScriptStackTrace: {0}" -f $_.ScriptStackTrace)
        ("Invocation: {0}" -f $_.InvocationInfo.PositionMessage)
        ""
    )

    Add-Content -LiteralPath $errorLogPath -Value $errorLines

    Write-Host ""
    Write-Host "EDI Search crashed." -ForegroundColor Red
    Write-Host ("Error: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-Host ("Details written to: {0}" -f $errorLogPath) -ForegroundColor Yellow
}
finally {
    if ($null -ne $script:ExitEvent) {
        try {
            Unregister-Event -SubscriptionId $script:ExitEvent.Id -ErrorAction SilentlyContinue
        } catch {}
    }
    if ($null -ne $script:ShutdownEvent) {
        $script:ShutdownEvent.Set() | Out-Null
        $script:ShutdownEvent.Dispose() | Out-Null
    }
    if ($null -ne $listener -and $listener.IsListening) {
        $listener.Stop()
    }
    if ($null -ne $listener) {
        $listener.Close()
    }
    Write-Host "EDI Search stopped." -ForegroundColor Yellow
}

# End of script
