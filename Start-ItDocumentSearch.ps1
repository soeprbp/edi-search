param(
    [string]$SourcePath = "I:\IT",
    [string]$IndexPath = (Join-Path $PSScriptRoot "data\it-document-index.clixml"),
    [string]$Prefix = "http://localhost:8788/",
    [int]$MaxResults = 200,
    [int]$ThrottleLimit = [Math]::Max(2, [Environment]::ProcessorCount),
    [int]$MaxPlainTextFileSizeMB = 8,
    [switch]$Rebuild
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:SupportedExtensions = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
@(
    ".txt", ".log", ".csv", ".tsv", ".md", ".rtf",
    ".json", ".xml", ".html", ".htm", ".yaml", ".yml", ".ini", ".config",
    ".ps1", ".psm1", ".psd1", ".sql", ".cmd", ".bat",
    ".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm",
    ".pdf"
) | ForEach-Object { [void]$script:SupportedExtensions.Add($_) }

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
        [int]$MaxLength = 260
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }

    $clean = ($Text -replace "\s+", " ").Trim()
    if ($clean.Length -le $MaxLength) { return $clean }
    return $clean.Substring(0, $MaxLength) + "..."
}

function Get-Snippet {
    param(
        [AllowNull()][string]$Content,
        [AllowNull()][string]$Query
    )

    if ([string]::IsNullOrWhiteSpace($Content)) { return "" }

    $normalized = ($Content -replace "\s+", " ").Trim()
    if ([string]::IsNullOrWhiteSpace($Query)) {
        return Trim-Text -Text $normalized
    }

    $index = $normalized.IndexOf($Query, [System.StringComparison]::OrdinalIgnoreCase)
    if ($index -lt 0) {
        return Trim-Text -Text $normalized
    }

    $start = [Math]::Max(0, $index - 100)
    $length = [Math]::Min(260, $normalized.Length - $start)
    $snippet = $normalized.Substring($start, $length)
    if ($start -gt 0) { $snippet = "..." + $snippet }
    if (($start + $length) -lt $normalized.Length) { $snippet += "..." }
    return $snippet
}

function Get-RelativePathSafe {
    param(
        [string]$RootPath,
        [string]$FullPath
    )

    if ($FullPath.StartsWith($RootPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $FullPath.Substring($RootPath.Length).TrimStart("\")
    }

    return $FullPath
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

function Test-SourceAccessible {
    param([string]$Path)

    try {
        return Test-Path -LiteralPath $Path
    }
    catch {
        return $false
    }
}

function Get-FullPathSafe {
    param([string]$Path)

    try {
        return [System.IO.Path]::GetFullPath($Path)
    }
    catch {
        return $Path
    }
}

function Test-PathUnderRoot {
    param(
        [string]$Path,
        [string]$RootPath
    )

    $fullPath = Get-FullPathSafe -Path $Path
    $fullRoot = (Get-FullPathSafe -Path $RootPath).TrimEnd('\')

    return $fullPath.StartsWith(($fullRoot + "\"), [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.Equals($fullRoot, [System.StringComparison]::OrdinalIgnoreCase)
}

function Assert-PathOutsideSource {
    param(
        [string]$Path,
        [string]$SourcePath,
        [string]$Label
    )

    if (Test-PathUnderRoot -Path $Path -RootPath $SourcePath) {
        throw "$Label must not be inside the source path. This tool is read-only against $SourcePath."
    }
}

function Test-SupportedDocument {
    param([System.IO.FileInfo]$File)

    return $script:SupportedExtensions.Contains($File.Extension)
}

function Get-Latin1Encoding {
    return [System.Text.Encoding]::GetEncoding("iso-8859-1")
}

function Read-PlainTextFile {
    param(
        [string]$Path,
        [int64]$MaxBytes
    )

    $fileInfo = Get-Item -LiteralPath $Path
    if ($fileInfo.Length -gt $MaxBytes) {
        return [pscustomobject]@{
            Content  = ""
            Strategy = "Skipped large text file"
            Error    = ("Plain-text file exceeds {0:N0} bytes." -f $MaxBytes)
        }
    }

    try {
        $content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        return [pscustomobject]@{
            Content  = $content
            Strategy = "Plain text"
            Error    = $null
        }
    }
    catch {
        return [pscustomobject]@{
            Content  = ""
            Strategy = "Plain text"
            Error    = $_.Exception.Message
        }
    }
}

function Read-ZipEntryText {
    param(
        [string]$ArchivePath,
        [string[]]$EntryNames
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
        foreach ($entryName in $EntryNames) {
            $entry = $archive.Entries | Where-Object { $_.FullName -eq $entryName } | Select-Object -First 1
            if ($null -eq $entry) { continue }

            $stream = $entry.Open()
            try {
                $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
                try {
                    return $reader.ReadToEnd()
                }
                finally {
                    $reader.Dispose()
                }
            }
            finally {
                $stream.Dispose()
            }
        }
    }
    finally {
        if ($null -ne $archive) {
            $archive.Dispose()
        }
    }

    return $null
}

function Convert-XmlToJoinedText {
    param([AllowNull()][string]$XmlText)

    if ([string]::IsNullOrWhiteSpace($XmlText)) {
        return ""
    }

    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $false
    $doc.LoadXml($XmlText)
    return ($doc.InnerText -replace "\s+", " ").Trim()
}

function Read-WordOpenXml {
    param([string]$Path)

    try {
        $parts = @(
            "word/document.xml",
            "word/header1.xml",
            "word/header2.xml",
            "word/header3.xml",
            "word/footer1.xml",
            "word/footer2.xml",
            "word/footer3.xml"
        )

        $chunks = foreach ($part in $parts) {
            $xml = Read-ZipEntryText -ArchivePath $Path -EntryNames @($part)
            if (-not [string]::IsNullOrWhiteSpace($xml)) {
                Convert-XmlToJoinedText -XmlText $xml
            }
        }

        return [pscustomobject]@{
            Content  = (($chunks | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join "`n").Trim()
            Strategy = "Open XML Word"
            Error    = $null
        }
    }
    catch {
        return [pscustomobject]@{
            Content  = ""
            Strategy = "Open XML Word"
            Error    = $_.Exception.Message
        }
    }
}

function Read-ExcelOpenXml {
    param([string]$Path)

    try {
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $sharedStrings = @{}
            $sharedStringsEntry = $archive.Entries | Where-Object { $_.FullName -eq "xl/sharedStrings.xml" } | Select-Object -First 1
            if ($null -ne $sharedStringsEntry) {
                $stream = $sharedStringsEntry.Open()
                try {
                    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
                    $xmlText = $reader.ReadToEnd()
                    $xml = [xml]$xmlText
                    $index = 0
                    foreach ($si in @($xml.sst.si)) {
                        $sharedStrings[$index] = (($si.InnerText -replace "\s+", " ").Trim())
                        $index++
                    }
                }
                finally {
                    if ($null -ne $reader) { $reader.Dispose() }
                    $stream.Dispose()
                }
            }

            $sheetEntries = @($archive.Entries | Where-Object { $_.FullName -like "xl/worksheets/*.xml" })
            $chunks = New-Object System.Collections.Generic.List[string]

            foreach ($sheetEntry in $sheetEntries) {
                $stream = $sheetEntry.Open()
                try {
                    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
                    $xml = [xml]$reader.ReadToEnd()

                    foreach ($cell in @($xml.worksheet.sheetData.row.c)) {
                        if ($cell.t -eq "s") {
                            $ref = [int]$cell.v
                            if ($sharedStrings.ContainsKey($ref)) {
                                $chunks.Add($sharedStrings[$ref])
                            }
                        }
                        elseif ($null -ne $cell.v) {
                            $chunks.Add([string]$cell.v)
                        }
                        elseif ($null -ne $cell.is) {
                            $chunks.Add(($cell.is.InnerText -replace "\s+", " ").Trim())
                        }
                    }
                }
                finally {
                    if ($null -ne $reader) { $reader.Dispose() }
                    $stream.Dispose()
                }
            }

            return [pscustomobject]@{
                Content  = (($chunks | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join "`n").Trim()
                Strategy = "Open XML Excel"
                Error    = $null
            }
        }
        finally {
            $archive.Dispose()
        }
    }
    catch {
        return [pscustomobject]@{
            Content  = ""
            Strategy = "Open XML Excel"
            Error    = $_.Exception.Message
        }
    }
}

function Read-PowerPointOpenXml {
    param([string]$Path)

    try {
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $slideEntries = @($archive.Entries | Where-Object { $_.FullName -like "ppt/slides/slide*.xml" } | Sort-Object FullName)
            $chunks = New-Object System.Collections.Generic.List[string]

            foreach ($slideEntry in $slideEntries) {
                $stream = $slideEntry.Open()
                try {
                    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
                    $xml = [xml]$reader.ReadToEnd()
                    $nodes = $xml.SelectNodes("//*[local-name()='t']")
                    foreach ($node in @($nodes)) {
                        $value = ($node.InnerText -replace "\s+", " ").Trim()
                        if (-not [string]::IsNullOrWhiteSpace($value)) {
                            $chunks.Add($value)
                        }
                    }
                }
                finally {
                    if ($null -ne $reader) { $reader.Dispose() }
                    $stream.Dispose()
                }
            }

            return [pscustomobject]@{
                Content  = (($chunks | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join "`n").Trim()
                Strategy = "Open XML PowerPoint"
                Error    = $null
            }
        }
        finally {
            $archive.Dispose()
        }
    }
    catch {
        return [pscustomobject]@{
            Content  = ""
            Strategy = "Open XML PowerPoint"
            Error    = $_.Exception.Message
        }
    }
}

function Inflate-PdfStream {
    param([byte[]]$Bytes)

    $input = $null
    $output = $null
    $deflate = $null
    try {
        $input = New-Object System.IO.MemoryStream(, $Bytes)
        $output = New-Object System.IO.MemoryStream
        $deflate = New-Object System.IO.Compression.DeflateStream($input, [System.IO.Compression.CompressionMode]::Decompress)
        $deflate.CopyTo($output)
        return $output.ToArray()
    }
    finally {
        if ($null -ne $deflate) { $deflate.Dispose() }
        if ($null -ne $output) { $output.Dispose() }
        if ($null -ne $input) { $input.Dispose() }
    }
}

function Decode-PdfLiteralString {
    param([string]$Value)

    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Value.Length; $i++) {
        $char = $Value[$i]
        if ($char -ne '\') {
            [void]$sb.Append($char)
            continue
        }

        if ($i -ge ($Value.Length - 1)) {
            break
        }

        $i++
        $next = $Value[$i]
        switch ($next) {
            'n' { [void]$sb.Append("`n") }
            'r' { [void]$sb.Append("`r") }
            't' { [void]$sb.Append("`t") }
            'b' { [void]$sb.Append([char]8) }
            'f' { [void]$sb.Append([char]12) }
            '(' { [void]$sb.Append('(') }
            ')' { [void]$sb.Append(')') }
            '\' { [void]$sb.Append('\') }
            default {
                if ($next -match '[0-7]') {
                    $octal = [string]$next
                    for ($j = 0; $j -lt 2 -and ($i + 1) -lt $Value.Length; $j++) {
                        if ($Value[$i + 1] -match '[0-7]') {
                            $i++
                            $octal += [string]$Value[$i]
                        }
                        else {
                            break
                        }
                    }
                    [void]$sb.Append([char][Convert]::ToInt32($octal, 8))
                }
                else {
                    [void]$sb.Append($next)
                }
            }
        }
    }

    return $sb.ToString()
}

function Extract-PdfTextFromContentStream {
    param([string]$StreamText)

    $chunks = New-Object System.Collections.Generic.List[string]

    foreach ($match in [regex]::Matches($StreamText, '\((?:\\.|[^\\)])*\)\s*Tj')) {
        $raw = $match.Value -replace '\)\s*Tj$', ''
        $raw = $raw.Substring(1)
        $chunks.Add((Decode-PdfLiteralString -Value $raw))
    }

    foreach ($match in [regex]::Matches($StreamText, '\[(?<body>.*?)\]\s*TJ', [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $parts = [regex]::Matches($match.Groups['body'].Value, '\((?:\\.|[^\\)])*\)')
        foreach ($part in $parts) {
            $raw = $part.Value.Substring(1, $part.Value.Length - 2)
            $chunks.Add((Decode-PdfLiteralString -Value $raw))
        }
    }

    foreach ($match in [regex]::Matches($StreamText, '\((?:\\.|[^\\)])*\)\s*[''"]')) {
        $raw = $match.Value -replace '\s*[''"]$', ''
        $raw = $raw.Substring(1, $raw.Length - 2)
        $chunks.Add((Decode-PdfLiteralString -Value $raw))
    }

    return (($chunks | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join " ").Trim()
}

function Read-PdfText {
    param([string]$Path)

    try {
        $latin1 = Get-Latin1Encoding
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        $text = $latin1.GetString($bytes)
        $chunks = New-Object System.Collections.Generic.List[string]

        foreach ($match in [regex]::Matches($text, '(?s)(?<dict><<.*?>>)\s*stream\r?\n')) {
            $streamStart = $match.Index + $match.Length
            $streamEnd = $text.IndexOf("endstream", $streamStart, [System.StringComparison]::Ordinal)
            if ($streamEnd -lt 0) { continue }

            $rawStreamText = $text.Substring($streamStart, $streamEnd - $streamStart)
            $rawBytes = $latin1.GetBytes($rawStreamText.Trim("`r", "`n"))
            $decodedBytes = $rawBytes

            if ($match.Groups["dict"].Value -match "/FlateDecode") {
                try {
                    $decodedBytes = Inflate-PdfStream -Bytes $rawBytes
                }
                catch {
                    continue
                }
            }

            $content = Extract-PdfTextFromContentStream -StreamText ($latin1.GetString($decodedBytes))
            if (-not [string]::IsNullOrWhiteSpace($content)) {
                $chunks.Add($content)
            }
        }

        $joined = (($chunks | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join "`n").Trim()
        $error = $null
        if ([string]::IsNullOrWhiteSpace($joined)) {
            $error = "No extractable text found. This may be an image-only PDF or use unsupported encoding."
        }

        return [pscustomobject]@{
            Content  = $joined
            Strategy = "Basic PDF text extraction"
            Error    = $error
        }
    }
    catch {
        return [pscustomobject]@{
            Content  = ""
            Strategy = "Basic PDF text extraction"
            Error    = $_.Exception.Message
        }
    }
}

function Get-DocumentContent {
    param(
        [System.IO.FileInfo]$File,
        [int64]$MaxPlainTextBytes
    )

    switch -Regex ($File.Extension) {
        '^\.(txt|log|csv|tsv|md|rtf|json|xml|html|htm|yaml|yml|ini|config|ps1|psm1|psd1|sql|cmd|bat)$' {
            return Read-PlainTextFile -Path $File.FullName -MaxBytes $MaxPlainTextBytes
        }
        '^\.(docx|docm)$' {
            return Read-WordOpenXml -Path $File.FullName
        }
        '^\.(xlsx|xlsm)$' {
            return Read-ExcelOpenXml -Path $File.FullName
        }
        '^\.(pptx|pptm)$' {
            return Read-PowerPointOpenXml -Path $File.FullName
        }
        '^\.pdf$' {
            return Read-PdfText -Path $File.FullName
        }
        default {
            return [pscustomobject]@{
                Content  = ""
                Strategy = "Unsupported"
                Error    = "Extension is not supported for content extraction."
            }
        }
    }
}

function New-IndexedItem {
    param(
        [System.IO.FileInfo]$File,
        [string]$RootPath,
        [int64]$MaxPlainTextBytes
    )

    $extracted = Get-DocumentContent -File $File -MaxPlainTextBytes $MaxPlainTextBytes

    [pscustomobject]@{
        Name              = $File.Name
        FullPath          = $File.FullName
        RelativePath      = Get-RelativePathSafe -RootPath $RootPath -FullPath $File.FullName
        Directory         = $File.DirectoryName
        Extension         = $File.Extension
        Length            = [int64]$File.Length
        LastWriteUtc      = $File.LastWriteTimeUtc.ToString("o")
        Content           = $extracted.Content
        ExtractionMethod  = $extracted.Strategy
        ExtractionError   = $extracted.Error
        ContentLength     = if ([string]::IsNullOrWhiteSpace($extracted.Content)) { 0 } else { $extracted.Content.Length }
    }
}

function Sync-ItDocumentIndex {
    param(
        [string]$SourcePath,
        [string]$OutputPath,
        [int]$Parallelism,
        [int64]$MaxPlainTextBytes,
        [switch]$ForceRebuild
    )

    if (-not (Test-SourceAccessible -Path $SourcePath)) {
        throw "Source path not found or not accessible: $SourcePath"
    }

    Ensure-ParentDirectory -Path $OutputPath

    $existingMap = if ($ForceRebuild) { @{} } else { Import-ExistingIndexItems -OutputPath $OutputPath }
    $allSupportedFiles = @(
        Get-ChildItem -LiteralPath $SourcePath -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { Test-SupportedDocument -File $_ }
    )

    $unchangedItems = New-Object System.Collections.Generic.List[object]
    $workItems = New-Object System.Collections.Generic.List[System.IO.FileInfo]

    foreach ($file in $allSupportedFiles) {
        $existing = $existingMap[$file.FullName]
        $currentLastWriteUtc = $file.LastWriteTimeUtc.ToString("o")

        if ($existing -and $existing.Length -eq [int64]$file.Length -and $existing.LastWriteUtc -eq $currentLastWriteUtc) {
            $unchangedItems.Add($existing)
        }
        else {
            $workItems.Add($file)
        }
    }

    Write-Host ("Supported files found: {0}" -f $allSupportedFiles.Count)
    Write-Host ("Reused from existing index: {0}" -f $unchangedItems.Count)
    Write-Host ("New or changed files: {0}" -f $workItems.Count)

    $indexedWorkItems = @()
    if ($workItems.Count -gt 0) {
        foreach ($workItem in $workItems) {
            $indexedWorkItems += New-IndexedItem -File $workItem -RootPath $SourcePath -MaxPlainTextBytes $MaxPlainTextBytes
        }
    }

    $combinedItems = @($unchangedItems.ToArray()) + @($indexedWorkItems)
    $sortedItems = @($combinedItems | Sort-Object FullPath)
    $errorCount = @($sortedItems | Where-Object { -not [string]::IsNullOrWhiteSpace($_.ExtractionError) }).Count

    $payload = [pscustomobject]@{
        GeneratedUtc = (Get-Date).ToUniversalTime().ToString("o")
        SourcePath   = $SourcePath
        FileCount    = $sortedItems.Count
        ErrorCount   = $errorCount
        Items        = $sortedItems
    }

    $payload | Export-Clixml -LiteralPath $OutputPath
    return $payload
}

function Load-ItDocumentIndex {
    param(
        [string]$SourcePath,
        [string]$OutputPath,
        [int]$Parallelism,
        [int64]$MaxPlainTextBytes,
        [switch]$ForceRebuild
    )

    try {
        return Sync-ItDocumentIndex -SourcePath $SourcePath -OutputPath $OutputPath -Parallelism $Parallelism -MaxPlainTextBytes $MaxPlainTextBytes -ForceRebuild:$ForceRebuild
    }
    catch {
        if (-not $ForceRebuild -and (Test-Path -LiteralPath $OutputPath)) {
            Write-Warning ("Could not refresh the source path right now. Using the last local index instead. Reason: {0}" -f $_.Exception.Message)
            return Import-Clixml -LiteralPath $OutputPath
        }

        throw
    }
}

function Search-ItDocumentIndex {
    param(
        [pscustomobject]$Index,
        [string]$Query,
        [string]$Mode,
        [int]$Limit
    )

    $query = if ($null -eq $Query) { "" } else { $Query.Trim() }
    $mode = if ([string]::IsNullOrWhiteSpace($Mode)) { "all" } else { $Mode.ToLowerInvariant() }

    $results = foreach ($item in $Index.Items) {
        $nameHit = $false
        $pathHit = $false
        $contentHit = $false
        $errorHit = $false

        if ([string]::IsNullOrEmpty($query)) {
            $nameHit = $true
        }
        else {
            if ($mode -in @("all", "name")) {
                $nameHit = $item.Name.IndexOf($query, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            }
            if ($mode -in @("all", "path")) {
                $pathHit = $item.RelativePath.IndexOf($query, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            }
            if ($mode -in @("all", "content") -and -not [string]::IsNullOrWhiteSpace($item.Content)) {
                $contentHit = $item.Content.IndexOf($query, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            }
            if ($mode -eq "errors" -and -not [string]::IsNullOrWhiteSpace($item.ExtractionError)) {
                $errorHit = $true
            }
        }

        if ($nameHit -or $pathHit -or $contentHit -or $errorHit) {
            $matchAreas = New-Object System.Collections.Generic.List[string]
            if ($nameHit) { $matchAreas.Add("name") }
            if ($pathHit) { $matchAreas.Add("path") }
            if ($contentHit) { $matchAreas.Add("content") }
            if ($errorHit) { $matchAreas.Add("errors") }

            [pscustomobject]@{
                Name             = $item.Name
                RelativePath     = $item.RelativePath
                FullPath         = $item.FullPath
                Extension        = $item.Extension
                Length           = $item.Length
                LastWriteUtc     = $item.LastWriteUtc
                MatchAreas       = ($matchAreas -join ", ")
                Snippet          = Get-Snippet -Content $item.Content -Query $query
                ExtractionMethod = $item.ExtractionMethod
                ExtractionError  = $item.ExtractionError
                ContentLength    = $item.ContentLength
            }
        }
    }

    return $results |
        Sort-Object @{ Expression = "LastWriteUtc"; Descending = $true }, @{ Expression = "Name"; Descending = $false } |
        Select-Object -First $Limit
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

    [pscustomobject]@{
        sourcePath   = $Index.SourcePath
        generatedUtc = $Index.GeneratedUtc
        fileCount    = $Index.FileCount
        errorCount   = $Index.ErrorCount
    }
}

function Get-PageHtml {
    param(
        [pscustomobject]$Index,
        [int]$MaxPlainTextFileSizeMB
    )

    $template = @'
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>IT Document Search</title>
  <style>
    :root {
      --bg: #f4f6f8;
      --panel: #ffffff;
      --ink: #13212b;
      --muted: #5b6b77;
      --line: #d9e0e6;
      --accent: #0e6ba8;
      --accent-soft: #d7ebf8;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Segoe UI", Tahoma, sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at top right, rgba(14,107,168,.10) 0, transparent 22rem),
        linear-gradient(180deg, #f9fbfc 0%, var(--bg) 100%);
    }
    .wrap {
      max-width: 1220px;
      margin: 0 auto;
      padding: 28px 18px 48px;
    }
    .hero, .search-panel, .results-panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 18px;
      box-shadow: 0 12px 32px rgba(24, 39, 56, .06);
    }
    .hero { padding: 22px; }
    h1 {
      margin: 0 0 8px;
      font-size: clamp(2rem, 5vw, 3.2rem);
      line-height: 1;
      letter-spacing: -.03em;
    }
    p {
      margin: 0;
      color: var(--muted);
    }
    .meta {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 12px;
      margin-top: 18px;
    }
    .meta-card {
      background: #f9fbfd;
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 14px 16px;
    }
    .meta-label {
      display: block;
      color: var(--muted);
      font-size: .8rem;
      text-transform: uppercase;
      letter-spacing: .08em;
    }
    .meta-value {
      display: block;
      margin-top: 6px;
      word-break: break-word;
    }
    .search-panel {
      margin-top: 20px;
      padding: 18px;
    }
    form {
      display: grid;
      grid-template-columns: minmax(0, 1.8fr) 190px 120px 140px;
      gap: 12px;
    }
    input, select, button {
      width: 100%;
      border-radius: 12px;
      border: 1px solid var(--line);
      padding: 12px 14px;
      font: inherit;
      background: #fff;
    }
    input:focus, select:focus {
      outline: 2px solid rgba(14,107,168,.18);
      border-color: var(--accent);
    }
    button {
      cursor: pointer;
      background: var(--accent);
      color: white;
      border: none;
    }
    button.secondary { background: #355c74; }
    .toolbar {
      display: flex;
      gap: 12px;
      align-items: center;
      margin-top: 14px;
      color: var(--muted);
      flex-wrap: wrap;
    }
    .results-panel {
      margin-top: 20px;
      padding: 8px;
    }
    .results-head {
      display: flex;
      justify-content: space-between;
      gap: 16px;
      align-items: center;
      padding: 12px 14px;
      color: var(--muted);
      flex-wrap: wrap;
    }
    .result {
      border-top: 1px solid #e8edf1;
      padding: 16px 14px;
    }
    .result:first-of-type { border-top: 0; }
    .result h3 {
      margin: 0;
      font-size: 1.06rem;
    }
    .path {
      margin-top: 4px;
      color: var(--muted);
      word-break: break-word;
      font-size: .94rem;
    }
    .snippet, .warning {
      margin-top: 10px;
      border-radius: 12px;
      padding: 10px 12px;
      white-space: pre-wrap;
    }
    .snippet {
      font-family: Consolas, "Courier New", monospace;
      background: #fbfcfd;
      border: 1px solid #e1e8ee;
    }
    .warning {
      background: #fff5e7;
      border: 1px solid #f2d2a6;
      color: #73440c;
    }
    .badges {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-top: 10px;
    }
    .badge {
      border-radius: 999px;
      padding: 5px 10px;
      background: var(--accent-soft);
      color: #0a4f7d;
      font-size: .82rem;
    }
    .empty {
      padding: 26px 16px 34px;
      text-align: center;
      color: var(--muted);
    }
    .loading { opacity: .72; }
    @media (max-width: 900px) {
      form { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <section class="hero">
      <h1>IT Document Search</h1>
      <p>Indexes standard document formats from a local or mapped folder into a local cache for faster filename, path, and content search. Image-only PDFs are not OCR'd.</p>
      <div class="meta">
        <div class="meta-card">
          <span class="meta-label">Source</span>
          <span class="meta-value">__SOURCE_PATH__</span>
        </div>
        <div class="meta-card">
          <span class="meta-label">Indexed Files</span>
          <span class="meta-value" id="fileCount">__FILE_COUNT__</span>
        </div>
        <div class="meta-card">
          <span class="meta-label">Extraction Warnings</span>
          <span class="meta-value" id="errorCount">__ERROR_COUNT__</span>
        </div>
        <div class="meta-card">
          <span class="meta-label">Index Built</span>
          <span class="meta-value" id="generatedUtc">__GENERATED_UTC__</span>
        </div>
      </div>
    </section>

    <section class="search-panel">
      <form id="searchForm">
        <input id="query" name="query" type="text" placeholder="Search filename, path, document text..." autocomplete="off">
        <select id="mode" name="mode">
          <option value="all">Name + Path + Contents</option>
          <option value="name">Filename Only</option>
          <option value="path">Path Only</option>
          <option value="content">Contents Only</option>
          <option value="errors">Extraction Warnings</option>
        </select>
        <input id="limit" name="limit" type="number" min="1" max="500" value="100">
        <button type="submit">Search</button>
      </form>
      <div class="toolbar">
        <button id="reindexBtn" class="secondary" type="button">Refresh Index</button>
        <span id="status">Ready. Plain-text files over __MAX_TEXT_MB__ MB are skipped for content extraction.</span>
      </div>
    </section>

    <section class="results-panel">
      <div class="results-head">
        <strong id="resultCount">No search yet</strong>
        <span>Tip: use the warnings mode to find files that could not be fully read.</span>
      </div>
      <div id="results" class="empty">Run a search to see matching documents.</div>
    </section>
  </div>

  <script>
    const form = document.getElementById('searchForm');
    const resultsEl = document.getElementById('results');
    const resultCountEl = document.getElementById('resultCount');
    const statusEl = document.getElementById('status');
    const generatedUtcEl = document.getElementById('generatedUtc');
    const fileCountEl = document.getElementById('fileCount');
    const errorCountEl = document.getElementById('errorCount');
    const reindexBtn = document.getElementById('reindexBtn');

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
          item.matchAreas ? `<span class="badge">${escapeHtml(item.matchAreas)}</span>` : '',
          item.extension ? `<span class="badge">${escapeHtml(item.extension)}</span>` : '<span class="badge">[no extension]</span>',
          `<span class="badge">${Number(item.length || 0).toLocaleString()} bytes</span>`,
          item.extractionMethod ? `<span class="badge">${escapeHtml(item.extractionMethod)}</span>` : '',
          `<span class="badge">${Number(item.contentLength || 0).toLocaleString()} chars</span>`
        ].join('');

        const warning = item.extractionError
          ? `<div class="warning">${escapeHtml(item.extractionError)}</div>`
          : '';

        return `
          <article class="result">
            <h3>${escapeHtml(item.name)}</h3>
            <div class="path">${escapeHtml(item.relativePath)}</div>
            <div class="path">${escapeHtml(item.fullPath)}</div>
            <div class="badges">${badges}</div>
            ${warning}
            <div class="snippet">${escapeHtml(item.snippet || '')}</div>
          </article>
        `;
      }).join('');
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
        renderResults(payload);
        statusEl.textContent = `Search finished in ${payload.elapsedMs} ms.`;
      } catch (error) {
        resultsEl.className = 'empty';
        resultsEl.textContent = error.message;
        statusEl.textContent = 'Search failed.';
      } finally {
        resultsEl.classList.remove('loading');
      }
    }

    async function rebuildIndex() {
      statusEl.textContent = 'Refreshing index...';
      reindexBtn.disabled = true;

      try {
        const response = await fetch('/api/reindex', { method: 'POST' });
        const payload = await response.json();
        generatedUtcEl.textContent = payload.summary.generatedUtc;
        fileCountEl.textContent = payload.summary.fileCount;
        errorCountEl.textContent = payload.summary.errorCount;
        statusEl.textContent = 'Index refreshed.';
      } catch (error) {
        statusEl.textContent = 'Refresh failed.';
      } finally {
        reindexBtn.disabled = false;
      }
    }

    form.addEventListener('submit', event => {
      event.preventDefault();
      runSearch();
    });

    reindexBtn.addEventListener('click', rebuildIndex);
  </script>
</body>
</html>
'@

    return $template.
        Replace("__SOURCE_PATH__", (Escape-Html $Index.SourcePath)).
        Replace("__FILE_COUNT__", [string]$Index.FileCount).
        Replace("__ERROR_COUNT__", [string]$Index.ErrorCount).
        Replace("__GENERATED_UTC__", (Escape-Html $Index.GeneratedUtc)).
        Replace("__MAX_TEXT_MB__", [string]$MaxPlainTextFileSizeMB)
}

$script:Index = $null
$listener = $null
$errorLogPath = Join-Path $PSScriptRoot "data\it-document-search-error.log"
$maxPlainTextBytes = [int64]$MaxPlainTextFileSizeMB * 1MB

try {
    Assert-PathOutsideSource -Path $IndexPath -SourcePath $SourcePath -Label "IndexPath"
    Assert-PathOutsideSource -Path $errorLogPath -SourcePath $SourcePath -Label "Error log path"

    Write-Host "Loading IT document index and refreshing changed files..."
    $script:Index = Load-ItDocumentIndex -SourcePath $SourcePath -OutputPath $IndexPath -Parallelism $ThrottleLimit -MaxPlainTextBytes $maxPlainTextBytes -ForceRebuild:$Rebuild
    Write-Host ("Index ready. Files: {0}. Generated: {1}" -f $script:Index.FileCount, $script:Index.GeneratedUtc)

    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add($Prefix)
    $listener.Start()

    Write-Host "IT Document Search running at $Prefix"
    Write-Host "Press Ctrl+C to stop."
    Start-Process -FilePath $Prefix | Out-Null

    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        try {
            switch ($request.Url.AbsolutePath) {
                "/" {
                    Write-TextResponse -Response $response -Body (Get-PageHtml -Index $script:Index -MaxPlainTextFileSizeMB $MaxPlainTextFileSizeMB)
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

                    $sw = [System.Diagnostics.Stopwatch]::StartNew()
                    $results = Search-ItDocumentIndex -Index $script:Index -Query $query -Mode $mode -Limit $limit
                    $sw.Stop()

                    Write-JsonResponse -Response $response -Payload @{
                        query = $query
                        mode = $mode
                        limit = $limit
                        elapsedMs = $sw.ElapsedMilliseconds
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
                                extractionMethod = $_.ExtractionMethod
                                extractionError = $_.ExtractionError
                                contentLength = $_.ContentLength
                            }
                        })
                    }
                }
                "/api/reindex" {
                    if ($request.HttpMethod -ne "POST") {
                        Write-JsonResponse -Response $response -Payload @{ error = "Method not allowed." } -StatusCode 405
                        break
                    }

                    $script:Index = Sync-ItDocumentIndex -SourcePath $SourcePath -OutputPath $IndexPath -Parallelism $ThrottleLimit -MaxPlainTextBytes $maxPlainTextBytes
                    Write-JsonResponse -Response $response -Payload @{
                        ok = $true
                        summary = Get-IndexSummary -Index $script:Index
                    }
                }
                default {
                    Write-TextResponse -Response $response -Body "Not found." -ContentType "text/plain; charset=utf-8" -StatusCode 404
                }
            }
        }
        catch {
            Write-JsonResponse -Response $response -Payload @{
                error = $_.Exception.Message
            } -StatusCode 500
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
    Write-Host "IT Document Search crashed during startup." -ForegroundColor Red
    Write-Host ("Error: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-Host ("Details were written to: {0}" -f $errorLogPath) -ForegroundColor Yellow
    Read-Host "Press Enter to close"
}
finally {
    if ($null -ne $listener -and $listener.IsListening) {
        $listener.Stop()
    }
    if ($null -ne $listener) {
        $listener.Close()
    }
}
