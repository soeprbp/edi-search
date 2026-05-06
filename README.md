# Welch Support Tools

This workspace now contains:

- a local launcher menu for Welch support tools
- a local web search tool for EDI files
- a local document search tool for `I:\IT`

## Start The Support Menu

From this folder:

```powershell
powershell -ExecutionPolicy Bypass -File .\Start-SupportMenu.ps1
```

This opens a local Windows launcher window directly, so there is no browser address to open.
By default, the backing PowerShell console is hidden while the launcher is open.
For troubleshooting, you can keep it visible with:

```powershell
powershell -ExecutionPolicy Bypass -File .\Start-SupportMenu.ps1 -ShowConsole
```

The menu reads its tool list from `support-tools.json`.

Add more tools by appending entries to the `tools` array. The first starter entry launches `Start-EdiSearch.ps1`.
Relative paths can point to sibling or parent script folders too, such as `..\\SomeOtherTool.ps1`.

Supported launch types right now:

- `powershell-script`
- `program`
- `url`

Example config entry:

```json
{
  "id": "my-tool",
  "name": "My Tool",
  "category": "Support",
  "description": "Short description of what it does.",
  "launchType": "powershell-script",
  "target": "MyTool.ps1",
  "arguments": [],
  "workingDirectory": ".",
  "openInBrowser": false
}
```

After changing `support-tools.json`, use the menu's **Reload Config** button in the launcher window to pick up the new entries.

## EDI Search

EDI Search provides two integrated tools in one web interface:

### File Search
- filename search
- path search
- full-content search
- copy eligible `.bak` files from `Processed` to the parent folder as `.cov`

### Log Search
- searches the EDI application log file (`\\svwpefs\WelchEncoreShare\WELCHPKG\EDI\Logs\EDI_log.txt`)
- pre-indexed at startup for fast searches (last 30 days only)
- date range filtering with start/end date pickers
- parsed log entries showing: timestamp, customer, ship-to, transaction type, PID, and message
- context lines before/after each match
- newest-first result ordering
- in-memory caching for repeated searches (exact-query cache + prefix reuse)

It does **not** change the folder structure in `\\svwpefs\WelchEncoreShare\WELCHPKG\EDI`.

It only:

1. reads files from that share
2. reads the log file for the last 30 days
3. writes its own local index to `data\edi-index.clixml`
4. serves a browser UI on `http://localhost:8787/`
5. optionally copies a selected `.bak` file to its parent folder as a `.cov` file when you confirm it in the UI

## Run EDI Search Directly

From this folder:

```powershell
powershell -ExecutionPolicy Bypass -File .\Start-EdiSearch.ps1
```

Then open:

```text
http://localhost:8787/
```

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-SharePath` | `\\svwpefs\WelchEncoreShare\WELCHPKG\EDI` | EDI share to index |
| `-LogPath` | `\\...\EDI\Logs\EDI_log.txt` | Log file to index |
| `-MaxResults` | `200` | Default max file-search results |
| `-ThrottleLimit` | CPU-based | Parallelism for indexing operations |
| `-MaxContentSize` | `50000` | Max bytes of file content to index per file |
| `-Rebuild` | `$false` | Force full rebuild of file index |

## Log Search Performance

- Log indexing reads and keeps only the most recent 30 days to reduce startup and query time.
- Date filters use index-aware range lookups (binary search) before text matching.
- Query matching uses fast loop-based search with early exit once result limit is reached.
- When additional matches exist beyond the current limit, the UI shows "Showing first X matches (more available)".
- Cache behavior:
  - exact cache hit: same query + date range + limit + current log index version
  - prefix reuse hit: narrower query reuses broader cached result set when safe
  - cache TTL: short-lived in-memory cache for active session speedups

## Refresh Behavior

- On every normal startup, the script scans the share and refreshes only files that are new or changed.
- Unchanged files are reused from the saved local index.
- Deleted files drop out automatically because the index is rebuilt from the current file list.
- Refreshing changed files is multi-threaded with PowerShell parallel workers.

## Full Rebuild

Force a full rebuild from scratch on startup:

```powershell
powershell -ExecutionPolicy Bypass -File .\Start-EdiSearch.ps1 -Rebuild
```

Or use the **Refresh Index** button in the page for a normal incremental refresh.

## Notes

- First run will take a while because it has to read and index the share.
- After that, startup should be much faster because only new or changed files are reread.
- Searches happen against the local index, so content searching is much faster.
- The index is separate from the share and can be deleted any time without affecting the EDI folders.
- The copy action is intentionally limited to `.bak` files under a `Processed` folder and will drop the copied file in the parent folder as `.cov`.
- It will not create missing folders or overwrite an existing target file.
- Log file indexing only processes the last 30 days by default for faster startup.
- Use the date range filters in the Log Search section to narrow down results.
- `Ctrl+C` now shuts down the local web server cleanly.

## Troubleshooting (EDI Search)

- **Ctrl+C does not stop the script**
  - Start the script in a normal PowerShell host (not a detached/background host).
  - If needed, close the window as a fallback; the listener is closed in `finally`.

- **Log search returns HTTP 500**
  - Check `data\edi-search-error.log` for the latest exception details.
  - Confirm `-LogPath` exists and is readable from your session.
  - Try narrowing the date range and rerun.

- **Log search still feels slow**
  - Use a date range first, then add text query.
  - Repeat similar searches to benefit from in-memory cache/prefix reuse.

## IT Document Search

- indexes supported files under `I:\IT` into `data\it-document-index.clixml`
- searches filename, path, and extracted document text in a local browser UI
- supports plain-text files plus `.docx`, `.docm`, `.xlsx`, `.xlsm`, `.pptx`, `.pptm`, and text-based `.pdf`
- tracks extraction warnings so unreadable or image-only PDFs are easy to spot

Run it directly:

```powershell
powershell -ExecutionPolicy Bypass -File .\Start-ItDocumentSearch.ps1
```

Then open:

```text
http://localhost:8788/
```

Notes:

- This tool does not depend on Copilot to read documents.
- It is read-only against the source folder and writes only to local files under this workspace, never back into `I:\IT`.
- PDFs are handled with a lightweight built-in text extractor, so scanned/image-only PDFs will usually index by name and path but may have no searchable content.
- Large plain-text files over the configured size limit are skipped for content extraction and flagged in the warnings view.
- The source path is configurable with `-SourcePath` if the drive mapping changes.

## Data Folder

The `data\` folder contains local index files and logs. These can be safely deleted to force a rebuild.

| File | Description |
|------|-------------|
| `edi-index.clixml` | Cached index of EDI files |
| `edi-search-error.log` | Error log from EDI Search web server |
| `it-document-index.clixml` | Cached index of IT documents |
| `it-document-search-error.log` | Error log from IT Document Search |
| `*.evtx` | Windows Event Log exports (Application, Security, System) |
