# DOC to PDF Converter (LibreOffice)

A .NET 8 console application that converts Microsoft Word documents (.doc and .docx) to PDF format using LibreOffice in headless mode. Optimized for batch processing with parallel execution, file checking, and pattern-based filtering.

## Key Features

- **Parallel Processing** - Convert multiple files simultaneously (4-8 threads recommended)
- **Smart Timestamp Checking** - Only converts files when source is newer than PDF
- **Pattern Filtering** - Selectively process files based on name patterns
- **Resume Capability** - Automatically skips already-converted files
- **Progress Tracking** - Real-time progress reporting and statistics
- **Comprehensive Logging** - Detailed logs with Serilog
- **Single-File Deployment** - Self-contained executable with embedded configuration
- **Cross-Platform** - Works on Windows, Linux, and macOS
- **Production Ready** - Stable parallel processing with excellent error isolation

## Requirements

### Runtime Requirements
- **LibreOffice** (free, open-source)
  - Windows: [Download LibreOffice](https://www.libreoffice.org/download/download/)
  - Linux: `sudo apt-get install libreoffice`
  - macOS: `brew install --cask libreoffice`

### For Development
- .NET 8 SDK
- Visual Studio 2022 or VS Code

## ?? Quick Start

### 1. Install LibreOffice
Download and install from [libreoffice.org](https://www.libreoffice.org/download/download/)

### 2. Configure Application
Edit `appsettings.json`:

```json
{
  "Converter": {
    "InputDirectory": "C:\\Input",
    "OutputDirectory": "C:\\Output",
  "LibreOfficePath": "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    "MaxDegreeOfParallelism": 8,
    "ChunkSize": 100,
    "TimeoutSeconds": 60,
    "FileNamePatterns": []
  }
}
```

### 3. Run Application

**Standard Run:**
```bash
dotnet run
```

**Or build and run executable:**
```bash
dotnet build -c Release
cd bin\Release\net8.0
.\DocToPdfConverterLibreOffice.exe
```

## Configuration Options

### Converter Settings

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `InputDirectory` | string | Required | Directory containing .doc/.docx files |
| `OutputDirectory` | string | Required | Directory for generated PDFs |
| `LibreOfficePath` | string | Auto-detect | Full path to soffice.exe (auto-detected if not specified) |
| `MaxDegreeOfParallelism` | int | 4 | Number of concurrent conversion processes (1-16) |
| `ChunkSize` | int | 100 | Files per processing chunk |
| `TimeoutSeconds` | int | 60 | Timeout per file conversion |
| `FileNamePatterns` | string[] | `[]` | File name patterns to filter (empty = all files) |

### Serilog Settings

| Setting | Default | Description |
|---------|---------|-------------|
| `MinimumLevel` | Information | Log level (Debug, Information, Warning, Error) |
| `WriteTo` | Console, File | Log destinations |
| `path` | logs/converter-.log | Log file path (daily rolling) |

## Configuration Examples

### Basic Configuration
```json
{
  "Converter": {
    "InputDirectory": "D:\\Documents",
    "OutputDirectory": "D:\\PDFs"
  }
}
```
LibreOffice path will be auto-detected.

### High-Performance Configuration
```json
{
  "Converter": {
  "InputDirectory": "D:\\Documents",
    "OutputDirectory": "D:\\PDFs",
    "MaxDegreeOfParallelism": 12,
    "ChunkSize": 200,
    "TimeoutSeconds": 90
  }
}
```
Optimized for 12+ core servers with large batches.

### Pattern Filtering Configuration
```json
{
  "Converter": {
    "InputDirectory": "D:\\Documents",
    "OutputDirectory": "D:\\PDFs",
    "FileNamePatterns": [
   "Invoice_2024_*.doc",
      "Report_Q1_*.doc",
      "*_110.doc"
    ]
  }
}
```
Only processes files matching specified patterns.

### Debug Configuration
```json
{
  "Converter": {
    "InputDirectory": "D:\\Test",
    "OutputDirectory": "D:\\TestOutput",
    "MaxDegreeOfParallelism": 1,
    "FileNamePatterns": ["Test_*"]
  },
  "Serilog": {
    "MinimumLevel": "Debug"
  }
}
```
Single-threaded with debug logging for troubleshooting.

## File Name Patterns

### Pattern Syntax

| Wildcard | Meaning | Example | Matches |
|----------|---------|---------|---------|
| `*` | Any characters | `Report_*` | Report_2024.doc, Report_Jan.doc |
| `?` | Single character | `Doc_????.doc` | Doc_2024.doc, Doc_Test.doc |

### Pattern Examples

```json
{
  "FileNamePatterns": [
    "*_110",           // Files ending with _110
    "110_*",           // Files starting with 110_
    "*_2024_*",        // Files containing _2024_
    "Invoice_*.doc",   // Invoice .doc files only
    "Doc_????.doc",    // Doc with exactly 4 characters
    "HR_*",            // HR department files
    "*.docx"           // All .docx files only
  ]
}
```

### Pattern Behavior

- **Empty array `[]`**: Process ALL files (default)
- **Multiple patterns**: Process files matching ANY pattern (OR logic)
- **Case insensitive**: `REPORT_*` matches `report_110.doc`

## Performance Tuning

### Recommended Settings by Scenario

#### Development/Testing
```json
{
  "MaxDegreeOfParallelism": 2,
  "ChunkSize": 50,
  "TimeoutSeconds": 60
}
```

#### Production - Standard Server
```json
{
  "MaxDegreeOfParallelism": 8,
  "ChunkSize": 100,
  "TimeoutSeconds": 60
}
```

#### Production - High-Performance Server
```json
{
  "MaxDegreeOfParallelism": 16,
  "ChunkSize": 200,
  "TimeoutSeconds": 90
}
```

#### Network Drive / Slow Storage
```json
{
  "MaxDegreeOfParallelism": 4,
  "ChunkSize": 50,
  "TimeoutSeconds": 120
}
```

### Performance Guidelines

| Hardware | MaxDegreeOfParallelism | Expected Throughput |
|----------|----------------------|---------------------|
| 4-core CPU, 8GB RAM | 4 | ~1 file/second |
| 8-core CPU, 16GB RAM | 8 | ~2 files/second |
| 16-core CPU, 32GB RAM | 16 | ~4 files/second |

**Memory Usage:** ~100-150 MB per concurrent process

## Smart Features

### 1. Smart Timestamp Checking

Automatically detects when source files are modified:

```
If PDF exists:
  ?? PDF newer than source? ? Skip ?
  ?? Source newer than PDF? ? Regenerate ??
```

**Example:**
```
Document1.doc (Modified: 2025-01-01 10:00)
Document1.pdf (Created: 2025-01-01 10:15)
Skipped (PDF is up-to-date)

Document2.doc (Modified: 2025-01-02 11:30)
Document2.pdf (Created: 2025-01-01 10:16)
Regenerated (Source modified after PDF)
```

### 2. Resume Capability

If conversion is interrupted, simply restart - already converted files are automatically skipped.

### 3. Automatic LibreOffice Detection

If `LibreOfficePath` is not specified, the application auto-detects LibreOffice in common locations:
- `C:\Program Files\LibreOffice\program\soffice.exe`
- `C:\Program Files (x86)\LibreOffice\program\soffice.exe`
- System PATH

## Example Output

```
[10:00:00 INF] DOC to PDF Converter (LibreOffice) started
[10:00:00 INF] LibreOffice Path: C:\Program Files\LibreOffice\program\soffice.exe
[10:00:00 INF] Input Directory: D:\Documents
[10:00:00 INF] Output Directory: D:\PDFs
[10:00:00 INF] Max Degree of Parallelism: 8
[10:00:00 INF] Chunk Size: 100
[10:00:00 INF] Timeout: 60 seconds
[10:00:00 INF] File Name Patterns: None (processing all files)
[10:00:00 INF] Found 24000 document file(s) to convert
[10:00:00 INF] Starting parallel conversion with 8 concurrent processes
[10:00:00 INF] Processing 24000 files in 240 chunks of 100
[10:00:00 INF] Processing chunk 1/240 (100 files)
[10:05:00 INF] Progress: 5% (1200/24000 files)
[10:10:00 INF] Progress: 10% (2400/24000 files)
...
[14:00:00 INF] Progress: 100% (24000/24000 files)
[14:00:00 INF] =====================================
[14:00:00 INF] Conversion completed in 14400.00 seconds
[14:00:00 INF] Success: 23950, Skipped: 0, Failed: 50, Total: 24000
[14:00:00 INF] Average time per file: 0.60 seconds
[14:00:00 INF] Throughput: 1.67 files/second
```

## Single-File Deployment

### Build Self-Contained Executable

**Windows (x64):**
```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

**Linux (x64):**
```bash
dotnet publish -c Release -r linux-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

**macOS (x64):**
```bash
dotnet publish -c Release -r osx-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

**macOS (ARM64):**
```bash
dotnet publish -c Release -r osx-arm64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

### Output Location
```
bin\Release\net8.0\{runtime}\publish\DocToPdfConverterLibreOffice.exe
```

### Running the Executable

**With Embedded Configuration:**
```bash
.\DocToPdfConverterLibreOffice.exe
```
Uses embedded `appsettings.json` from build time.

**With External Configuration:**
1. Create `appsettings.json` next to the .exe
2. Run the executable
3. External config takes precedence over embedded

### Deployment Steps

1. **Publish the application:**
   ```bash
   dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
   ```

2. **Locate the executable:**
   ```
   bin\Release\net8.0\win-x64\publish\DocToPdfConverterLibreOffice.exe
   ```

3. **Deploy to target system:**
   - Copy the .exe file
   - (Optional) Copy `appsettings.json` for custom configuration
   - Install LibreOffice on target system

4. **Run:**
   ```bash
   DocToPdfConverterLibreOffice.exe
   ```

**No .NET runtime installation required!** The executable is fully self-contained.

## Benefits

### Performance
- **Fast Parallel Processing** - 8x speed improvement with 8 threads
- **Efficient Resource Usage** - Chunked processing prevents memory issues
- **Smart Skipping** - Only converts files that need updating

### Reliability
- **Excellent Stability** - Process-based isolation (not COM)
- **Error Isolation** - One file failure doesn't stop batch
- **Resume Capability** - Safe to restart after interruption
- **Comprehensive Logging** - Detailed logs for troubleshooting

### Flexibility
- **Cross-Platform** - Works on Windows, Linux, macOS
- **Pattern Filtering** - Process specific subsets of files
- **Configurable** - Adjust parallelism, timeouts, and behavior
- **Free** - Uses open-source LibreOffice (no licensing costs)

### Production Ready
- **Unattended Operation** - Stable for long-running batches
- **Detailed Statistics** - Success/failure counts and metrics
- **Single-File Deployment** - Easy distribution
- **Embedded Configuration** - Works out-of-the-box

## Limitations

### LibreOffice Required
- Must have LibreOffice installed on target system
- Conversion quality depends on LibreOffice rendering

### Conversion Fidelity
- Generally excellent for standard documents
- Complex formatting may differ slightly from Word
- Macros are not executed during conversion

### Performance Constraints
- **Disk I/O**: SSD recommended for best performance
- **Network Drives**: Slower; reduce parallelism to 2-4
- **Memory**: ~150 MB per concurrent process

### Platform-Specific
- `where` command (Windows) for auto-detection
- Use `which` on Linux/macOS if modifying code

## Troubleshooting

### LibreOffice Not Found
**Error:** `LibreOffice installation not found`

**Solution:**
1. Install LibreOffice
2. Or specify path manually in `appsettings.json`:
   ```json
   {
     "LibreOfficePath": "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
   }
   ```

### Conversion Failures (Exit Code 1)
**Error:** `LibreOffice exited with code 1`

**Solution:**
- The application automatically creates unique temp profiles per process
- If issues persist, reduce `MaxDegreeOfParallelism` to 4

### Timeouts
**Error:** `LibreOffice conversion timed out`

**Solution:**
- Increase `TimeoutSeconds` in configuration
- Reduce `MaxDegreeOfParallelism`
- Check for corrupted source files

### High Memory Usage
**Solution:**
- Reduce `MaxDegreeOfParallelism`
- Reduce `ChunkSize`
- Ensure sufficient RAM (150 MB × parallelism)

### No Files Processed
**Solution:**
- Check `InputDirectory` path is correct
- Verify `.doc` or `.docx` files exist
- Check `FileNamePatterns` if configured
- Review logs for filtering information

## Use Cases

### Daily Archival
Convert updated documents to PDF daily:
```json
{
  "FileNamePatterns": []
}
```
Smart timestamp checking ensures only modified files are converted.

### Department-Specific Processing
Process only specific department files:
```json
{
  "FileNamePatterns": ["HR_*", "Finance_*"]
}
```

### Batch Processing by Date
Process files from specific time periods:
```json
{
  "FileNamePatterns": ["*_2024_*", "*_Q1_*"]
}
```

### Large-Scale Migration
Convert 100,000+ documents:
```json
{
  "MaxDegreeOfParallelism": 16,
  "ChunkSize": 200
}
```

## Support

### Logs
Check `logs/converter-YYYY-MM-DD.log` for detailed information

### Debug Mode
Enable detailed logging:
```json
{
  "Serilog": {
  "MinimumLevel": "Debug"
  }
}
```

## Version

- **Target Framework:** .NET 8.0
- **Language Version:** C# 12.0
- **LibreOffice:** Compatible with all recent versions

## License

This project uses:
- **LibreOffice** (Mozilla Public License 2.0)
- **Serilog** (Apache License 2.0)

## Quick Tips

1. **Start Small**: Test with 100 files before processing 24,000
2. **Monitor First Hour**: Watch for any issues before leaving unattended
3. **Use SSD**: Dramatically improves performance
4. **Pattern Filtering**: Great for incremental processing
5. **Smart Timestamps**: Run daily - only changed files convert
6. **Check Logs**: Review logs/converter-*.log for issues
7. **Tune Parallelism**: Adjust `MaxDegreeOfParallelism` based on your CPU

## Performance Example

**Scenario: 24,000 files, 4 seconds each**

| Configuration | Time |
|---------------|------|
| Sequential (1 thread) | ~27 hours |
| Parallel (4 threads) | ~7 hours |
| Parallel (8 threads) | ~3.5 hours |
| Parallel (16 threads) | ~2 hours |

**With smart timestamp checking on subsequent runs:**
- Only 100 of 24,000 files changed
- **Time: ~5 seconds!** (99.98% time saved)

---
