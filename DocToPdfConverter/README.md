# DOC to PDF Converter

A console application that converts Microsoft Word (.doc) files to PDF format using NetOfficeFw.Word library.

## Features

- ? Batch conversion of all .doc files in a directory
- ? Configurable input and output directories via `appsettings.json`
- ? Comprehensive logging with Serilog (console + file)
- ? Robust exception handling
- ? Detailed conversion statistics

## Prerequisites

- .NET 8.0 Runtime
- Microsoft Word installed on the machine (required for NetOfficeFw.Word)

## Configuration

Edit `appsettings.json` to configure the input and output directories:

```json
{
  "Converter": {
    "InputDirectory": "C:\\Input",
    "OutputDirectory": "C:\\Output"
  }
}
```

### Configuration Options

- **InputDirectory**: Directory containing .doc files to convert
- **OutputDirectory**: Directory where PDF files will be saved (created automatically if doesn't exist)

## Usage

1. Configure input/output directories in `appsettings.json`
2. Place your .doc files in the input directory
3. Run the application:
   ```
   dotnet run
   ```
   or execute the compiled exe:
   ```
   DocToPdfConverter.exe
   ```

## Logging

Logs are written to:
- **Console**: Real-time progress and status
- **File**: `logs/converter-{date}.log` (rolls daily)

Log levels:
- `Information`: Normal operations
- `Warning`: No files found, etc.
- `Error`: Conversion failures
- `Fatal`: Application-level failures

## Output

The application will:
1. Scan the input directory for all .doc files
2. Convert each file to PDF format
3. Save PDFs in the output directory with the same filename
4. Log success/failure for each conversion
5. Display summary statistics at the end

## Error Handling

- Invalid/missing directories are validated before processing
- Each file conversion is wrapped in error handling
- Failed conversions are logged but don't stop the batch process
- COM objects are properly disposed to prevent memory leaks

## Example Output

```
[2025-11-06 10:15:30.123] [INF] DOC to PDF Converter started
[2025-11-06 10:15:30.234] [INF] Input Directory: C:\Input
[2025-11-06 10:15:30.235] [INF] Output Directory: C:\Output
[2025-11-06 10:15:30.450] [INF] Found 3 .doc file(s) to convert
[2025-11-06 10:15:30.451] [INF] Converting: document1.doc
[2025-11-06 10:15:32.890] [INF] Successfully converted: document1.doc -> document1.pdf
[2025-11-06 10:15:32.891] [INF] Converting: document2.doc
[2025-11-06 10:15:35.234] [INF] Successfully converted: document2.doc -> document2.pdf
[2025-11-06 10:15:35.235] [INF] Converting: document3.doc
[2025-11-06 10:15:37.567] [INF] Successfully converted: document3.doc -> document3.pdf
[2025-11-06 10:15:37.568] [INF] Conversion completed. Success: 3, Failed: 0
```

## Troubleshooting

**Issue**: Application fails to start
- Verify Microsoft Word is installed
- Ensure .NET 8.0 runtime is installed

**Issue**: "Input directory does not exist"
- Check the path in `appsettings.json` is correct and accessible
- Ensure proper escaping of backslashes in JSON (`\\` not `\`)

**Issue**: Conversion fails for specific files
- Check if the .doc file is corrupted
- Ensure Word can open the file manually
- Review the log file for detailed error messages

## License

This project uses NetOfficeFw.Word which is subject to its own license terms.
