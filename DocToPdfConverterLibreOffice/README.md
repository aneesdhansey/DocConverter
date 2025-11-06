# DOC to PDF Converter (LibreOffice CLI)

A console application that converts Microsoft Word (.doc and .docx) files to PDF format using LibreOffice command-line interface.

## Features

- ? Batch conversion of all .doc and .docx files in a directory
- ? Configurable input and output directories via `appsettings.json`
- ? Auto-detection of LibreOffice installation
- ? Comprehensive logging with Serilog (console + file)
- ? Robust exception handling with timeout protection
- ? Cross-platform support (Windows, Linux, macOS)
- ? No Microsoft Office required
- ? Detailed conversion statistics

## Prerequisites

- .NET 8.0 Runtime
- **LibreOffice** installed on the machine
  - Download from: https://www.libreoffice.org/download/

## Installation

### Windows
Download and install LibreOffice from the official website. The application will automatically detect the installation in:
- `C:\Program Files\LibreOffice\program\soffice.exe`
- `C:\Program Files (x86)\LibreOffice\program\soffice.exe`

### Linux
```bash
sudo apt-get install libreoffice
# or
sudo yum install libreoffice
```

### macOS
```bash
brew install --cask libreoffice
```

## Configuration

Edit `appsettings.json` to configure the directories and LibreOffice path:

```json
{
  "Converter": {
    "InputDirectory": "C:\\Input",
    "OutputDirectory": "C:\\Output",
    "LibreOfficePath": "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
  }
}
```

### Configuration Options

- **InputDirectory**: Directory containing .doc/.docx files to convert
- **OutputDirectory**: Directory where PDF files will be saved (created automatically if doesn't exist)
- **LibreOfficePath**: Path to LibreOffice executable (optional - will auto-detect if not specified)

### Platform-Specific Paths

**Windows:**
```json
"LibreOfficePath": "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
```

**Linux:**
```json
"LibreOfficePath": "/usr/bin/soffice"
```

**macOS:**
```json
"LibreOfficePath": "/Applications/LibreOffice.app/Contents/MacOS/soffice"
```

## Usage

1. Configure input/output directories in `appsettings.json`
2. Place your .doc or .docx files in the input directory
3. Run the application:
   ```bash
   dotnet run
   ```
   or execute the compiled executable:
   ```bash
   DocToPdfConverterLibreOffice.exe
   ```

## How It Works

The application uses LibreOffice's headless mode with the following command:
```bash
soffice --headless --convert-to pdf "input.doc" --outdir "output_directory"
```

### Conversion Process

1. **Validate Configuration**: Checks input/output directories and LibreOffice path
2. **Auto-Detection**: Attempts to find LibreOffice if path not configured
3. **Scan Files**: Finds all .doc and .docx files in input directory
4. **Convert**: Processes each file using LibreOffice CLI with 60-second timeout
5. **Report**: Displays conversion statistics

## Logging

Logs are written to:
- **Console**: Real-time progress and status
- **File**: `logs/converter-{date}.log` (rolls daily)

Log levels:
- `Information`: Normal operations
- `Warning`: Missing configuration, LibreOffice warnings
- `Error`: Conversion failures
- `Fatal`: Application-level failures
- `Debug`: LibreOffice command output (when debugging)

## Advantages Over COM Automation

| Feature | LibreOffice CLI | MS Word COM |
|---------|----------------|-------------|
| Microsoft Office Required | ? No | ? Yes |
| Cross-Platform | ? Yes | ? Windows only |
| License Cost | ? Free | ?? Paid |
| Server Deployment | ? Easy | ?? Complex |
| Process Isolation | ? Better | ?? Limited |
| Supports .docx | ? Yes | ? Yes |
| Conversion Speed | ?? Moderate | ? Fast |

## Output

The application will:
1. Auto-detect or validate LibreOffice installation
2. Scan the input directory for all .doc/.docx files
3. Convert each file to PDF format
4. Save PDFs in the output directory with the same filename
5. Log success/failure for each conversion with timeout protection
6. Display summary statistics at the end

## Error Handling

- Invalid/missing directories are validated before processing
- LibreOffice path validation with auto-detection fallback
- Each file conversion has a 60-second timeout
- Failed conversions are logged but don't stop the batch process
- Process output and errors are captured and logged
- Proper process cleanup and disposal

## Example Output

```
[2025-11-06 10:15:30.123] [INF] DOC to PDF Converter (LibreOffice) started
[2025-11-06 10:15:30.234] [INF] Auto-detected LibreOffice at: C:\Program Files\LibreOffice\program\soffice.exe
[2025-11-06 10:15:30.235] [INF] LibreOffice Path: C:\Program Files\LibreOffice\program\soffice.exe
[2025-11-06 10:15:30.236] [INF] Input Directory: C:\Input
[2025-11-06 10:15:30.237] [INF] Output Directory: C:\Output
[2025-11-06 10:15:30.450] [INF] Found 5 document file(s) to convert
[2025-11-06 10:15:30.451] [INF] Converting: document1.doc
[2025-11-06 10:15:32.890] [INF] Successfully converted: document1.doc -> document1.pdf
[2025-11-06 10:15:32.891] [INF] Converting: document2.docx
[2025-11-06 10:15:35.234] [INF] Successfully converted: document2.docx -> document2.pdf
[2025-11-06 10:15:35.235] [INF] Converting: report.doc
[2025-11-06 10:15:37.567] [INF] Successfully converted: report.doc -> report.pdf
[2025-11-06 10:15:37.568] [INF] Conversion completed. Success: 3, Failed: 0
```

## Troubleshooting

**Issue**: "LibreOffice installation not found"
- Install LibreOffice from https://www.libreoffice.org/download/
- Or manually configure the path in `appsettings.json`

**Issue**: "LibreOffice executable not found at: [path]"
- Verify LibreOffice is installed at the specified path
- Check the path matches your LibreOffice version
- Use auto-detection by removing or leaving the `LibreOfficePath` empty

**Issue**: "LibreOffice conversion timed out"
- Large or complex documents may take longer
- Consider increasing the timeout in the code (default: 60 seconds)
- Check if LibreOffice is hanging on system dialogs

**Issue**: Conversion fails for specific files
- Ensure the document isn't password protected
- Check if LibreOffice can open the file manually
- Review the log file for detailed error messages
- Some corrupted files may fail to convert

**Issue**: Permission errors
- Ensure the application has write access to the output directory
- Check that LibreOffice has read access to input files
- Run with appropriate permissions if accessing network paths

## Performance Tips

1. **Parallel Processing**: For large batches, consider modifying to use `Parallel.ForEach`
2. **Warm-up**: First conversion may be slower as LibreOffice initializes
3. **File Size**: Larger documents take longer to convert
4. **System Resources**: LibreOffice spawns separate processes for each conversion

## Command Line Arguments

Currently, the application reads configuration from `appsettings.json`. To support command-line arguments, you can extend the code to override configuration values.

## License

This project uses LibreOffice which is licensed under Mozilla Public License v2.0.
