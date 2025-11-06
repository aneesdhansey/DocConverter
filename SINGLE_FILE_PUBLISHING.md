# Single-File Publishing Guide

Both DocToPdfConverter projects now support single-file, self-contained executable publishing.

## What Changed

### Configuration Loading Strategy

The applications now use a **three-tier fallback strategy**:

1. **External appsettings.json** (next to the .exe) - Checked first
2. **Embedded appsettings.json** (inside the .exe) - Used if external file not found
3. **Code-based defaults** - Fallback if both above fail

### Technical Changes

- `appsettings.json` is now embedded as a resource in the executable
- Configuration loading checks for external file first, then embedded resource
- Serilog has a fallback code-based configuration

## Benefits

? **Single-File Distribution**: One .exe file contains everything
? **Flexible Configuration**: Can still use external config for easy updates
? **Self-Contained**: No .NET runtime installation required
? **Works Everywhere**: Same binary works with or without external config

## Publishing Commands

### Windows x64 (Single-File, Self-Contained)

```bash
# DocToPdfConverter (NetOffice version)
dotnet publish DocToPdfConverter/DocToPdfConverter.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true

# DocToPdfConverterLibreOffice
dotnet publish DocToPdfConverterLibreOffice/DocToPdfConverterLibreOffice.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

### Linux x64 (Single-File, Self-Contained)

```bash
# DocToPdfConverterLibreOffice (Linux compatible)
dotnet publish DocToPdfConverterLibreOffice/DocToPdfConverterLibreOffice.csproj -c Release -r linux-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

### Framework-Dependent (Smaller size, requires .NET 8)

```bash
# Windows x64
dotnet publish DocToPdfConverter/DocToPdfConverter.csproj -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true

# Linux x64
dotnet publish DocToPdfConverterLibreOffice/DocToPdfConverterLibreOffice.csproj -c Release -r linux-x64 --self-contained false -p:PublishSingleFile=true
```

## Output Locations

After publishing, find your executable at:
```
bin/Release/net8.0/{runtime}/publish/
```

For example:
- Windows: `bin/Release/net8.0/win-x64/publish/DocToPdfConverter.exe`
- Linux: `bin/Release/net8.0/linux-x64/publish/DocToPdfConverterLibreOffice`

## Configuration Options

### Option 1: Use Embedded Configuration (Easiest)

Just run the .exe file:
```bash
DocToPdfConverter.exe
```

The application will use the embedded `appsettings.json` with default settings.

**To change settings**: Edit `appsettings.json` before publishing and rebuild.

### Option 2: Use External Configuration (Most Flexible)

Create an `appsettings.json` file next to the .exe:

```json
{
  "Converter": {
    "InputDirectory": "C:\\MyDocuments",
    "OutputDirectory": "C:\\MyPDFs"
  },
  "Serilog": {
    "MinimumLevel": "Information",
    "WriteTo": [
      {
        "Name": "Console"
      },
      {
        "Name": "File",
        "Args": {
          "path": "logs/converter-.log",
          "rollingInterval": "Day"
        }
   }
    ]
  }
}
```

Then run the .exe - it will use the external file.

**Benefits**: 
- Change configuration without rebuilding
- Different configs for different environments
- Easy to deploy updates

## File Size Comparison

| Publish Type | Approximate Size |
|-------------|-----------------|
| Self-Contained, Single-File | ~65-80 MB |
| Framework-Dependent, Single-File | ~5-10 MB |
| Normal Publish (multiple files) | ~500 KB (requires .NET 8) |

## Trimming (Reduce Size)

For even smaller executables, enable trimming:

```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishTrimmed=true
```

?? **Warning**: Trimming may cause issues with reflection-based libraries. Test thoroughly!

## Platform-Specific Notes

### Windows
- NetOfficeFw version only works on Windows (requires MS Word COM)
- LibreOffice version works on Windows if LibreOffice is installed

### Linux
- Only LibreOffice version is compatible
- Requires LibreOffice installed: `sudo apt-get install libreoffice`
- Make executable: `chmod +x DocToPdfConverterLibreOffice`

### macOS
- LibreOffice version should work (not tested)
- Publish with `-r osx-x64` or `-r osx-arm64`
- Requires LibreOffice: `brew install --cask libreoffice`

## Deployment Checklist

### For NetOfficeFw Version (Windows):
- ? Publish as single-file executable
- ? Ensure Microsoft Word is installed on target machine
- ? (Optional) Place appsettings.json next to .exe for custom config
- ? Ensure target directories exist or app can create them

### For LibreOffice Version (Any OS):
- ? Publish as single-file executable for target platform
- ? Ensure LibreOffice is installed on target machine
- ? (Optional) Place appsettings.json next to .exe for custom config
- ? (Linux/Mac) Set executable permissions

## Example: Complete Publishing Workflow

```bash
# 1. Navigate to solution directory
cd C:\Users\anees\source\repos\DocConverter

# 2. Clean previous builds
dotnet clean

# 3. Publish (Windows x64, self-contained, single-file)
dotnet publish DocToPdfConverter/DocToPdfConverter.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true

# 4. Navigate to output
cd DocToPdfConverter\bin\Release\net8.0\win-x64\publish

# 5. (Optional) Create external appsettings.json for custom config
notepad appsettings.json

# 6. Run the application
.\DocToPdfConverter.exe
```

## Docker Deployment (LibreOffice Version)

Example Dockerfile:

```dockerfile
FROM mcr.microsoft.com/dotnet/runtime:8.0

# Install LibreOffice
RUN apt-get update && \
    apt-get install -y libreoffice && \
  apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Copy the single-file executable
COPY DocToPdfConverterLibreOffice /app/DocToPdfConverterLibreOffice
RUN chmod +x /app/DocToPdfConverterLibreOffice

# Optional: Copy external config
COPY appsettings.json /app/appsettings.json

WORKDIR /app
ENTRYPOINT ["./DocToPdfConverterLibreOffice"]
```

## Troubleshooting

### Issue: "appsettings.json not found as external file or embedded resource"
- **Cause**: Build didn't embed the resource properly
- **Solution**: Clean and rebuild: `dotnet clean && dotnet build`

### Issue: Configuration not loading
- **Check**: Is there an external `appsettings.json` next to the .exe?
- **Check**: Is it valid JSON?
- **Fallback**: App will use code-based defaults and log a warning

### Issue: Single-file publish fails
- **Check**: .NET 8 SDK is installed
- **Check**: Correct runtime identifier (win-x64, linux-x64, etc.)
- **Try**: Remove `obj` and `bin` folders and rebuild

### Issue: Large executable size
- **Try**: Framework-dependent publish (requires .NET 8 on target)
- **Try**: Enable trimming (test thoroughly first)
- **Accept**: Self-contained includes the entire .NET runtime

## Best Practices

1. **Development**: Use external `appsettings.json` for easy testing
2. **Production**: Embed defaults, allow external override for environment-specific settings
3. **Distribution**: Include sample `appsettings.json` with download
4. **Updates**: If only config changes, users can update just the JSON file
5. **Security**: Don't embed sensitive data; use external config or environment variables

## Related Documentation

- [.NET Single-File Deployment](https://learn.microsoft.com/en-us/dotnet/core/deploying/single-file)
- [Publishing .NET Apps](https://learn.microsoft.com/en-us/dotnet/core/deploying/)
- [Runtime Identifiers](https://learn.microsoft.com/en-us/dotnet/core/rid-catalog)
