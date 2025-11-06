# Single-File Publish - Quick Reference

## ? What Was Fixed

Both applications now work correctly with single-file, self-contained publishing.

## ?? Changes Made

### 1. Project Files (.csproj)
**Before:**
```xml
<None Update="appsettings.json">
  <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
</None>
```

**After:**
```xml
<EmbeddedResource Include="appsettings.json" />
```

### 2. Configuration Loading (Program.cs)

**Added:**
- `BuildConfiguration()` method - Loads config from external file or embedded resource
- `ConfigureSerilog()` method - Sets up logging with fallback
- `using System.Reflection;` - Required for embedded resource access

**Configuration Priority:**
1. External `appsettings.json` (if exists next to .exe)
2. Embedded `appsettings.json` (inside the .exe)
3. Code-based defaults (hardcoded fallback)

## ?? How to Publish

### Quick Command (Windows, Self-Contained)
```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

### Quick Command (Linux, Self-Contained)
```bash
dotnet publish -c Release -r linux-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

## ?? Usage After Publishing

### Option A: Use Embedded Config
Just run the .exe - uses embedded settings:
```bash
DocToPdfConverter.exe
```

### Option B: Use External Config
Create `appsettings.json` next to the .exe with your settings, then run:
```bash
DocToPdfConverter.exe
```

## ?? How It Works

```
???????????????????????????????????????????????
?  Application Starts   ?
???????????????????????????????????????????????
    ?
        ?
???????????????????????????????????????????????
?  Look for appsettings.json next to .exe     ?
???????????????????????????????????????????????
               ?
      ?????????????
       ?   Found?  ?
         ?????????????
       Yes ????????? No
         ?  ?
         ?     ?
    ??????????  ??????????????????????????
    ?  Use   ?  ? Load from embedded   ?
    ?External?  ? resource inside .exe   ?
    ??????????  ??????????????????????????
         ?       ?
      ??????????????
     ?
        ?
    ??????????????????
       ? Configure      ?
       ? Serilog        ?
 ??????????????????
        ?
   ?
  ??????????????????
       ? Run Application?
       ??????????????????
```

## ? Benefits

| Feature | Before | After |
|---------|--------|-------|
| **Single .exe distribution** | ? Required appsettings.json alongside | ? Works standalone |
| **Configuration flexibility** | ? Fixed | ? Can override with external file |
| **Self-contained deployment** | ?? Broke without external config | ? Fully self-contained |
| **Size** | ~500 KB + .NET runtime | ~65 MB (includes everything) |

## ?? Configuration Examples

### Embedded (Default)
Edit `appsettings.json` before publishing:
```json
{
  "Converter": {
    "InputDirectory": "C:\\Input",
    "OutputDirectory": "C:\\Output"
  }
}
```

### External Override
Create next to .exe after publishing:
```json
{
  "Converter": {
    "InputDirectory": "D:\\MyDocs",
    "OutputDirectory": "D:\\MyPDFs"
  }
}
```

## ?? Quick Start

```bash
# 1. Publish
cd DocToPdfConverter
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true

# 2. Find the .exe
cd bin\Release\net8.0\win-x64\publish

# 3. (Optional) Create custom config
echo { "Converter": { "InputDirectory": "C:\\Docs", "OutputDirectory": "C:\\PDFs" } } > appsettings.json

# 4. Run
.\DocToPdfConverter.exe
```

## ?? Troubleshooting

| Problem | Solution |
|---------|----------|
| Config not found error | Clean and rebuild: `dotnet clean && dotnet build` |
| Using wrong config | Delete external `appsettings.json` to use embedded |
| Settings not applied | Check JSON syntax in external file |
| Large file size | Use framework-dependent: remove `--self-contained` |

## ?? File Size Comparison

```
Self-Contained + Single-File:    ~65-80 MB  ? No runtime needed
Framework-Dependent + Single:    ~5-10 MB   ?? Needs .NET 8
Normal Publish (multi-file):     ~500 KB    ?? Needs .NET 8 + all files
```

## ? Testing Checklist

- [ ] Publish as single-file
- [ ] Run without external appsettings.json (uses embedded)
- [ ] Create external appsettings.json (uses external)
- [ ] Modify external config and verify changes take effect
- [ ] Delete external config and verify fallback to embedded
- [ ] Check logs are created correctly
- [ ] Verify conversion works end-to-end

## ?? Key Files Modified

1. `DocToPdfConverter/DocToPdfConverter.csproj`
2. `DocToPdfConverter/Program.cs`
3. `DocToPdfConverterLibreOffice/DocToPdfConverterLibreOffice.csproj`
4. `DocToPdfConverterLibreOffice/Program.cs`

All changes maintain backward compatibility with normal (multi-file) publishing.
