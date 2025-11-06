# DOC to PDF Converter - Comparison Guide

This repository contains two implementations for converting DOC/DOCX files to PDF:

## 1. DocToPdfConverter (NetOfficeFw.Word)
Uses Microsoft Word COM automation via NetOfficeFw.Word library.

## 2. DocToPdfConverterLibreOffice (LibreOffice CLI)
Uses LibreOffice command-line interface in headless mode.

---

## Feature Comparison

| Feature | NetOfficeFw.Word | LibreOffice CLI |
|---------|------------------|-----------------|
| **Microsoft Office Required** | ? Yes | ? No |
| **Free/Open Source** | ? No (requires Office license) | ? Yes |
| **Cross-Platform** | ? Windows only | ? Windows, Linux, macOS |
| **Conversion Quality** | ?? Excellent (native) | ? Very Good |
| **Speed** | ?? Fast | ? Moderate |
| **Server Deployment** | ?? Complex (Office licensing issues) | ? Easy |
| **Memory Usage** | ?? Higher (COM overhead) | ? Lower (process-based) |
| **Process Isolation** | ?? Limited | ? Excellent |
| **File Format Support** | .doc only | .doc and .docx |
| **Setup Complexity** | ?? Requires Office installation | ? Just install LibreOffice |
| **Stability** | ?? COM can be fragile | ? More stable |
| **Concurrent Processing** | ?? Limited (COM threading) | ? Better (separate processes) |

---

## When to Use Each Version

### Use NetOfficeFw.Word When:

? **You already have Microsoft Office installed**
- Office is already part of your infrastructure
- No additional software installation needed

? **You need the highest fidelity conversion**
- Using native Word engine ensures perfect formatting
- Complex documents with advanced features

? **You're working exclusively on Windows**
- Desktop applications
- Windows-only environments

? **Speed is critical**
- Fastest conversion times
- COM overhead is acceptable for your use case

? **Avoid When:**
- Deploying to servers (licensing issues)
- Need cross-platform support
- Working in containerized environments
- Processing high volumes concurrently

---

### Use LibreOffice CLI When:

? **Cross-platform requirement**
- Need to run on Windows, Linux, or macOS
- Docker containers or Kubernetes

? **Server/Production deployment**
- No Microsoft Office licensing concerns
- Easy to automate and scale

? **Cost is a concern**
- Free and open-source
- No licensing fees

? **High concurrency needed**
- Process isolation allows better parallel processing
- More stable under load

? **Both .doc and .docx support needed**
- Handles both formats out of the box

? **Working with cloud infrastructure**
- AWS, Azure, GCP Linux instances
- Serverless functions (with layers)

? **Avoid When:**
- Need absolute highest conversion fidelity
- Already invested in Microsoft ecosystem
- Converting very large batches where speed is critical

---

## Performance Comparison

### Conversion Time (approximate)

| Document Size | NetOfficeFw.Word | LibreOffice CLI |
|--------------|------------------|-----------------|
| Small (< 1MB) | ~1-2 seconds | ~2-3 seconds |
| Medium (1-5MB) | ~2-5 seconds | ~4-8 seconds |
| Large (> 5MB) | ~5-15 seconds | ~10-30 seconds |

*Times vary based on document complexity and system specifications*

### Resource Usage

| Metric | NetOfficeFw.Word | LibreOffice CLI |
|--------|------------------|-----------------|
| Initial Memory | ~150-200 MB | ~100-150 MB |
| Per Conversion | ~50-100 MB | ~80-120 MB |
| CPU Usage | Moderate | Moderate-High |
| Startup Time | Slow (COM init) | Fast (direct process) |

---

## Deployment Scenarios

### Desktop Application
**Recommendation**: NetOfficeFw.Word
- Users likely have Office installed
- Optimal performance
- Native Windows integration

### Web Application (IIS/Windows Server)
**Recommendation**: LibreOffice CLI
- Avoid COM licensing issues
- Better process isolation
- Easier to scale

### Docker/Containers
**Recommendation**: LibreOffice CLI
- No Office licensing in containers
- Native Linux support
- Smaller image size

### Azure/AWS/Cloud
**Recommendation**: LibreOffice CLI
- Most cloud VMs run Linux
- Cost-effective
- Easy to scale horizontally

### Enterprise Server (Windows)
**Recommendation**: Depends
- If Office licenses available: Either
- If no Office: LibreOffice CLI
- Consider licensing costs vs. performance needs

---

## Code Differences

### NetOfficeFw.Word Approach
```csharp
// Uses COM automation
var wordApp = new Word.Application();
var doc = wordApp.Documents.Open(inputPath);
doc.ExportAsFixedFormat(outputPath, WdExportFormat.wdExportFormatPDF);
```

### LibreOffice CLI Approach
```csharp
// Uses command-line process
var process = new Process
{
    StartInfo = new ProcessStartInfo
    {
        FileName = "soffice",
        Arguments = "--headless --convert-to pdf input.doc --outdir output"
    }
};
process.Start();
process.WaitForExit();
```

---

## Installation Requirements

### NetOfficeFw.Word
1. Microsoft Office (Word) installed
2. NetOfficeFw.Word NuGet package
3. Windows OS

### LibreOffice CLI
1. LibreOffice installed (free download)
2. No special packages required
3. Any OS (Windows, Linux, macOS)

---

## Recommendation Matrix

| Your Scenario | Recommended Solution |
|--------------|---------------------|
| Desktop app with Office installed | **NetOfficeFw.Word** |
| Web app on Windows | **LibreOffice CLI** |
| Web app on Linux | **LibreOffice CLI** |
| Docker container | **LibreOffice CLI** |
| High-volume server | **LibreOffice CLI** |
| Maximum fidelity needed | **NetOfficeFw.Word** |
| Cost-sensitive project | **LibreOffice CLI** |
| Quick prototype | **LibreOffice CLI** |
| Enterprise with Office | **NetOfficeFw.Word** |
| Multi-platform support | **LibreOffice CLI** |

---

## Migration Between Versions

Both implementations share the same configuration structure and logging approach, making it easy to switch:

1. **Same Configuration**: Both use `appsettings.json` with similar structure
2. **Same Logging**: Both use Serilog with identical configuration
3. **Same Interface**: Input/output directories work the same way
4. **Similar Output**: Both produce similar PDF quality

To switch:
1. Update the project reference
2. Adjust the `LibreOfficePath` configuration
3. No code changes needed in calling applications

---

## Conclusion

- **NetOfficeFw.Word**: Best for Windows environments with Office already installed
- **LibreOffice CLI**: Best for servers, cross-platform, and cost-sensitive deployments

Both implementations are production-ready with proper error handling, logging, and configuration management.
