using Microsoft.Extensions.Configuration;
using Serilog;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using System.Reflection;
using System.Collections.Concurrent;
using System.Diagnostics;
using DocConverter.FileNameService;

namespace DocToPdfConverterNetOffice;

class Program
{
    private static readonly object _logLock = new object();

    static void Main(string[] args)
    {
        // Build configuration - supporting both external and embedded appsettings.json
        var configuration = BuildConfiguration();

        // Configure Serilog
        ConfigureSerilog(configuration);

        try
        {
            Log.Information("DOC to PDF Converter started");

            // Get configuration
            var inputDirectory = configuration["Converter:InputDirectory"];
            var outputDirectory = configuration["Converter:OutputDirectory"];
            var enableParallel = configuration.GetValue<bool>("Converter:EnableParallelProcessing", false);
            var maxDegreeOfParallelism = configuration.GetValue<int>("Converter:MaxDegreeOfParallelism", 2);
            var chunkSize = configuration.GetValue<int>("Converter:ChunkSize", 50);
            var fileNamePatterns = configuration.GetSection("Converter:FileNamePatterns").Get<string[]>() ?? Array.Empty<string>();

            if (string.IsNullOrWhiteSpace(inputDirectory) || string.IsNullOrWhiteSpace(outputDirectory))
            {
                Log.Error("Input or output directory not configured in appsettings.json");
                return;
            }

            Log.Information("Input Directory: {InputDirectory}", inputDirectory);
            Log.Information("Output Directory: {OutputDirectory}", outputDirectory);
            Log.Information("Parallel Processing: {Enabled}", enableParallel ? "ENABLED" : "DISABLED");

            if (fileNamePatterns.Length > 0)
            {
                Log.Information("File Name Patterns: {Patterns}", string.Join(", ", fileNamePatterns));
            }
            else
            {
                Log.Information("File Name Patterns: None (processing all files)");
            }

            if (enableParallel)
            {
                Log.Warning("⚠️  WARNING: Parallel processing with COM automation can be unstable!");
                Log.Information("Max Degree of Parallelism: {MaxDegreeOfParallelism} (Recommended: 2-3)", maxDegreeOfParallelism);
                Log.Information("Chunk Size: {ChunkSize}", chunkSize);

                if (maxDegreeOfParallelism > 4)
                {
                    Log.Warning("⚠️  MaxDegreeOfParallelism > 4 is NOT recommended for COM automation!");
                    Log.Warning("⚠️  High risk of crashes, hangs, or memory issues!");
                }
            }

            // Validate input directory exists
            if (!Directory.Exists(inputDirectory))
            {
                Log.Error("Input directory does not exist: {InputDirectory}", inputDirectory);
                return;
            }

            // Create output directory if it doesn't exist
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
                Log.Information("Created output directory: {OutputDirectory}", outputDirectory);
            }

            // Get all .doc files from input directory
            var docFiles = Directory.GetFiles(inputDirectory, "*.doc", SearchOption.TopDirectoryOnly);

            // Apply file name patterns if specified
            if (fileNamePatterns.Length > 0)
            {
                var allFiles = docFiles;
                docFiles = docFiles.Where(file =>
                   {
                       var fileName = Path.GetFileName(file);
                       return fileNamePatterns.Any(pattern => MatchesPattern(fileName, pattern));
                   }).ToArray();

                Log.Information("Filtered from {TotalFiles} to {FilteredFiles} files using patterns", allFiles.Length, docFiles.Length);
            }

            if (docFiles.Length == 0)
            {
                Log.Warning("No .doc files found in input directory: {InputDirectory}", inputDirectory);
                return;
            }

            Log.Information("Found {Count} .doc file(s) to convert", docFiles.Length);

            // Initialize file name converter service based on configuration
            var fileNameConverterType = configuration.GetValue<string>("Converter:FileNameConverterType", "Database");
            IFileNameConverterService fileNameConverterService;
            
            if (fileNameConverterType.Equals("Excel", StringComparison.OrdinalIgnoreCase))
            {
                var excelFilePath = configuration.GetValue<string>("Converter:ExcelFilePath");
                fileNameConverterService = new ExcelFileNameConverterService(excelFilePath);
                Log.Information("File name converter service initialized (Excel mode)");
                if (!string.IsNullOrWhiteSpace(excelFilePath))
                {
                    Log.Information("Excel file path: {ExcelFilePath}", excelFilePath);
                }
            }
            else
            {
                fileNameConverterService = new FileNameConverterService(configuration);
                Log.Information("File name converter service initialized (Database mode)");
            }

            // Start conversion
            var stopwatch = Stopwatch.StartNew();
            List<ConversionResult> results;

            if (enableParallel)
            {
                results = ProcessFilesInParallel(docFiles, outputDirectory, maxDegreeOfParallelism, chunkSize, fileNameConverterService);
            }
            else
            {
                results = ProcessFilesSequentially(docFiles, outputDirectory, fileNameConverterService);
            }

            stopwatch.Stop();

            // Calculate statistics
            var successCount = results.Count(r => r.Success && !r.Skipped);
            var skippedCount = results.Count(r => r.Skipped);
            var failureCount = results.Count(r => !r.Success);
            var totalFiles = docFiles.Length;
            var avgTimePerFile = totalFiles > 0 ? stopwatch.Elapsed.TotalSeconds / totalFiles : 0;

            Log.Information("=====================================");
            Log.Information("Conversion completed in {Duration:F2} seconds", stopwatch.Elapsed.TotalSeconds);
            Log.Information("Success: {SuccessCount}, Skipped: {SkippedCount}, Failed: {FailureCount}, Total: {Total}",
                 successCount, skippedCount, failureCount, totalFiles);
            Log.Information("Average time per file: {AvgTime:F2} seconds", avgTimePerFile);

            if (stopwatch.Elapsed.TotalSeconds > 0)
            {
                Log.Information("Throughput: {FilesPerSecond:F2} files/second", totalFiles / stopwatch.Elapsed.TotalSeconds);
            }

            // Log failed files
            var failedFiles = results.Where(r => !r.Success).ToList();
            if (failedFiles.Any())
            {
                Log.Warning("Failed files ({Count}):", failedFiles.Count);
                foreach (var failed in failedFiles.Take(10))
                {
                    Log.Warning("  - {FileName}: {Error}", Path.GetFileName(failed.FilePath), failed.ErrorMessage);
                }
                if (failedFiles.Count > 10)
                {
                    Log.Warning("  ... and {Count} more", failedFiles.Count - 10);
                }
            }
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "Application terminated unexpectedly");
        }
        finally
        {
            Log.CloseAndFlush();
        }
    }

    static List<ConversionResult> ProcessFilesSequentially(string[] docFiles, string outputDirectory, IFileNameConverterService fileNameConverterService)
    {
        var results = new List<ConversionResult>();
        var totalFiles = docFiles.Length;

        Log.Information("Processing files sequentially (SAFE mode)");

        for (int i = 0; i < docFiles.Length; i++)
        {
            var docFile = docFiles[i];
            var result = ConvertDocToPdfSafe(docFile, outputDirectory, fileNameConverterService);
            results.Add(result);

            // Report progress every 5%
            var progress = (int)(((i + 1) * 100.0) / totalFiles);
            if (progress % 5 == 0 || i == totalFiles - 1)
            {
                Log.Information("Progress: {Progress}% ({Processed}/{Total} files)", progress, i + 1, totalFiles);
            }
        }

        return results;
    }

    static List<ConversionResult> ProcessFilesInParallel(string[] docFiles, string outputDirectory, int maxDegreeOfParallelism, int chunkSize, IFileNameConverterService fileNameConverterService)
    {
        var results = new ConcurrentBag<ConversionResult>();
        var totalFiles = docFiles.Length;
        var processedFiles = 0;
        var lastReportedProgress = 0;

        Log.Warning("⚠️  Starting PARALLEL conversion - Monitor for COM errors!");
        Log.Information("Using {MaxDegreeOfParallelism} concurrent Word instances", maxDegreeOfParallelism);

        // Clamp parallelism to safe limits
        if (maxDegreeOfParallelism > 4)
        {
            maxDegreeOfParallelism = 4;
            Log.Warning("Clamped MaxDegreeOfParallelism to 4 for safety");
        }

        var options = new ParallelOptions
        {
            MaxDegreeOfParallelism = maxDegreeOfParallelism
        };

        // Process in chunks to avoid memory issues
        var chunks = docFiles.Chunk(chunkSize).ToList();
        Log.Information("Processing {TotalFiles} files in {ChunkCount} chunks of {ChunkSize}", totalFiles, chunks.Count, chunkSize);

        for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++)
        {
            var chunk = chunks[chunkIndex];
            Log.Information("Processing chunk {ChunkNumber}/{TotalChunks} ({FileCount} files)",
      chunkIndex + 1, chunks.Count, chunk.Length);

            Parallel.ForEach(chunk, options, (docFile) =>
     {
         var result = ConvertDocToPdfSafe(docFile, outputDirectory, fileNameConverterService);
         results.Add(result);

         var currentProcessed = Interlocked.Increment(ref processedFiles);
         var currentProgress = (int)((currentProcessed * 100.0) / totalFiles);

         // Report progress every 5%
         if (currentProgress >= lastReportedProgress + 5 || currentProcessed == totalFiles)
         {
             lock (_logLock)
             {
                 if (currentProgress >= lastReportedProgress + 5 || currentProcessed == totalFiles)
                 {
                     lastReportedProgress = currentProgress;
                     Log.Information("Progress: {Progress}% ({Processed}/{Total} files)",
                      currentProgress, currentProcessed, totalFiles);
                 }
             }
         }
     });

            // Longer pause between chunks to allow COM cleanup
            if (chunkIndex < chunks.Count - 1)
            {
                Log.Debug("Pausing between chunks for COM cleanup...");
                Thread.Sleep(500); // Longer pause for COM stability
                GC.Collect(); // Force garbage collection to release COM objects
                GC.WaitForPendingFinalizers();
            }
        }

        return results.ToList();
    }

    static ConversionResult ConvertDocToPdfSafe(string docFile, string outputDirectory, IFileNameConverterService fileNameConverterService)
    {
        try
        {
            var sourceFileName = Path.GetFileName(docFile);
            var convertedFileName = fileNameConverterService.GetConvertedFileName(sourceFileName);
            var pdfFile = Path.Combine(outputDirectory, convertedFileName);

            // Smart skip logic: Check if PDF exists AND is newer than source file
            if (File.Exists(pdfFile))
            {
                var sourceModified = File.GetLastWriteTime(docFile);
                var pdfCreated = File.GetLastWriteTime(pdfFile);

                if (pdfCreated >= sourceModified)
                {
                    // PDF is up-to-date, skip conversion
                    Log.Debug("Skipping {FileName} - PDF is up-to-date (Source: {SourceTime}, PDF: {PdfTime})",
                          Path.GetFileName(docFile), sourceModified, pdfCreated);
                    return new ConversionResult
                    {
                        FilePath = docFile,
                        Success = true,
                        Skipped = true
                    };
                }
                else
                {
                    // Source file was modified after PDF creation, regenerate
                    Log.Information("Regenerating {FileName} - Source modified after PDF (Source: {SourceTime}, PDF: {PdfTime})",
             Path.GetFileName(docFile), sourceModified, pdfCreated);
                }
            }

            ConvertDocToPdf(docFile, pdfFile);

            return new ConversionResult
            {
                FilePath = docFile,
                Success = true
            };
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to convert: {FileName}", Path.GetFileName(docFile));
            return new ConversionResult
            {
                FilePath = docFile,
                Success = false,
                ErrorMessage = ex.Message
            };
        }
    }

    static void ConvertDocToPdf(string inputPath, string outputPath)
    {
        Word.Application wordApplication = null;
        Word.Document document = null;
        bool documentCreated = false;
        bool applicationCreated = false;

        try
        {
            // Create Word Application instance (each thread gets its own instance)
            wordApplication = new Word.Application
            {
                DisplayAlerts = WdAlertLevel.wdAlertsNone,
                Visible = false,
                ScreenUpdating = false // Disable screen updates for performance
            };
            applicationCreated = true;

            // Open the document
            document = wordApplication.Documents.Open(inputPath);
            documentCreated = true;

            // Export to PDF
            document.ExportAsFixedFormat(
outputPath,
     WdExportFormat.wdExportFormatPDF,
    false,
    WdExportOptimizeFor.wdExportOptimizeForPrint,
 WdExportRange.wdExportAllDocument,
     from: 0,
         to: 0,
         item: WdExportItem.wdExportDocumentContent,
          includeDocProps: true,
   keepIRM: true,
     createBookmarks: WdExportCreateBookmarks.wdExportCreateWordBookmarks,
        docStructureTags: true,
 bitmapMissingFonts: true,
       useISO19005_1: false
            );
        }
        finally
        {
            // Critical: Proper COM cleanup to prevent memory leaks and crashes
            // Close and dispose document first
            if (documentCreated && document != null)
            {
                try
                {
                    document.Close(false);
                }
                catch (Exception ex)
                {
                    Log.Debug(ex, "Error closing document (non-fatal)");
                }

                try
                {
                    document.Dispose();
                }
                catch (Exception ex)
                {
                    Log.Debug(ex, "Error disposing document (non-fatal)");
                }
            }

            // Quit and dispose application
            if (applicationCreated && wordApplication != null)
            {
                try
                {
                    wordApplication.Quit(false);
                }
                catch (Exception ex)
                {
                    Log.Debug(ex, "Error quitting Word (non-fatal)");
                }

                try
                {
                    wordApplication.Dispose();
                }
                catch (Exception ex)
                {
                    Log.Debug(ex, "Error disposing Word application (non-fatal)");
                }
            }

            // Note: NetOffice handles COM object release internally through Dispose()
            // Manual Marshal.ReleaseComObject is not needed and can cause issues
        }
    }

    static IConfiguration BuildConfiguration()
    {
        var builder = new ConfigurationBuilder();

        // First, try to load from external file (for normal deployment)
        var externalConfigPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
        if (File.Exists(externalConfigPath))
        {
            builder.AddJsonFile(externalConfigPath, optional: true, reloadOnChange: false);
        }
        else
        {
            // Fall back to embedded resource (for single-file publish)
            var assembly = Assembly.GetExecutingAssembly();
     var resourceName = "DocToPdfConverterNetOffice.appsettings.json";

            var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream != null)
            {
                // Don't use 'using' here - AddJsonStream will handle the stream
                builder.AddJsonStream(stream);
            }
            else
            {
                // If neither exists, create minimal configuration
                throw new FileNotFoundException("appsettings.json not found as external file or embedded resource");
            }
        }

        return builder.Build();
    }

    static void ConfigureSerilog(IConfiguration configuration)
    {
        try
        {
            // Try to configure from configuration file
            Log.Logger = new LoggerConfiguration()
        .ReadFrom.Configuration(configuration)
        .CreateLogger();
        }
        catch
        {
            // Fallback to code-based configuration if config file is missing or invalid
            Log.Logger = new LoggerConfiguration()
              .MinimumLevel.Information()
              .WriteTo.Console()
         .WriteTo.File(
                  path: Path.Combine("logs", "converter-.log"),
                 rollingInterval: RollingInterval.Day,
                 outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
           .CreateLogger();

            Log.Warning("Using default Serilog configuration");
        }
    }

    static bool MatchesPattern(string fileName, string pattern)
    {
        // Convert wildcard pattern to regex
        // * matches any characters, ? matches single character
        var regexPattern = "^" + System.Text.RegularExpressions.Regex.Escape(pattern)
       .Replace("\\*", ".*")
      .Replace("\\?", ".")
+ "$";

        return System.Text.RegularExpressions.Regex.IsMatch(fileName, regexPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    }
}

class ConversionResult
{
    public string FilePath { get; set; } = string.Empty;
    public bool Success { get; set; }
    public bool Skipped { get; set; }
    public string? ErrorMessage { get; set; }
}
