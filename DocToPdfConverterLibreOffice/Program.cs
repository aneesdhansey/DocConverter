using Microsoft.Extensions.Configuration;
using Serilog;
using System.Diagnostics;
using System.Text;
using System.Reflection;
using System.Collections.Concurrent;

namespace DocToPdfConverterLibreOffice;

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
            Log.Information("DOC to PDF Converter (LibreOffice) started");

            // Get configuration
            var inputDirectory = configuration["Converter:InputDirectory"];
            var outputDirectory = configuration["Converter:OutputDirectory"];
            var libreOfficePath = configuration["Converter:LibreOfficePath"];
            var maxDegreeOfParallelism = configuration.GetValue<int>("Converter:MaxDegreeOfParallelism", 4);
            var chunkSize = configuration.GetValue<int>("Converter:ChunkSize", 100);
            var timeoutSeconds = configuration.GetValue<int>("Converter:TimeoutSeconds", 60);
            var fileNamePatterns = configuration.GetSection("Converter:FileNamePatterns").Get<string[]>() ?? Array.Empty<string>();

            if (string.IsNullOrWhiteSpace(inputDirectory) || string.IsNullOrWhiteSpace(outputDirectory))
            {
                Log.Error("Input or output directory not configured in appsettings.json");
                return;
            }

            if (string.IsNullOrWhiteSpace(libreOfficePath))
            {
                Log.Warning("LibreOffice path not configured. Attempting to auto-detect...");
                libreOfficePath = FindLibreOffice();

                if (string.IsNullOrWhiteSpace(libreOfficePath))
                {
                    Log.Error("LibreOffice installation not found. Please install LibreOffice or configure the path in appsettings.json");
                    return;
                }
            }

            if (!File.Exists(libreOfficePath))
            {
                Log.Error("LibreOffice executable not found at: {LibreOfficePath}", libreOfficePath);
                return;
            }

            Log.Information("LibreOffice Path: {LibreOfficePath}", libreOfficePath);
            Log.Information("Input Directory: {InputDirectory}", inputDirectory);
            Log.Information("Output Directory: {OutputDirectory}", outputDirectory);
            Log.Information("Max Degree of Parallelism: {MaxDegreeOfParallelism}", maxDegreeOfParallelism);
            Log.Information("Chunk Size: {ChunkSize}", chunkSize);
            Log.Information("Timeout: {TimeoutSeconds} seconds", timeoutSeconds);

            if (fileNamePatterns.Length > 0)
            {
                Log.Information("File Name Patterns: {Patterns}", string.Join(", ", fileNamePatterns));
            }
            else
            {
                Log.Information("File Name Patterns: None (processing all files)");
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

            // Get all .doc and .docx files from input directory
            var docFiles = Directory.GetFiles(inputDirectory, "*.doc", SearchOption.TopDirectoryOnly)
        .Concat(Directory.GetFiles(inputDirectory, "*.docx", SearchOption.TopDirectoryOnly))
      .ToArray();

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
                Log.Warning("No .doc or .docx files found in input directory: {InputDirectory}", inputDirectory);
                return;
            }

            Log.Information("Found {Count} document file(s) to convert", docFiles.Length);

            // Start conversion with parallel processing
            var stopwatch = Stopwatch.StartNew();
            var results = ProcessFilesInParallel(docFiles, outputDirectory, libreOfficePath, maxDegreeOfParallelism, chunkSize, timeoutSeconds);
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

    static List<ConversionResult> ProcessFilesInParallel(string[] docFiles, string outputDirectory, string libreOfficePath,
    int maxDegreeOfParallelism, int chunkSize, int timeoutSeconds)
    {
        var results = new ConcurrentBag<ConversionResult>();
        var totalFiles = docFiles.Length;
        var processedFiles = 0;
        var lastReportedProgress = 0;

        Log.Information("Starting parallel conversion with {MaxDegreeOfParallelism} concurrent processes", maxDegreeOfParallelism);

        var options = new ParallelOptions
        {
            MaxDegreeOfParallelism = maxDegreeOfParallelism
        };

        // Process in chunks to avoid overwhelming the system
        var chunks = docFiles.Chunk(chunkSize).ToList();
        Log.Information("Processing {TotalFiles} files in {ChunkCount} chunks of {ChunkSize}", totalFiles, chunks.Count, chunkSize);

        for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++)
        {
            var chunk = chunks[chunkIndex];
            Log.Information("Processing chunk {ChunkNumber}/{TotalChunks} ({FileCount} files)",
      chunkIndex + 1, chunks.Count, chunk.Length);

            Parallel.ForEach(chunk, options, (docFile) =>
                 {
                     var result = ConvertDocToPdfSafe(libreOfficePath, docFile, outputDirectory, timeoutSeconds);
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

            // Small pause between chunks to allow system to breathe
            if (chunkIndex < chunks.Count - 1)
            {
                Thread.Sleep(100);
            }
        }

        return results.ToList();
    }

    static ConversionResult ConvertDocToPdfSafe(string libreOfficePath, string docFile, string outputDirectory, int timeoutSeconds)
    {
        try
        {
            var fileName = Path.GetFileNameWithoutExtension(docFile);
            var pdfFile = Path.Combine(outputDirectory, $"{fileName}.pdf");


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

            ConvertDocToPdf(libreOfficePath, docFile, outputDirectory, timeoutSeconds);

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

    static void ConvertDocToPdf(string libreOfficePath, string inputPath, string outputDirectory, int timeoutSeconds)
    {
        // Create a unique temporary profile directory for this conversion
        // This prevents conflicts when running multiple LibreOffice instances in parallel
        var tempProfileDir = Path.Combine(Path.GetTempPath(), $"LibreOfficeProfile_{Guid.NewGuid()}");

        try
        {
            Directory.CreateDirectory(tempProfileDir);

            // LibreOffice command line arguments:
            // --headless: Run without GUI
            // --convert-to pdf: Convert to PDF format
            // --outdir: Specify output directory
            // -env:UserInstallation: Use unique profile directory to avoid conflicts
            var arguments = $"--headless --convert-to pdf \"{inputPath}\" --outdir \"{outputDirectory}\" -env:UserInstallation=file:///{tempProfileDir.Replace("\\", "/")}";
            var processStartInfo = new ProcessStartInfo
            {
                FileName = libreOfficePath,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = outputDirectory
            };

            using var process = new Process { StartInfo = processStartInfo };

            var outputBuilder = new StringBuilder();
            var errorBuilder = new StringBuilder();

            process.OutputDataReceived += (sender, e) =>
          {
              if (!string.IsNullOrEmpty(e.Data))
              {
                  outputBuilder.AppendLine(e.Data);
                  Log.Debug("LibreOffice output: {Output}", e.Data);
              }
          };

            process.ErrorDataReceived += (sender, e) =>
          {
              if (!string.IsNullOrEmpty(e.Data))
              {
                  errorBuilder.AppendLine(e.Data);
                  Log.Debug("LibreOffice error: {Error}", e.Data);
              }
          };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            // Wait for conversion to complete (with configurable timeout)
            if (!process.WaitForExit(timeoutSeconds * 1000))
            {
                process.Kill();
                throw new TimeoutException($"LibreOffice conversion timed out after {timeoutSeconds} seconds");
            }

            if (process.ExitCode != 0)
            {
                var errorMessage = errorBuilder.Length > 0
      ? errorBuilder.ToString()
             : $"LibreOffice exited with code {process.ExitCode}";
                throw new Exception($"Conversion failed: {errorMessage}");
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error during LibreOffice conversion: {ex.Message}", ex);
        }
        finally
        {
            // Clean up temporary profile directory
            try
            {
                if (Directory.Exists(tempProfileDir))
                {
                    Directory.Delete(tempProfileDir, recursive: true);
                }
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Failed to delete temporary profile directory: {TempProfileDir}", tempProfileDir);
                // Non-fatal, just log it
            }
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
            var resourceName = "DocToPdfConverterLibreOffice.appsettings.json";

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

    static string? FindLibreOffice()
    {
        // Common LibreOffice installation paths on Windows
        var possiblePaths = new[]
             {
         @"C:\Program Files\LibreOffice\program\soffice.exe",
     @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
   @"C:\Program Files\LibreOffice 7\program\soffice.exe",
     @"C:\Program Files (x86)\LibreOffice 7\program\soffice.exe"
     };

        foreach (var path in possiblePaths)
        {
            if (File.Exists(path))
            {
                Log.Information("Auto-detected LibreOffice at: {Path}", path);
                return path;
            }
        }

        // Try to find via environment PATH
        try
        {
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "where",
                    Arguments = "soffice.exe",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                }
            };

            process.Start();
            var output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();

            if (process.ExitCode == 0 && !string.IsNullOrWhiteSpace(output))
            {
                var path = output.Split('\n').FirstOrDefault()?.Trim();
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    Log.Information("Found LibreOffice in PATH: {Path}", path);
                    return path;
                }
            }
        }
        catch
        {
            // Ignore errors during PATH search
        }

        return null;
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
