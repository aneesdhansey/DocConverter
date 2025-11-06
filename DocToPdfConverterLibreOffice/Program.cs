using Microsoft.Extensions.Configuration;
using Serilog;
using System.Diagnostics;
using System.Text;
using System.Reflection;

namespace DocToPdfConverterLibreOffice;

class Program
{
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

            if (docFiles.Length == 0)
            {
                Log.Warning("No .doc or .docx files found in input directory: {InputDirectory}", inputDirectory);
                return;
            }

            Log.Information("Found {Count} document file(s) to convert", docFiles.Length);

            // Convert each file
            int successCount = 0;
            int failureCount = 0;

            foreach (var docFile in docFiles)
            {
                try
                {
                    var fileName = Path.GetFileNameWithoutExtension(docFile);
                    var pdfFile = Path.Combine(outputDirectory, $"{fileName}.pdf");

                    Log.Information("Converting: {FileName}", Path.GetFileName(docFile));

                    ConvertDocToPdf(libreOfficePath, docFile, outputDirectory);

                    successCount++;
                    Log.Information("Successfully converted: {FileName} -> {OutputFileName}",
                        Path.GetFileName(docFile), $"{fileName}.pdf");
                }
                catch (Exception ex)
                {
                    failureCount++;
                    Log.Error(ex, "Failed to convert: {FileName}", Path.GetFileName(docFile));
                }
            }

            Log.Information("Conversion completed. Success: {SuccessCount}, Failed: {FailureCount}",
                successCount, failureCount);
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

    static void ConvertDocToPdf(string libreOfficePath, string inputPath, string outputDirectory)
    {
        try
        {
            // LibreOffice command line arguments:
            // --headless: Run without GUI
            // --convert-to pdf: Convert to PDF format
            // --outdir: Specify output directory
            var arguments = $"--headless --convert-to pdf \"{inputPath}\" --outdir \"{outputDirectory}\"";

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
                  Log.Warning("LibreOffice error: {Error}", e.Data);
              }
          };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            // Wait for conversion to complete (with timeout)
            if (!process.WaitForExit(60000)) // 60 second timeout
            {
                process.Kill();
                throw new TimeoutException("LibreOffice conversion timed out after 60 seconds");
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
}
