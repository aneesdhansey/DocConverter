using Microsoft.Extensions.Configuration;
using Serilog;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using System.Reflection;

namespace DocToPdfConverter;

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
            Log.Information("DOC to PDF Converter started");

            // Get input and output directories from configuration
            var inputDirectory = configuration["Converter:InputDirectory"];
            var outputDirectory = configuration["Converter:OutputDirectory"];

            if (string.IsNullOrWhiteSpace(inputDirectory) || string.IsNullOrWhiteSpace(outputDirectory))
            {
                Log.Error("Input or output directory not configured in appsettings.json");
                return;
            }

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

            // Get all .doc files from input directory
            var docFiles = Directory.GetFiles(inputDirectory, "*.doc", SearchOption.TopDirectoryOnly);

            if (docFiles.Length == 0)
            {
                Log.Warning("No .doc files found in input directory: {InputDirectory}", inputDirectory);
                return;
            }

            Log.Information("Found {Count} .doc file(s) to convert", docFiles.Length);

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

                    ConvertDocToPdf(docFile, pdfFile);

                    successCount++;
                    Log.Information("Successfully converted: {FileName} -> {OutputFileName}",
                Path.GetFileName(docFile), Path.GetFileName(pdfFile));
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
            var resourceName = "DocToPdfConverter.appsettings.json";

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

    static void ConvertDocToPdf(string inputPath, string outputPath)
    {
        Word.Application wordApplication = null;
        Word.Document document = null;

        try
        {
            // Create Word Application instance
            wordApplication = new Word.Application
            {
                DisplayAlerts = WdAlertLevel.wdAlertsNone,
                Visible = false
            };

            // Open the document
            document = wordApplication.Documents.Open(inputPath);

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
            // Clean up
            if (document != null)
            {
                document.Close(false);
                document.Dispose();
            }

            if (wordApplication != null)
            {
                wordApplication.Quit(false);
                wordApplication.Dispose();
            }
        }
    }
}
