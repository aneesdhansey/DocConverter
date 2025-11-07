using Microsoft.Data.SqlClient;
using Dapper;
using DocConverter.FileNameService.Models;
using Microsoft.Extensions.Configuration;

namespace DocConverter.FileNameService;

public class FileNameConverterService : IFileNameConverterService
{
    private readonly Dictionary<int, Department> _departmentCache;
    private readonly IConfiguration _configuration;

    public FileNameConverterService(IConfiguration configuration)
    {
        _configuration = configuration;
        _departmentCache = new Dictionary<int, Department>();
        LoadDepartmentInfo();
    }

    private void LoadDepartmentInfo()
    {
        try
        {
            var connectionString = _configuration.GetConnectionString("DefaultConnection");
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                // If no connection string is configured, continue with empty cache
                // This allows the service to fall back to returning original filenames
                return;
            }

            using var connection = new SqlConnection(connectionString);
            connection.Open();

            var departments = connection.Query<Department>(
                "SELECT DepartmentSectionId, DepartmentName, DepartmentSection FROM Departments"
            );

            foreach (var dept in departments)
            {
                _departmentCache[dept.DepartmentSectionId] = dept;
            }
        }
        catch (Exception)
        {
            // If there's any error loading department info, continue with empty cache
            // This allows the service to fall back to returning original filenames
        }
    }

    private static string SanitizeForFileName(string input)
    {
        if (string.IsNullOrEmpty(input))
        {
            return input;
        }

        // Replace invalid filename characters with underscore
        var invalidChars = Path.GetInvalidFileNameChars();
        foreach (var c in invalidChars)
        {
            input = input.Replace(c, '_');
        }

        return input;
    }

    public string GetConvertedFileName(string sourceFileName)
    {
        try
        {
            // Handle null or empty input
            if (string.IsNullOrWhiteSpace(sourceFileName))
            {
                return sourceFileName ?? string.Empty;
            }

            // Extract the file name without extension
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(sourceFileName);
            
            // Parse the format XXX-YYYY
            var parts = fileNameWithoutExtension.Split('-');
            if (parts.Length != 2)
            {
                // Invalid format, return original filename
                return sourceFileName;
            }

            // Parse department section ID (XXX) and doc number (YYYY)
            if (!int.TryParse(parts[0], out int departmentSectionId) || 
                !int.TryParse(parts[1], out int docNumber))
            {
                // Invalid format, return original filename
                return sourceFileName;
            }

            // Look up department info
            if (!_departmentCache.TryGetValue(departmentSectionId, out var department))
            {
                // Department info not found, return original filename
                return sourceFileName;
            }

            // Sanitize department names to ensure valid filename
            var sanitizedDepartmentName = SanitizeForFileName(department.DepartmentName);
            var sanitizedDepartmentSection = SanitizeForFileName(department.DepartmentSection);

            // Build the new filename: {DepartmentName}_{DepartmentSection}_{XXX}_{YYYY}.pdf
            var newFileName = $"{sanitizedDepartmentName}_{sanitizedDepartmentSection}_{departmentSectionId}_{docNumber}.pdf";
            
            return newFileName;
        }
        catch (Exception)
        {
            // Any exception during conversion, return original filename (or empty if null)
            return sourceFileName ?? string.Empty;
        }
    }
}
