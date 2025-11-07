using DocConverter.FileNameService.Models;
using ClosedXML.Excel;

namespace DocConverter.FileNameService;

public class ExcelFileNameConverterService : IFileNameConverterService
{
    private readonly Dictionary<int, ExcelDepartment> _departmentCache;
    private readonly string _excelFilePath;

    public ExcelFileNameConverterService(string? excelFilePath = null)
    {
        // If no path provided, look for Excel file next to the executable
        _excelFilePath = excelFilePath ?? Path.Combine(AppContext.BaseDirectory, "Departments.xlsx");
        _departmentCache = new Dictionary<int, ExcelDepartment>();
        LoadDepartmentInfo();
    }

    private void LoadDepartmentInfo()
    {
        try
        {
            if (!File.Exists(_excelFilePath))
            {
                // If Excel file doesn't exist, continue with empty cache
                // This allows the service to fall back to returning original filenames
                return;
            }

            using var workbook = new XLWorkbook(_excelFilePath);
            var worksheet = workbook.Worksheets.FirstOrDefault();
            
            if (worksheet == null)
            {
                return;
            }

            // Read data from Excel
            // Expected columns: DepartmentSectionNumber, DepartmentNumber, DepartmentAbbreviation, SectionAbbreviation
            var rowCount = worksheet.LastRowUsed()?.RowNumber() ?? 0;
            
            // Start from row 2 (assuming row 1 has headers)
            for (int row = 2; row <= rowCount; row++)
            {
                try
                {
                    var deptSectionNumber = worksheet.Cell(row, 1).GetValue<int>();
                    var deptNumber = worksheet.Cell(row, 2).GetValue<int>();
                    var deptAbbreviation = worksheet.Cell(row, 3).GetValue<string>();
                    var sectionAbbreviation = worksheet.Cell(row, 4).GetValue<string>();

                    var dept = new ExcelDepartment
                    {
                        DepartmentSectionNumber = deptSectionNumber,
                        DepartmentNumber = deptNumber,
                        DepartmentAbbreviation = deptAbbreviation,
                        SectionAbbreviation = sectionAbbreviation
                    };

                    _departmentCache[deptSectionNumber] = dept;
                }
                catch
                {
                    // Skip rows with invalid data
                    continue;
                }
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
            
            // Parse the format {DepartmentSectionNumber}-{DocNumber}
            var parts = fileNameWithoutExtension.Split('-');
            if (parts.Length != 2)
            {
                // Invalid format, return original filename
                return sourceFileName;
            }

            // Parse department section number and doc number
            if (!int.TryParse(parts[0], out int departmentSectionNumber) || 
                !int.TryParse(parts[1], out int docNumber))
            {
                // Invalid format, return original filename
                return sourceFileName;
            }

            // Look up department info
            if (!_departmentCache.TryGetValue(departmentSectionNumber, out var department))
            {
                // Department info not found, return original filename
                return sourceFileName;
            }

            // Sanitize department names to ensure valid filename
            var sanitizedDeptAbbrev = SanitizeForFileName(department.DepartmentAbbreviation);
            var sanitizedSectionAbbrev = SanitizeForFileName(department.SectionAbbreviation);

            // Build the new filename: {DepartmentAbbreviation}_{SectionAbbreviation}_{DepartmentSectionNumber}_{DocumentNumber}.pdf
            var newFileName = $"{sanitizedDeptAbbrev}_{sanitizedSectionAbbrev}_{departmentSectionNumber}_{docNumber}.pdf";
            
            return newFileName;
        }
        catch (Exception)
        {
            // Any exception during conversion, return original filename (or empty if null)
            return sourceFileName ?? string.Empty;
        }
    }
}
