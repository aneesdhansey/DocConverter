using Xunit;
using ClosedXML.Excel;
using System.IO;

namespace DocConverter.FileNameService.Tests;

public class ExcelFileNameConverterServiceTests
{
    private string CreateTestExcelFile()
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"TestDepartments_{Guid.NewGuid()}.xlsx");
        
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Departments");
            
            // Add headers
            worksheet.Cell(1, 1).Value = "DepartmentSectionNumber";
            worksheet.Cell(1, 2).Value = "DepartmentNumber";
            worksheet.Cell(1, 3).Value = "DepartmentAbbreviation";
            worksheet.Cell(1, 4).Value = "SectionAbbreviation";
            
            // Add test data
            worksheet.Cell(2, 1).Value = 123;
            worksheet.Cell(2, 2).Value = 10;
            worksheet.Cell(2, 3).Value = "ENG";
            worksheet.Cell(2, 4).Value = "MECH";
            
            worksheet.Cell(3, 1).Value = 456;
            worksheet.Cell(3, 2).Value = 20;
            worksheet.Cell(3, 3).Value = "HR";
            worksheet.Cell(3, 4).Value = "REC";
            
            workbook.SaveAs(tempPath);
        }
        
        return tempPath;
    }

    [Fact]
    public void GetConvertedFileName_WithValidFormatAndExcelData_ReturnsConvertedFileName()
    {
        // Arrange
        var excelPath = CreateTestExcelFile();
        try
        {
            var service = new ExcelFileNameConverterService(excelPath);

            // Act
            var result = service.GetConvertedFileName("123-456.doc");

            // Assert
            Assert.Equal("ENG_MECH_123_456.pdf", result);
        }
        finally
        {
            if (File.Exists(excelPath))
                File.Delete(excelPath);
        }
    }

    [Fact]
    public void GetConvertedFileName_WithInvalidFormat_ReturnsOriginalFileName()
    {
        // Arrange
        var excelPath = CreateTestExcelFile();
        try
        {
            var service = new ExcelFileNameConverterService(excelPath);

            // Act & Assert
            Assert.Equal("invalid.doc", service.GetConvertedFileName("invalid.doc"));
            Assert.Equal("123.doc", service.GetConvertedFileName("123.doc"));
            Assert.Equal("abc-def.doc", service.GetConvertedFileName("abc-def.doc"));
        }
        finally
        {
            if (File.Exists(excelPath))
                File.Delete(excelPath);
        }
    }

    [Fact]
    public void GetConvertedFileName_WithNonExistentDepartment_ReturnsOriginalFileName()
    {
        // Arrange
        var excelPath = CreateTestExcelFile();
        try
        {
            var service = new ExcelFileNameConverterService(excelPath);

            // Act
            var result = service.GetConvertedFileName("999-888.doc");

            // Assert
            Assert.Equal("999-888.doc", result);
        }
        finally
        {
            if (File.Exists(excelPath))
                File.Delete(excelPath);
        }
    }

    [Fact]
    public void GetConvertedFileName_WithNullInput_ReturnsEmptyString()
    {
        // Arrange
        var service = new ExcelFileNameConverterService("nonexistent.xlsx");

        // Act
        var result = service.GetConvertedFileName(null!);

        // Assert
        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void Constructor_WithNonExistentExcelFile_DoesNotThrow()
    {
        // Act & Assert - should not throw
        var service = new ExcelFileNameConverterService("nonexistent.xlsx");
        Assert.NotNull(service);
    }

    [Fact]
    public void GetConvertedFileName_WithMultipleDepartments_ReturnsCorrectFileName()
    {
        // Arrange
        var excelPath = CreateTestExcelFile();
        try
        {
            var service = new ExcelFileNameConverterService(excelPath);

            // Act
            var result1 = service.GetConvertedFileName("123-100.doc");
            var result2 = service.GetConvertedFileName("456-200.doc");

            // Assert
            Assert.Equal("ENG_MECH_123_100.pdf", result1);
            Assert.Equal("HR_REC_456_200.pdf", result2);
        }
        finally
        {
            if (File.Exists(excelPath))
                File.Delete(excelPath);
        }
    }
}
