# DocConverter.FileNameService

A library for converting document file names based on department information from either a database or Excel file.

## Overview

This service extracts department section IDs from source file names and converts them to a more descriptive format using department information from either a SQL Server database or an Excel file.

## Features

**Two Service Implementations:**

1. **Database-driven (`FileNameConverterService`)**
   - Parses filenames in the format `XXX-YYYY` (with any extension) where:
     - `XXX` is the department section ID
     - `YYYY` is the document number
   - Converts to: `{DepartmentName}_{DepartmentSection}_{XXX}_{YYYY}.pdf`
   - Uses Dapper for efficient database access

2. **Excel-driven (`ExcelFileNameConverterService`)**
   - Parses filenames in the format `XXX-YYYY` (with any extension) where:
     - `XXX` is the department section number
     - `YYYY` is the document number
   - Converts to: `{DepartmentAbbreviation}_{SectionAbbreviation}_{XXX}_{YYYY}.pdf`
   - Uses ClosedXML for Excel file reading

**Common Features:**
- Sanitizes department names to remove invalid filename characters
- Loads all department info at startup for performance
- Graceful error handling - returns original filename on any error

## Usage

### Installation

Reference the `DocConverter.FileNameService` project in your application.

### Configuration

#### Switching Between Database and Excel

Add configuration to your `appsettings.json`:

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=localhost;Database=YourDatabase;User Id=YourUser;Password=YourPassword;TrustServerCertificate=True;"
  },
  "Converter": {
    "FileNameConverterType": "Database",  // or "Excel"
    "ExcelFilePath": "Departments.xlsx"   // Optional: custom path for Excel file
  }
}
```

**Configuration Options:**
- `FileNameConverterType`: Set to `"Database"` or `"Excel"` to choose the service type
- `ExcelFilePath`: (Optional) Path to Excel file. If not specified, defaults to `Departments.xlsx` next to the executable

#### Database Mode Configuration

When using `FileNameConverterType: "Database"`, ensure you have a valid connection string configured.

**Database Schema:**

```sql
CREATE TABLE Departments (
    DepartmentSectionId INT PRIMARY KEY,
    DepartmentName NVARCHAR(255) NOT NULL,
    DepartmentSection NVARCHAR(255) NOT NULL
);
```

#### Excel Mode Configuration

When using `FileNameConverterType: "Excel"`, place your Excel file at the configured location.

**Excel File Structure:**

| Column A (1) | Column B (2) | Column C (3) | Column D (4) |
|--------------|--------------|--------------|--------------|
| DepartmentSectionNumber | DepartmentNumber | DepartmentAbbreviation | SectionAbbreviation |
| 123 | 10 | ENG | MECH |
| 456 | 20 | HR | REC |

### Code Examples

**Database Service:**
```csharp
using DocConverter.FileNameService;
using Microsoft.Extensions.Configuration;

var service = new FileNameConverterService(configuration);
var result = service.GetConvertedFileName("123-456.doc");
// Returns: "Engineering_Mechanical_123_456.pdf" (if dept 123 exists in DB)
```

**Excel Service:**
```csharp
using DocConverter.FileNameService;

// Using default path (Departments.xlsx next to executable)
var service = new ExcelFileNameConverterService();
var result = service.GetConvertedFileName("123-456.doc");
// Returns: "ENG_MECH_123_456.pdf" (if dept 123 exists in Excel)

// Using custom Excel file path
var service = new ExcelFileNameConverterService("C:\\Data\\Departments.xlsx");
var result = service.GetConvertedFileName("123-456.doc");
```

**Dynamic Selection (Based on Configuration):**
```csharp
var fileNameConverterType = configuration.GetValue<string>("Converter:FileNameConverterType", "Database");
IFileNameConverterService service;

if (fileNameConverterType.Equals("Excel", StringComparison.OrdinalIgnoreCase))
{
    var excelFilePath = configuration.GetValue<string>("Converter:ExcelFilePath");
    service = new ExcelFileNameConverterService(excelFilePath);
}
else
{
    service = new FileNameConverterService(configuration);
}
```

## Error Handling

Both services handle errors gracefully:
- If the data source (database/Excel) is unavailable, they return the original filename
- If the filename format is invalid, they return the original filename
- If the department section ID is not found, they return the original filename
- If any exception occurs, they return the original filename

## Integration

This library is used by both converter projects:
- `DocToPdfConverterLibreOffice` - Uses LibreOffice CLI for conversion
- `DocToPdfConverterNetOffice` - Uses NetOffice COM automation for conversion

Both applications support switching between Database and Excel modes via configuration.

