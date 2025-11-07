# DocConverter.FileNameService

A library for converting document file names based on department information from a database.

## Overview

This service extracts department section IDs from source file names and converts them to a more descriptive format using department information from a SQL Server database.

## Features

- Parses filenames in the format `XXX-YYYY` (with any extension) where:
  - `XXX` is the department section ID
  - `YYYY` is the document number
- Converts to: `{DepartmentName}_{DepartmentSection}_{XXX}_{YYYY}.pdf`
- Sanitizes department names to remove invalid filename characters
- Loads all department info at startup for performance
- Graceful error handling - returns original filename on any error
- Uses Dapper for efficient database access

## Usage

### Installation

Reference the `DocConverter.FileNameService` project in your application.

### Configuration

Add a connection string to your `appsettings.json`:

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=localhost;Database=YourDatabase;User Id=YourUser;Password=YourPassword;TrustServerCertificate=True;"
  }
}
```

### Database Schema

The service expects a `Departments` table with the following structure:

```sql
CREATE TABLE Departments (
    DepartmentSectionId INT PRIMARY KEY,
    DepartmentName NVARCHAR(255) NOT NULL,
    DepartmentSection NVARCHAR(255) NOT NULL
);
```

### Code Example

```csharp
using DocConverter.FileNameService;
using Microsoft.Extensions.Configuration;

// Initialize the service
var service = new FileNameConverterService(configuration);

// Convert a filename
var originalFileName = "123-456.doc";
var convertedFileName = service.GetConvertedFileName(originalFileName);
// Result: "Engineering_Mechanical_123_456.pdf" (assuming dept 123 exists in DB)
```

## Error Handling

The service handles errors gracefully:
- If the database is unavailable, it returns the original filename
- If the filename format is invalid, it returns the original filename
- If the department section ID is not found, it returns the original filename
- If any exception occurs, it returns the original filename

## Integration

This library is used by both converter projects:
- `DocToPdfConverterLibreOffice` - Uses LibreOffice CLI for conversion
- `DocToPdfConverterNetOffice` - Uses NetOffice COM automation for conversion
