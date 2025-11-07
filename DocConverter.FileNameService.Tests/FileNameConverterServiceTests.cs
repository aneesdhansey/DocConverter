using Microsoft.Extensions.Configuration;
using Xunit;

namespace DocConverter.FileNameService.Tests;

public class FileNameConverterServiceTests
{
    private static IConfiguration CreateConfiguration(string? connectionString = null)
    {
        var configBuilder = new ConfigurationBuilder();
        if (connectionString != null)
        {
            configBuilder.AddInMemoryCollection(new Dictionary<string, string?>
            {
                { "ConnectionStrings:DefaultConnection", connectionString }
            });
        }
        return configBuilder.Build();
    }

    [Fact]
    public void GetConvertedFileName_WithValidFormat_ReturnsOriginalWhenNoDatabaseConnection()
    {
        // Arrange
        var configuration = CreateConfiguration();
        var service = new FileNameConverterService(configuration);

        // Act
        var result = service.GetConvertedFileName("123-456.doc");

        // Assert
        Assert.Equal("123-456.doc", result);
    }

    [Fact]
    public void GetConvertedFileName_WithInvalidFormat_ReturnsOriginalFileName()
    {
        // Arrange
        var configuration = CreateConfiguration();
        var service = new FileNameConverterService(configuration);

        // Act & Assert
        Assert.Equal("invalid.doc", service.GetConvertedFileName("invalid.doc"));
        Assert.Equal("123.doc", service.GetConvertedFileName("123.doc"));
        Assert.Equal("abc-def.doc", service.GetConvertedFileName("abc-def.doc"));
        Assert.Equal("123-abc.doc", service.GetConvertedFileName("123-abc.doc"));
        Assert.Equal("abc-123.doc", service.GetConvertedFileName("abc-123.doc"));
    }

    [Fact]
    public void GetConvertedFileName_WithException_ReturnsOriginalFileName()
    {
        // Arrange
        var configuration = CreateConfiguration();
        var service = new FileNameConverterService(configuration);

        // Act - even with null, the service should handle it gracefully
        var result = service.GetConvertedFileName(null!);

        // Assert - should not throw exception and return empty string
        Assert.NotNull(result);
        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void Constructor_WithNoConnectionString_DoesNotThrow()
    {
        // Arrange & Act
        var configuration = CreateConfiguration();
        var service = new FileNameConverterService(configuration);
        
        // Assert - should not throw
        Assert.NotNull(service);
    }

    [Fact]
    public void GetConvertedFileName_WithValidFormatButNoDepartmentData_ReturnsOriginalFileName()
    {
        // Arrange
        var configuration = CreateConfiguration();
        var service = new FileNameConverterService(configuration);

        // Act
        var result = service.GetConvertedFileName("999-888.doc");

        // Assert - Should return original since department 999 won't exist
        Assert.Equal("999-888.doc", result);
    }
}
