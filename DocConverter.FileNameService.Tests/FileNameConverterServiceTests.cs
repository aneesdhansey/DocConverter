using Microsoft.Extensions.Configuration;
using Moq;
using Xunit;

namespace DocConverter.FileNameService.Tests;

public class FileNameConverterServiceTests
{
    [Fact]
    public void GetConvertedFileName_WithValidFormat_ReturnsOriginalWhenNoDatabaseConnection()
    {
        // Arrange
        var mockConfiguration = new Mock<IConfiguration>();
        var mockConnectionStringsSection = new Mock<IConfigurationSection>();
        
        mockConnectionStringsSection.Setup(x => x.Value).Returns((string)null);
        mockConfiguration.Setup(x => x.GetSection("ConnectionStrings")).Returns(mockConnectionStringsSection.Object);
        
        var service = new FileNameConverterService(mockConfiguration.Object);

        // Act
        var result = service.GetConvertedFileName("123-456.doc");

        // Assert
        Assert.Equal("123-456.doc", result);
    }

    [Fact]
    public void GetConvertedFileName_WithInvalidFormat_ReturnsOriginalFileName()
    {
        // Arrange
        var mockConfiguration = new Mock<IConfiguration>();
        var mockConnectionStringsSection = new Mock<IConfigurationSection>();
        
        mockConnectionStringsSection.Setup(x => x.Value).Returns((string)null);
        mockConfiguration.Setup(x => x.GetSection("ConnectionStrings")).Returns(mockConnectionStringsSection.Object);
        
        var service = new FileNameConverterService(mockConfiguration.Object);

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
        var mockConfiguration = new Mock<IConfiguration>();
        var mockConnectionStringsSection = new Mock<IConfigurationSection>();
        
        mockConnectionStringsSection.Setup(x => x.Value).Returns((string)null);
        mockConfiguration.Setup(x => x.GetSection("ConnectionStrings")).Returns(mockConnectionStringsSection.Object);
        
        var service = new FileNameConverterService(mockConfiguration.Object);

        // Act - even with null, the service should handle it gracefully
        var result = service.GetConvertedFileName(null);

        // Assert - should not throw exception and return original
        Assert.NotNull(result);
    }

    [Fact]
    public void Constructor_WithNoConnectionString_DoesNotThrow()
    {
        // Arrange
        var mockConfiguration = new Mock<IConfiguration>();
        var mockConnectionStringsSection = new Mock<IConfigurationSection>();
        
        mockConnectionStringsSection.Setup(x => x.Value).Returns((string)null);
        mockConfiguration.Setup(x => x.GetSection("ConnectionStrings")).Returns(mockConnectionStringsSection.Object);
        
        // Act & Assert - should not throw
        var service = new FileNameConverterService(mockConfiguration.Object);
        Assert.NotNull(service);
    }

    [Fact]
    public void GetConvertedFileName_WithValidFormatButNoDepartmentData_ReturnsOriginalFileName()
    {
        // Arrange
        var mockConfiguration = new Mock<IConfiguration>();
        var mockConnectionStringsSection = new Mock<IConfigurationSection>();
        
        mockConnectionStringsSection.Setup(x => x.Value).Returns((string)null);
        mockConfiguration.Setup(x => x.GetSection("ConnectionStrings")).Returns(mockConnectionStringsSection.Object);
        
        var service = new FileNameConverterService(mockConfiguration.Object);

        // Act
        var result = service.GetConvertedFileName("999-888.doc");

        // Assert - Should return original since department 999 won't exist
        Assert.Equal("999-888.doc", result);
    }
}
