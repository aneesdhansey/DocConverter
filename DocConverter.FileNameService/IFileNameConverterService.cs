namespace DocConverter.FileNameService;

public interface IFileNameConverterService
{
    string GetConvertedFileName(string sourceFileName);
}
