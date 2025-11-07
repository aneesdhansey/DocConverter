namespace DocConverter.FileNameService.Models;

public class ExcelDepartment
{
    public int DepartmentSectionNumber { get; set; }
    public int DepartmentNumber { get; set; }
    public string DepartmentAbbreviation { get; set; } = string.Empty;
    public string SectionAbbreviation { get; set; } = string.Empty;
}
