using MicroadsOffice.Excel.Ui;

namespace MicroadsOffice.Excel;

/// <summary>Entry point of Microads Office Excel.</summary>
public static class ExcelApp
{
    public static Form CreateMainForm(string[] files) => new ExcelForm(files);
}
