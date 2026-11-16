namespace MicroadsOffice;

/// <summary>Entry point of Microads Office PDF (reader, digital signing, printing).</summary>
public static class PdfApp
{
    public static Form CreateMainForm(string[] files) => new MainForm(files);
}
