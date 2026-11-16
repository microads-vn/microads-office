using MicroadsOffice.Word.UI;

namespace MicroadsOffice.Word;

/// <summary>Entry point of Microads Office Word.</summary>
public static class WordApp
{
    /// <summary>Creates the main Word window (implements <see cref="MicroadsOffice.UI.IOfficeAppWindow"/>).</summary>
    public static Form CreateMainForm(string[] files) => new WordMainForm(files);
}
