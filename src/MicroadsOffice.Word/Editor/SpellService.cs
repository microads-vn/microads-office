using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.Editor;

/// <summary>
/// Spelling through the Windows Spell Checking API (Windows 8+). Uses Vietnamese when the
/// language pack is installed, otherwise English; words in another script are not flagged.
/// </summary>
public sealed class SpellService
{
    [ComImport, Guid("8E018A9D-2415-4677-BF08-794EA61F94BB"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface ISpellCheckerFactory
    {
        IEnumString SupportedLanguages { get; }
        int IsSupported([MarshalAs(UnmanagedType.LPWStr)] string languageTag);
        ISpellChecker CreateSpellChecker([MarshalAs(UnmanagedType.LPWStr)] string languageTag);
    }

    [ComImport, Guid("B6FD0B71-E2BC-4653-8D05-F197E412770B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface ISpellChecker
    {
        [return: MarshalAs(UnmanagedType.LPWStr)] string get_LanguageTag();
        IEnumSpellingError Check([MarshalAs(UnmanagedType.LPWStr)] string text);
        IEnumString Suggest([MarshalAs(UnmanagedType.LPWStr)] string word);
        void Add([MarshalAs(UnmanagedType.LPWStr)] string word);
        void Ignore([MarshalAs(UnmanagedType.LPWStr)] string word);
    }

    [ComImport, Guid("803E3BD4-2828-4410-8290-418D1D73C762"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IEnumSpellingError
    {
        [PreserveSig] int Next(out ISpellingError? value);
    }

    [ComImport, Guid("B7C82D61-FBE8-4B47-9B27-6C0D2E0DE0A3"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface ISpellingError
    {
        uint StartIndex { get; }
        uint Length { get; }
        int CorrectiveAction { get; }
        string Replacement { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
    }

    [ComImport, Guid("7AB36653-1796-484B-BDFA-E74F1DB7C1DC")]
    private class SpellCheckerFactoryClass { }

    private ISpellChecker? _checker;
    private readonly ConditionalWeakTable<Paragraph, Entry> _cache = new();
    private sealed class Entry { public int Version; public List<(int, int)> Errors = new(); }

    public bool Available => _checker != null;
    public string Language { get; private set; } = "";
    public bool Enabled { get; set; } = true;
    /// <summary>Set when the document is Vietnamese but only a non-Vietnamese dictionary exists: nothing is flagged.</summary>
    public bool Suppressed { get; set; }

    public SpellService()
    {
        try
        {
            var f = (ISpellCheckerFactory)new SpellCheckerFactoryClass();
            foreach (var tag in new[] { "vi-VN", "en-US", "en-GB" })
            {
                if (f.IsSupported(tag) != 0)
                {
                    _checker = f.CreateSpellChecker(tag);
                    Language = tag;
                    break;
                }
            }
        }
        catch { _checker = null; }
    }

    private bool IsVietnamese => Language.StartsWith("vi", StringComparison.OrdinalIgnoreCase);

    public IReadOnlyList<(int, int)> Errors(Paragraph p)
    {
        if (!Enabled || _checker == null || Suppressed) return [];
        if (_cache.TryGetValue(p, out var e) && e.Version == p.Version) return e.Errors;
        e = new Entry { Version = p.Version, Errors = CheckText(p.Text) };
        _cache.AddOrUpdate(p, e);
        return e.Errors;
    }

    public List<(int, int)> CheckText(string text)
    {
        var list = new List<(int, int)>();
        if (_checker == null || string.IsNullOrWhiteSpace(text)) return list;
        // Vietnamese paragraph but no Vietnamese dictionary installed: do not flag anything
        if (!IsVietnamese && LooksVietnamese(text)) return list;
        var clean = text.Replace(Run.ObjectChar, ' ').Replace(Run.LineBreak, ' ').Replace(Run.PageBreak, ' ').Replace('\t', ' ');
        try
        {
            var en = _checker.Check(clean);
            while (en.Next(out var err) == 0 && err != null)
            {
                int s = (int)err.StartIndex, l = (int)err.Length;
                if (s + l <= clean.Length)
                {
                    var word = clean.Substring(s, l);
                    if (!Skip(word)) list.Add((s, l));
                }
                Marshal.ReleaseComObject(err);
            }
        }
        catch { }
        return list;
    }

    private const string VietChars = "ăâđêôơưàảãáạằẳẵắặầẩẫấậèẻẽéẹềểễếệìỉĩíịòỏõóọồổỗốộờởỡớợùủũúụừửữứựỳỷỹýỵ";

    public static bool LooksVietnamese(string text)
    {
        int n = 0;
        foreach (var c in text) if (VietChars.Contains(char.ToLowerInvariant(c)) && ++n >= 2) return true;
        return false;
    }

    private bool Skip(string word)
    {
        if (word.Any(char.IsDigit) || word.Contains('@') || word.Contains("://")) return true;
        if (word.Length > 1 && word.All(c => !char.IsLetter(c) || char.IsUpper(c))) return true; // ACRONYMS
        if (!IsVietnamese && word.Any(c => c > 127)) return true;                                 // Vietnamese text with an English dictionary
        return false;
    }

    public List<string> Suggest(string word)
    {
        var r = new List<string>();
        if (_checker == null) return r;
        try
        {
            var e = _checker.Suggest(word);
            var arr = new string[1];
            while (r.Count < 8 && e.Next(1, arr, IntPtr.Zero) == 0) r.Add(arr[0]);
        }
        catch { }
        return r;
    }

    public void Ignore(string word) { try { _checker?.Ignore(word); } catch { } _cache.Clear(); }
    public void AddToDictionary(string word) { try { _checker?.Add(word); } catch { } _cache.Clear(); }
}
