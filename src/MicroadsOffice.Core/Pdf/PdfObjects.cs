using System.Globalization;
using System.Text;

namespace MicroadsOffice.Core.Pdf;

public abstract class PdfObject
{
    public abstract void WriteTo(PdfByteWriter w);

    public byte[] ToBytes()
    {
        var w = new PdfByteWriter();
        WriteTo(w);
        return w.ToArray();
    }

    public override string ToString() => Encoding.Latin1.GetString(ToBytes());
}

public sealed class PdfNull : PdfObject
{
    public static readonly PdfNull Instance = new();
    public override void WriteTo(PdfByteWriter w) => w.Write("null");
}

public sealed class PdfBool(bool value) : PdfObject
{
    public bool Value { get; } = value;
    public override void WriteTo(PdfByteWriter w) => w.Write(Value ? "true" : "false");
}

public sealed class PdfNumber : PdfObject
{
    public double Value { get; }
    public bool IsInteger { get; }

    public PdfNumber(long v) { Value = v; IsInteger = true; }
    public PdfNumber(double v) { Value = v; IsInteger = Math.Abs(v - Math.Round(v)) < 1e-9 && Math.Abs(v) < 1e15; }

    public int IntValue => (int)Math.Round(Value);
    public long LongValue => (long)Math.Round(Value);

    public override void WriteTo(PdfByteWriter w) => w.Write(Format(Value));

    public static string Format(double v)
    {
        if (Math.Abs(v - Math.Round(v)) < 1e-9) return ((long)Math.Round(v)).ToString(CultureInfo.InvariantCulture);
        var s = v.ToString("0.#####", CultureInfo.InvariantCulture);
        return s == "-0" ? "0" : s;
    }
}

public sealed class PdfName : PdfObject
{
    public string Value { get; }
    public PdfName(string value) => Value = value;

    public override void WriteTo(PdfByteWriter w)
    {
        var sb = new StringBuilder("/");
        // Names are stored Latin-1 decoded (1 char == 1 byte) so that round-tripping preserves raw bytes.
        foreach (var b in Encoding.Latin1.GetBytes(Value))
        {
            if (b < 0x21 || b > 0x7E || "#/%()<>[]{}".IndexOf((char)b) >= 0)
                sb.Append('#').Append(b.ToString("X2"));
            else sb.Append((char)b);
        }
        w.Write(sb.ToString());
    }

    public override bool Equals(object? obj) => obj is PdfName n && n.Value == Value;
    public override int GetHashCode() => Value.GetHashCode();
}

public sealed class PdfString : PdfObject
{
    public byte[] Bytes { get; }
    public bool IsHex { get; }

    public PdfString(byte[] bytes, bool isHex = false) { Bytes = bytes; IsHex = isHex; }

    /// <summary>Creates a PDF text string (PDFDocEncoding when ASCII, otherwise UTF-16BE with BOM).</summary>
    public static PdfString FromText(string text)
    {
        bool ascii = text.All(c => c >= 0x20 && c < 0x7F);
        if (ascii) return new PdfString(Encoding.ASCII.GetBytes(text));
        var be = Encoding.BigEndianUnicode.GetBytes(text);
        var bytes = new byte[be.Length + 2];
        bytes[0] = 0xFE; bytes[1] = 0xFF;
        Buffer.BlockCopy(be, 0, bytes, 2, be.Length);
        return new PdfString(bytes, true);
    }

    public string ToText()
    {
        var b = Bytes;
        if (b.Length >= 2 && b[0] == 0xFE && b[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(b, 2, b.Length - 2);
        if (b.Length >= 2 && b[0] == 0xFF && b[1] == 0xFE) return Encoding.Unicode.GetString(b, 2, b.Length - 2);
        if (b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) return Encoding.UTF8.GetString(b, 3, b.Length - 3);
        return PdfDocEncoding.Decode(b);
    }

    public override void WriteTo(PdfByteWriter w)
    {
        if (IsHex)
        {
            w.Write("<");
            w.Write(Convert.ToHexString(Bytes));
            w.Write(">");
            return;
        }
        var sb = new StringBuilder("(");
        foreach (var b in Bytes)
        {
            switch (b)
            {
                case (byte)'(': sb.Append("\\("); break;
                case (byte)')': sb.Append("\\)"); break;
                case (byte)'\\': sb.Append("\\\\"); break;
                case (byte)'\r': sb.Append("\\r"); break;
                case (byte)'\n': sb.Append("\\n"); break;
                default:
                    if (b < 0x20 || b > 0x7E) sb.Append('\\').Append(Convert.ToString(b, 8).PadLeft(3, '0'));
                    else sb.Append((char)b);
                    break;
            }
        }
        sb.Append(')');
        w.Write(sb.ToString());
    }
}

public sealed class PdfArray : PdfObject
{
    public List<PdfObject> Items { get; } = new();
    public PdfArray() { }
    public PdfArray(IEnumerable<PdfObject> items) => Items.AddRange(items);
    public int Count => Items.Count;
    public PdfObject this[int i] => Items[i];
    public void Add(PdfObject o) => Items.Add(o);

    public static PdfArray Of(params double[] nums) => new(nums.Select(n => (PdfObject)new PdfNumber(n)));

    public override void WriteTo(PdfByteWriter w)
    {
        w.Write("[");
        for (int i = 0; i < Items.Count; i++)
        {
            if (i > 0) w.Write(" ");
            Items[i].WriteTo(w);
        }
        w.Write("]");
    }
}

public class PdfDictionary : PdfObject
{
    private readonly List<KeyValuePair<string, PdfObject>> _entries = new();

    public IEnumerable<KeyValuePair<string, PdfObject>> Entries => _entries;
    public IEnumerable<string> Keys => _entries.Select(e => e.Key);
    public int Count => _entries.Count;

    public PdfObject? this[string key]
    {
        get
        {
            foreach (var e in _entries) if (e.Key == key) return e.Value;
            return null;
        }
        set
        {
            int idx = _entries.FindIndex(e => e.Key == key);
            if (value == null) { if (idx >= 0) _entries.RemoveAt(idx); return; }
            if (idx >= 0) _entries[idx] = new(key, value);
            else _entries.Add(new(key, value));
        }
    }

    public bool ContainsKey(string key) => _entries.Any(e => e.Key == key);
    public PdfDictionary Set(string key, PdfObject value) { this[key] = value; return this; }
    public PdfDictionary Set(string key, string name) { this[key] = new PdfName(name); return this; }
    public PdfDictionary Set(string key, double number) { this[key] = new PdfNumber(number); return this; }

    public string? GetName(string key) => this[key] is PdfName n ? n.Value : null;

    /// <summary>Shallow copy (values are shared).</summary>
    public PdfDictionary CloneShallow()
    {
        var d = new PdfDictionary();
        foreach (var e in _entries) d._entries.Add(e);
        return d;
    }

    public override void WriteTo(PdfByteWriter w)
    {
        w.Write("<<");
        foreach (var e in _entries)
        {
            new PdfName(e.Key).WriteTo(w);
            w.Write(" ");
            e.Value.WriteTo(w);
            w.Write("\n");
        }
        w.Write(">>");
    }
}

public sealed class PdfStream : PdfObject
{
    public PdfDictionary Dict { get; }
    /// <summary>Encoded (raw) stream bytes as stored in the file.</summary>
    public byte[] Data { get; set; }

    public PdfStream(PdfDictionary dict, byte[] data) { Dict = dict; Data = data; }

    public override void WriteTo(PdfByteWriter w)
    {
        Dict["Length"] = new PdfNumber(Data.Length);
        Dict.WriteTo(w);
        w.Write("\nstream\n");
        w.Write(Data);
        w.Write("\nendstream");
    }
}

public sealed class PdfReference(int number, int generation) : PdfObject
{
    public int Number { get; } = number;
    public int Generation { get; } = generation;
    public override void WriteTo(PdfByteWriter w) => w.Write($"{Number} {Generation} R");
    public override bool Equals(object? obj) => obj is PdfReference r && r.Number == Number && r.Generation == Generation;
    public override int GetHashCode() => HashCode.Combine(Number, Generation);
}

/// <summary>Growable byte buffer used to serialise PDF syntax.</summary>
public sealed class PdfByteWriter
{
    private readonly MemoryStream _ms = new();
    public long Position => _ms.Length;
    public void Write(string s) { var b = Encoding.Latin1.GetBytes(s); _ms.Write(b, 0, b.Length); }
    public void Write(byte[] b) => _ms.Write(b, 0, b.Length);
    public void Write(ReadOnlySpan<byte> b) => _ms.Write(b);
    public byte[] ToArray() => _ms.ToArray();
}

internal static class PdfDocEncoding
{
    private static readonly char[] High = BuildHigh();

    private static char[] BuildHigh()
    {
        var t = new char[256];
        for (int i = 0; i < 256; i++) t[i] = (char)i;
        string s80 = "•†‡…—–ƒ⁄‹›−‰„“”‘’‚™ﬁﬂŁŒŠŸŽıłœšž�";
        for (int i = 0; i < s80.Length; i++) t[0x80 + i] = s80[i];
        t[0xA0] = '€';
        return t;
    }

    public static string Decode(byte[] b)
    {
        var c = new char[b.Length];
        for (int i = 0; i < b.Length; i++) c[i] = High[b[i]];
        return new string(c);
    }
}
