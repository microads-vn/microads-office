using MicroadsOffice.Core.Office;
using System.Globalization;
using System.Text;

namespace MicroadsOffice.Core.Pdf;

public sealed class PdfParseException(string message) : Exception(message);

/// <summary>Tokeniser / object parser for PDF syntax operating on an in-memory byte buffer.</summary>
internal sealed class PdfParser
{
    private readonly byte[] _d;
    private readonly Func<int, PdfObject?>? _resolver;
    public int Pos;

    public PdfParser(byte[] data, int pos, Func<int, PdfObject?>? resolver = null)
    {
        _d = data; Pos = pos; _resolver = resolver;
    }

    public static bool IsWhite(byte b) => b is 0 or 9 or 10 or 12 or 13 or 32;
    public static bool IsDelim(byte b) => b is (byte)'(' or (byte)')' or (byte)'<' or (byte)'>' or (byte)'[' or (byte)']' or (byte)'{' or (byte)'}' or (byte)'/' or (byte)'%';
    private bool Eof => Pos >= _d.Length;

    public void SkipWs()
    {
        while (Pos < _d.Length)
        {
            byte b = _d[Pos];
            if (IsWhite(b)) { Pos++; continue; }
            if (b == '%')
            {
                while (Pos < _d.Length && _d[Pos] != '\n' && _d[Pos] != '\r') Pos++;
                continue;
            }
            break;
        }
    }

    public string ReadKeyword()
    {
        SkipWs();
        int s = Pos;
        while (Pos < _d.Length && !IsWhite(_d[Pos]) && !IsDelim(_d[Pos])) Pos++;
        return Encoding.Latin1.GetString(_d, s, Pos - s);
    }

    public bool TryReadInt(out long value)
    {
        value = 0;
        SkipWs();
        int s = Pos;
        if (Pos < _d.Length && (_d[Pos] == '+' || _d[Pos] == '-')) Pos++;
        int digits = Pos;
        while (Pos < _d.Length && _d[Pos] >= '0' && _d[Pos] <= '9') Pos++;
        if (Pos == digits || (Pos < _d.Length && _d[Pos] == '.')) { Pos = s; return false; }
        if (Pos < _d.Length && !IsWhite(_d[Pos]) && !IsDelim(_d[Pos])) { Pos = s; return false; }
        if (!long.TryParse(Encoding.Latin1.GetString(_d, s, Pos - s), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value))
        { Pos = s; return false; }
        return true;
    }

    public PdfObject ParseObject(int depth = 0)
    {
        if (depth > 200) throw new PdfParseException(Loc.T("Cấu trúc lồng nhau quá sâu", "Structure nested too deeply"));
        SkipWs();
        if (Eof) throw new PdfParseException(Loc.T("Kết thúc tệp bất ngờ", "Unexpected end of file"));
        byte b = _d[Pos];
        switch (b)
        {
            case (byte)'/': return ParseName();
            case (byte)'(': return ParseLiteralString();
            case (byte)'[':
            {
                Pos++;
                var arr = new PdfArray();
                while (true)
                {
                    SkipWs();
                    if (Eof) throw new PdfParseException(Loc.T("Mảng chưa đóng", "Unclosed array"));
                    if (_d[Pos] == ']') { Pos++; break; }
                    arr.Add(ParseObject(depth + 1));
                }
                return arr;
            }
            case (byte)'<':
                if (Pos + 1 < _d.Length && _d[Pos + 1] == '<') return ParseDictionary(depth);
                return ParseHexString();
        }
        if (b is (byte)'+' or (byte)'-' or (byte)'.' || (b >= '0' && b <= '9'))
        {
            int save = Pos;
            if (TryReadInt(out long n1))
            {
                int afterFirst = Pos;
                if (TryReadInt(out long n2) && n2 >= 0)
                {
                    SkipWs();
                    if (Pos < _d.Length && _d[Pos] == 'R' && (Pos + 1 >= _d.Length || IsWhite(_d[Pos + 1]) || IsDelim(_d[Pos + 1])))
                    {
                        Pos++;
                        return new PdfReference((int)n1, (int)n2);
                    }
                }
                Pos = afterFirst;
                return new PdfNumber(n1);
            }
            Pos = save;
            return ParseReal();
        }
        var kw = ReadKeyword();
        return kw switch
        {
            "true" => new PdfBool(true),
            "false" => new PdfBool(false),
            "null" => PdfNull.Instance,
            "" => throw new PdfParseException(Loc.T($"Ký tự không hợp lệ tại {Pos}", $"Invalid character at {Pos}")),
            _ => throw new PdfParseException(Loc.T($"Từ khóa không mong đợi '{kw}' tại {Pos}", $"Unexpected keyword '{kw}' at {Pos}"))
        };
    }

    private PdfNumber ParseReal()
    {
        int s = Pos;
        while (Pos < _d.Length && (_d[Pos] is (byte)'+' or (byte)'-' or (byte)'.' || (_d[Pos] >= '0' && _d[Pos] <= '9'))) Pos++;
        var str = Encoding.Latin1.GetString(_d, s, Pos - s);
        // tolerate malformed numbers like "--5" or "5.-3"
        var cleaned = new StringBuilder();
        bool dot = false;
        for (int i = 0; i < str.Length; i++)
        {
            char c = str[i];
            if (c == '-' && cleaned.Length == 0) cleaned.Append(c);
            else if (c == '.' && !dot) { dot = true; cleaned.Append(c); }
            else if (char.IsDigit(c)) cleaned.Append(c);
        }
        double.TryParse(cleaned.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out double v);
        return new PdfNumber(v);
    }

    public PdfName ParseName()
    {
        Pos++; // '/'
        var bytes = new List<byte>();
        while (Pos < _d.Length && !IsWhite(_d[Pos]) && !IsDelim(_d[Pos]))
        {
            byte c = _d[Pos++];
            if (c == '#' && Pos + 1 < _d.Length && IsHex(_d[Pos]) && IsHex(_d[Pos + 1]))
            {
                bytes.Add((byte)(HexVal(_d[Pos]) * 16 + HexVal(_d[Pos + 1])));
                Pos += 2;
            }
            else bytes.Add(c);
        }
        return new PdfName(Encoding.Latin1.GetString(bytes.ToArray()));
    }

    private static bool IsHex(byte b) => (b >= '0' && b <= '9') || (b >= 'a' && b <= 'f') || (b >= 'A' && b <= 'F');
    private static int HexVal(byte b) => b <= '9' ? b - '0' : (b | 0x20) - 'a' + 10;

    private PdfString ParseHexString()
    {
        Pos++; // '<'
        var bytes = new List<byte>();
        int hi = -1;
        while (Pos < _d.Length && _d[Pos] != '>')
        {
            byte c = _d[Pos++];
            if (!IsHex(c)) continue;
            if (hi < 0) hi = HexVal(c);
            else { bytes.Add((byte)(hi * 16 + HexVal(c))); hi = -1; }
        }
        if (hi >= 0) bytes.Add((byte)(hi * 16));
        Pos++; // '>'
        return new PdfString(bytes.ToArray(), true);
    }

    private PdfString ParseLiteralString()
    {
        Pos++; // '('
        var bytes = new List<byte>();
        int depth = 1;
        while (Pos < _d.Length)
        {
            byte c = _d[Pos++];
            if (c == '\\')
            {
                if (Pos >= _d.Length) break;
                byte e = _d[Pos++];
                switch (e)
                {
                    case (byte)'n': bytes.Add(10); break;
                    case (byte)'r': bytes.Add(13); break;
                    case (byte)'t': bytes.Add(9); break;
                    case (byte)'b': bytes.Add(8); break;
                    case (byte)'f': bytes.Add(12); break;
                    case (byte)'\r': if (Pos < _d.Length && _d[Pos] == '\n') Pos++; break;
                    case (byte)'\n': break;
                    default:
                        if (e >= '0' && e <= '7')
                        {
                            int v = e - '0';
                            for (int k = 0; k < 2 && Pos < _d.Length && _d[Pos] >= '0' && _d[Pos] <= '7'; k++)
                                v = v * 8 + (_d[Pos++] - '0');
                            bytes.Add((byte)v);
                        }
                        else bytes.Add(e);
                        break;
                }
                continue;
            }
            if (c == '(') depth++;
            else if (c == ')') { if (--depth == 0) break; }
            else if (c == '\r')
            {
                if (Pos < _d.Length && _d[Pos] == '\n') Pos++;
                bytes.Add(10);
                continue;
            }
            bytes.Add(c);
        }
        return new PdfString(bytes.ToArray());
    }

    private PdfObject ParseDictionary(int depth)
    {
        Pos += 2;
        var dict = new PdfDictionary();
        while (true)
        {
            SkipWs();
            if (Eof) throw new PdfParseException(Loc.T("Từ điển chưa đóng", "Unclosed dictionary"));
            if (_d[Pos] == '>' && Pos + 1 < _d.Length && _d[Pos + 1] == '>') { Pos += 2; break; }
            if (_d[Pos] != '/')
            {
                // Garbage inside dictionary: skip one token to be tolerant.
                ParseObject(depth + 1);
                continue;
            }
            var key = ParseName();
            SkipWs();
            if (!Eof && _d[Pos] == '>' && Pos + 1 < _d.Length && _d[Pos + 1] == '>') { dict[key.Value] = PdfNull.Instance; continue; }
            var val = ParseObject(depth + 1);
            dict[key.Value] = val;
        }
        return dict;
    }

    /// <summary>Parses "num gen obj ... endobj" at the current position (handles streams).</summary>
    public (int Num, int Gen, PdfObject Obj) ParseIndirectObject()
    {
        if (!TryReadInt(out long num) || !TryReadInt(out long gen)) throw new PdfParseException(Loc.T($"Không tìm thấy đối tượng tại {Pos}", $"No object found at {Pos}"));
        var kw = ReadKeyword();
        if (kw != "obj") throw new PdfParseException(Loc.T($"Thiếu 'obj' tại {Pos}", $"Missing 'obj' at {Pos}"));
        var obj = ParseObject();
        if (obj is PdfDictionary dict)
        {
            int save = Pos;
            SkipWs();
            if (Pos + 6 <= _d.Length && Encoding.Latin1.GetString(_d, Pos, 6) == "stream")
            {
                Pos += 6;
                if (Pos < _d.Length && _d[Pos] == '\r') Pos++;
                if (Pos < _d.Length && _d[Pos] == '\n') Pos++;
                int start = Pos;
                int length = -1;
                var lenObj = dict["Length"];
                if (lenObj is PdfReference lr && _resolver != null) lenObj = _resolver(lr.Number);
                if (lenObj is PdfNumber ln) length = ln.IntValue;
                bool ok = length >= 0 && start + length <= _d.Length && LooksLikeEndStream(start + length);
                if (!ok)
                {
                    int idx = IndexOf(_d, "endstream"u8, start);
                    if (idx < 0) throw new PdfParseException(Loc.T("Thiếu endstream", "Missing endstream"));
                    int end = idx;
                    if (end > start && _d[end - 1] == '\n') end--;
                    if (end > start && _d[end - 1] == '\r') end--;
                    length = end - start;
                }
                var data = new byte[length];
                Buffer.BlockCopy(_d, start, data, 0, length);
                Pos = start + length;
                int es = IndexOf(_d, "endstream"u8, Pos);
                if (es >= 0) Pos = es + 9;
                return ((int)num, (int)gen, new PdfStream(dict, data));
            }
            Pos = save;
        }
        return ((int)num, (int)gen, obj);
    }

    private bool LooksLikeEndStream(int p)
    {
        while (p < _d.Length && IsWhite(_d[p])) p++;
        return p + 9 <= _d.Length && _d.AsSpan(p, 9).SequenceEqual("endstream"u8);
    }

    public static int IndexOf(byte[] data, ReadOnlySpan<byte> pattern, int start)
    {
        if (start >= data.Length) return -1;
        int i = data.AsSpan(start).IndexOf(pattern);
        return i < 0 ? -1 : start + i;
    }

    public static int LastIndexOf(byte[] data, ReadOnlySpan<byte> pattern, int searchFrom, int searchLen)
    {
        int s = Math.Max(0, searchFrom);
        int len = Math.Min(searchLen, data.Length - s);
        int i = data.AsSpan(s, len).LastIndexOf(pattern);
        return i < 0 ? -1 : s + i;
    }
}
