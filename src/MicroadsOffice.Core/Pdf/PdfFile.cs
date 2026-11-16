using MicroadsOffice.Core.Office;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace MicroadsOffice.Core.Pdf;

/// <summary>
/// Random-access reader for the low-level object structure of a PDF file
/// (cross-reference tables, cross-reference streams, object streams, broken-file reconstruction).
/// </summary>
public sealed class PdfFile
{
    public readonly record struct XrefEntry(int Type, long Offset, int Gen, int StreamNum, int IndexInStream);

    private readonly Dictionary<int, XrefEntry> _xref = new();
    private readonly Dictionary<int, PdfObject?> _cache = new();
    private readonly Dictionary<int, (byte[] Data, int First, Dictionary<int, int> Offsets)> _objStreams = new();
    private List<PdfReference>? _pages;

    public byte[] Data { get; }
    public PdfDictionary Trailer { get; private set; } = new();
    public long LastXrefOffset { get; private set; }
    public bool LastXrefIsStream { get; private set; }
    /// <summary>True when the cross-reference information had to be rebuilt by scanning the file.</summary>
    public bool Reconstructed { get; private set; }
    public bool IsEncrypted => Trailer["Encrypt"] != null;
    public IReadOnlyDictionary<int, XrefEntry> Xref => _xref;

    public int Size
    {
        get
        {
            int size = Trailer["Size"] is PdfNumber n ? n.IntValue : 0;
            int max = _xref.Count == 0 ? 0 : _xref.Keys.Max() + 1;
            return Math.Max(size, max);
        }
    }

    private PdfFile(byte[] data) => Data = data;

    public static PdfFile Load(byte[] data)
    {
        var f = new PdfFile(data);
        if (PdfParser.IndexOf(data, "%PDF-"u8, 0) is < 0 or > 1024)
            throw new PdfParseException(Loc.T("Không phải tệp PDF", "Not a PDF file"));
        try
        {
            f.LoadXrefChain();
            if (!f.ValidateStructure()) throw new PdfParseException(Loc.T("Bảng xref không khớp", "The xref table does not match"));
        }
        catch (Exception)
        {
            f.Reconstruct();
            if (!f.ValidateStructure()) throw new PdfParseException(Loc.T("Không thể đọc cấu trúc tài liệu PDF.", "Could not read the PDF document structure."));
        }
        return f;
    }

    private bool ValidateStructure()
    {
        try
        {
            if (Trailer["Root"] is not PdfReference) return false;
            if (Catalog is null) return false;
            return GetPageRefs().Count > 0;
        }
        catch { return false; }
    }

    // ------------------------------------------------------------------ xref loading

    private void LoadXrefChain()
    {
        int sx = PdfParser.LastIndexOf(Data, "startxref"u8, Data.Length - 4096, 4096);
        if (sx < 0) throw new PdfParseException(Loc.T("Thiếu startxref", "Missing startxref"));
        var p = new PdfParser(Data, sx + 9);
        if (!p.TryReadInt(out long offset)) throw new PdfParseException(Loc.T("startxref không hợp lệ", "Invalid startxref"));
        LastXrefOffset = offset;
        var visited = new HashSet<long>();
        bool first = true;
        long? next = offset;
        while (next is long off && visited.Add(off))
        {
            if (off < 0 || off >= Data.Length) throw new PdfParseException(Loc.T("Offset xref ngoài phạm vi", "xref offset out of range"));
            var (trailer, isStream) = ReadXrefSection(off);
            if (first)
            {
                Trailer = trailer;
                LastXrefIsStream = isStream;
                first = false;
            }
            next = trailer["Prev"] is PdfNumber prev ? prev.LongValue : null;
        }
    }

    private (PdfDictionary Trailer, bool IsStream) ReadXrefSection(long off)
    {
        var p = new PdfParser(Data, (int)off);
        p.SkipWs();
        int save = p.Pos;
        if (p.ReadKeyword() == "xref")
        {
            var section = new Dictionary<int, XrefEntry>();
            while (true)
            {
                int before = p.Pos;
                if (!p.TryReadInt(out long start)) { p.Pos = before; break; }
                if (!p.TryReadInt(out long count)) throw new PdfParseException(Loc.T("Bảng xref hỏng", "Corrupted xref table"));
                for (long i = 0; i < count; i++)
                {
                    if (!p.TryReadInt(out long o) || !p.TryReadInt(out long g)) throw new PdfParseException(Loc.T("Mục xref hỏng", "Corrupted xref entry"));
                    var t = p.ReadKeyword();
                    int num = (int)(start + i);
                    section[num] = new XrefEntry(t == "n" ? 1 : 0, o, (int)g, 0, 0);
                }
            }
            if (p.ReadKeyword() != "trailer") throw new PdfParseException(Loc.T("Thiếu trailer", "Missing trailer"));
            var trailer = p.ParseObject() as PdfDictionary ?? throw new PdfParseException(Loc.T("Trailer không hợp lệ", "Invalid trailer"));
            if (trailer["XRefStm"] is PdfNumber xs)
            {
                try
                {
                    var (_, _, sObj) = new PdfParser(Data, (int)xs.LongValue, ResolveForParser).ParseIndirectObject();
                    if (sObj is PdfStream st)
                        foreach (var (num, e) in DecodeXrefStream(st))
                            if (!section.TryGetValue(num, out var ex) || ex.Type == 0) section[num] = e;
                }
                catch { /* hybrid stream unreadable: fall back to table only */ }
            }
            foreach (var (num, e) in section) _xref.TryAdd(num, e);
            return (trailer, false);
        }
        p.Pos = save;
        var (_, _, obj) = p.ParseIndirectObject();
        if (obj is not PdfStream xstream || xstream.Dict.GetName("Type") != "XRef")
            throw new PdfParseException(Loc.T("Không tìm thấy bảng xref", "xref table not found"));
        foreach (var (num, e) in DecodeXrefStream(xstream)) _xref.TryAdd(num, e);
        return (xstream.Dict, true);
    }

    private PdfObject? ResolveForParser(int num)
    {
        // Only used for /Length lookups while parsing; avoid infinite recursion via the cache marker.
        if (_resolving.Contains(num)) return null;
        return GetObject(num);
    }

    private readonly HashSet<int> _resolving = new();

    private IEnumerable<(int, XrefEntry)> DecodeXrefStream(PdfStream st)
    {
        var data = DecodeStream(st);
        var w = (st.Dict["W"] as PdfArray)?.Items.Select(i => (i as PdfNumber)?.IntValue ?? 0).ToArray()
                ?? throw new PdfParseException(Loc.T("Xref stream thiếu /W", "xref stream is missing /W"));
        int size = st.Dict["Size"] is PdfNumber sz ? sz.IntValue : 0;
        var index = st.Dict["Index"] is PdfArray ia ? ia.Items.Select(i => (i as PdfNumber)?.IntValue ?? 0).ToArray() : [0, size];
        int rowLen = w.Sum();
        int pos = 0;
        var result = new List<(int, XrefEntry)>();
        for (int k = 0; k + 1 < index.Length; k += 2)
        {
            int start = index[k], count = index[k + 1];
            for (int i = 0; i < count && pos + rowLen <= data.Length; i++)
            {
                long f1 = w[0] == 0 ? 1 : ReadBE(data, pos, w[0]);
                long f2 = ReadBE(data, pos + w[0], w[1]);
                long f3 = ReadBE(data, pos + w[0] + w[1], w[2]);
                pos += rowLen;
                int num = start + i;
                result.Add(f1 switch
                {
                    1 => (num, new XrefEntry(1, f2, (int)f3, 0, 0)),
                    2 => (num, new XrefEntry(2, 0, 0, (int)f2, (int)f3)),
                    _ => (num, new XrefEntry(0, 0, (int)f3, 0, 0)),
                });
            }
        }
        return result;
    }

    private static long ReadBE(byte[] d, int pos, int len)
    {
        long v = 0;
        for (int i = 0; i < len; i++) v = (v << 8) | d[pos + i];
        return v;
    }

    // ------------------------------------------------------------------ reconstruction

    private void Reconstruct()
    {
        Reconstructed = true;
        _xref.Clear();
        _cache.Clear();
        _objStreams.Clear();
        _pages = null;
        var text = Encoding.Latin1.GetString(Data);
        var objRegex = new Regex(@"(?<![0-9])(\d{1,10})\s+(\d{1,5})\s+obj\b", RegexOptions.Compiled);
        foreach (Match m in objRegex.Matches(text))
        {
            int num = int.Parse(m.Groups[1].Value);
            int gen = int.Parse(m.Groups[2].Value);
            _xref[num] = new XrefEntry(1, m.Index, gen, 0, 0); // later definitions win
        }
        // Index objects living inside object streams.
        foreach (var (num, e) in _xref.ToList())
        {
            if (e.Type != 1) continue;
            try
            {
                if (GetObject(num) is PdfStream s && s.Dict.GetName("Type") == "ObjStm")
                {
                    var os = LoadObjStm(num);
                    foreach (var (inner, idx) in os.Offsets.Select((kv, i) => (kv.Key, i)))
                        // keep whichever definition appears later in the file (newer incremental update wins)
                        if (!_xref.TryGetValue(inner, out var ex) || (ex.Type == 1 && ex.Offset < e.Offset))
                            _xref[inner] = new XrefEntry(2, 0, 0, num, idx);
                }
            }
            catch { }
        }
        _cache.Clear();
        // Find a trailer: classic trailer dictionaries first (last one wins), then xref streams, then catalog scan.
        PdfDictionary trailer = new();
        int ti = 0;
        while ((ti = text.IndexOf("trailer", ti, StringComparison.Ordinal)) >= 0)
        {
            try
            {
                var p = new PdfParser(Data, ti + 7);
                if (p.ParseObject() is PdfDictionary d && d["Root"] is PdfReference) trailer = d;
            }
            catch { }
            ti += 7;
        }
        if (trailer["Root"] is not PdfReference)
        {
            foreach (var num in _xref.Keys.OrderBy(k => k))
            {
                try
                {
                    var o = GetObject(num);
                    var dict = o is PdfStream s ? s.Dict : o as PdfDictionary;
                    if (dict?.GetName("Type") == "XRef" && dict["Root"] is PdfReference) trailer = dict;
                }
                catch { }
            }
        }
        if (trailer["Root"] is not PdfReference)
        {
            foreach (var num in _xref.Keys.OrderBy(k => k))
            {
                try
                {
                    if (GetObject(num) is PdfDictionary d && d.GetName("Type") == "Catalog")
                    {
                        trailer = new PdfDictionary();
                        trailer["Root"] = new PdfReference(num, _xref[num].Gen);
                    }
                }
                catch { }
            }
        }
        var t = new PdfDictionary();
        foreach (var key in new[] { "Root", "Info", "ID", "Encrypt" })
            if (trailer[key] != null) t[key] = trailer[key]!;
        t["Size"] = new PdfNumber(_xref.Count == 0 ? 1 : _xref.Keys.Max() + 1);
        Trailer = t;
        LastXrefOffset = -1;
        LastXrefIsStream = true;
    }

    // ------------------------------------------------------------------ object access

    public PdfObject? GetObject(int num)
    {
        if (_cache.TryGetValue(num, out var cached)) return cached;
        if (!_xref.TryGetValue(num, out var e) || e.Type == 0) return null;
        PdfObject? obj = null;
        _resolving.Add(num);
        try
        {
            if (e.Type == 1)
            {
                if (e.Offset < 0 || e.Offset >= Data.Length) throw new PdfParseException(Loc.T($"Offset đối tượng {num} sai", $"Wrong offset for object {num}"));
                var p = new PdfParser(Data, (int)e.Offset, ResolveForParser);
                var (n, _, o) = p.ParseIndirectObject();
                if (n != num) throw new PdfParseException(Loc.T($"Đối tượng {num} không khớp offset", $"Object {num} does not match its offset"));
                obj = o;
            }
            else
            {
                var os = LoadObjStm(e.StreamNum);
                if (os.Offsets.TryGetValue(num, out int rel))
                {
                    var p = new PdfParser(os.Data, os.First + rel, ResolveForParser);
                    obj = p.ParseObject();
                }
            }
        }
        finally { _resolving.Remove(num); }
        _cache[num] = obj;
        return obj;
    }

    private (byte[] Data, int First, Dictionary<int, int> Offsets) LoadObjStm(int streamNum)
    {
        if (_objStreams.TryGetValue(streamNum, out var cached)) return cached;
        if (GetObject(streamNum) is not PdfStream s) throw new PdfParseException(Loc.T("Object stream không hợp lệ", "Invalid object stream"));
        var data = DecodeStream(s);
        int n = s.Dict["N"] is PdfNumber nn ? nn.IntValue : 0;
        int first = s.Dict["First"] is PdfNumber ff ? ff.IntValue : 0;
        var p = new PdfParser(data, 0);
        var offsets = new Dictionary<int, int>();
        for (int i = 0; i < n; i++)
        {
            if (!p.TryReadInt(out long on) || !p.TryReadInt(out long oo)) break;
            offsets.TryAdd((int)on, (int)oo);
        }
        var res = (data, first, offsets);
        _objStreams[streamNum] = res;
        return res;
    }

    public int GetGeneration(int num) => _xref.TryGetValue(num, out var e) && e.Type == 1 ? e.Gen : 0;

    public PdfObject? Resolve(PdfObject? o)
    {
        int guard = 0;
        while (o is PdfReference r && guard++ < 32) o = GetObject(r.Number);
        return o;
    }

    public PdfDictionary? ResolveDict(PdfObject? o) => Resolve(o) switch
    {
        PdfDictionary d => d,
        PdfStream s => s.Dict,
        _ => null
    };

    public PdfDictionary? Catalog => ResolveDict(Trailer["Root"]);

    public IReadOnlyList<PdfReference> GetPageRefs()
    {
        if (_pages != null) return _pages;
        var list = new List<PdfReference>();
        var visited = new HashSet<int>();
        void Walk(PdfObject? node, int depth)
        {
            if (depth > 64 || node is not PdfReference r || !visited.Add(r.Number)) return;
            var d = ResolveDict(r);
            if (d == null) return;
            var type = d.GetName("Type");
            if (type == "Pages" || (type == null && d["Kids"] != null))
            {
                if (Resolve(d["Kids"]) is PdfArray kids)
                    foreach (var k in kids.Items) Walk(k, depth + 1);
            }
            else list.Add(r);
        }
        Walk(Catalog?["Pages"], 0);
        _pages = list;
        return list;
    }

    // ------------------------------------------------------------------ filters

    public static byte[] DecodeStream(PdfStream s)
    {
        var filter = s.Dict["Filter"];
        var parms = s.Dict["DecodeParms"];
        var filters = filter switch
        {
            PdfName n => new List<string> { n.Value },
            PdfArray a => a.Items.OfType<PdfName>().Select(x => x.Value).ToList(),
            _ => new List<string>()
        };
        var data = s.Data;
        for (int i = 0; i < filters.Count; i++)
        {
            var f = filters[i];
            var p = parms is PdfArray pa ? (i < pa.Count ? pa[i] as PdfDictionary : null) : parms as PdfDictionary;
            if (f is "FlateDecode" or "Fl")
            {
                data = Inflate(data);
                data = ApplyPredictor(data, p);
            }
            else throw new PdfParseException(Loc.T("Bộ lọc không hỗ trợ: ", "Unsupported filter: ") + f);
        }
        return data;
    }

    public static byte[] Inflate(byte[] input)
    {
        try
        {
            using var ms = new MemoryStream(input);
            using var z = new ZLibStream(ms, CompressionMode.Decompress);
            using var outMs = new MemoryStream();
            z.CopyTo(outMs);
            return outMs.ToArray();
        }
        catch
        {
            // Some producers omit/garble the zlib header: try raw deflate, keep partial output on error.
            using var outMs = new MemoryStream();
            try
            {
                using var ms = new MemoryStream(input, 2, Math.Max(0, input.Length - 2));
                using var z = new DeflateStream(ms, CompressionMode.Decompress);
                var buf = new byte[8192];
                int n;
                while ((n = z.Read(buf, 0, buf.Length)) > 0) outMs.Write(buf, 0, n);
            }
            catch { }
            return outMs.ToArray();
        }
    }

    private static byte[] ApplyPredictor(byte[] data, PdfDictionary? p)
    {
        if (p == null) return data;
        int predictor = p["Predictor"] is PdfNumber pr ? pr.IntValue : 1;
        if (predictor < 10) return data;
        int colors = p["Colors"] is PdfNumber c ? c.IntValue : 1;
        int bpc = p["BitsPerComponent"] is PdfNumber b ? b.IntValue : 8;
        int columns = p["Columns"] is PdfNumber col ? col.IntValue : 1;
        int bpp = Math.Max(1, colors * bpc / 8);
        int rowLen = (colors * bpc * columns + 7) / 8;
        var output = new MemoryStream();
        var prev = new byte[rowLen];
        var cur = new byte[rowLen];
        int pos = 0;
        while (pos + 1 + rowLen <= data.Length)
        {
            int type = data[pos++];
            Buffer.BlockCopy(data, pos, cur, 0, rowLen);
            pos += rowLen;
            for (int i = 0; i < rowLen; i++)
            {
                int left = i >= bpp ? cur[i - bpp] : 0;
                int up = prev[i];
                int ul = i >= bpp ? prev[i - bpp] : 0;
                cur[i] = type switch
                {
                    1 => (byte)(cur[i] + left),
                    2 => (byte)(cur[i] + up),
                    3 => (byte)(cur[i] + ((left + up) >> 1)),
                    4 => (byte)(cur[i] + Paeth(left, up, ul)),
                    _ => cur[i]
                };
            }
            output.Write(cur, 0, rowLen);
            (prev, cur) = (cur, prev);
        }
        return output.ToArray();
    }

    private static int Paeth(int a, int b, int c)
    {
        int p = a + b - c, pa = Math.Abs(p - a), pb = Math.Abs(p - b), pc = Math.Abs(p - c);
        return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
    }
}
