using MicroadsOffice.Core.Office;
using System.Security.Cryptography;

namespace MicroadsOffice.Core.Pdf;

/// <summary>
/// Builds a PDF incremental update (ISO 32000-1 §7.5.6): new/changed objects are appended after the
/// original bytes, followed by a new cross-reference section that chains to the previous one.
/// The original bytes are never modified, so existing digital signatures stay valid.
/// </summary>
public sealed class IncrementalUpdate
{
    private readonly PdfFile _file;
    private readonly SortedDictionary<int, (int Gen, PdfObject? Obj)> _objects = new();
    private int _next;

    /// <summary>Offset in the final output where each written object starts (filled by Build).</summary>
    public Dictionary<int, long> WrittenOffsets { get; } = new();

    public IncrementalUpdate(PdfFile file)
    {
        _file = file;
        _next = file.Size;
    }

    public PdfFile File => _file;

    public PdfReference Reserve()
    {
        var r = new PdfReference(_next++, 0);
        _objects[r.Number] = (0, null);
        return r;
    }

    public PdfReference Add(PdfObject obj)
    {
        var r = Reserve();
        _objects[r.Number] = (0, obj);
        return r;
    }

    public void Set(PdfReference r, PdfObject obj) => _objects[r.Number] = (r.Generation, obj);

    /// <summary>Replace an existing object of the original file (same object number and generation).</summary>
    public void Replace(int num, PdfObject obj) => _objects[num] = (_file.GetGeneration(num), obj);

    /// <summary>Returns the pending (modified) version of an object if any, else the original.</summary>
    public PdfObject? Get(int num) => _objects.TryGetValue(num, out var v) && v.Obj != null ? v.Obj : _file.GetObject(num);

    public byte[] Build()
    {
        var orig = _file.Data;
        var w = new PdfByteWriter();
        w.Write(orig);
        byte last = orig.Length > 0 ? orig[^1] : (byte)'\n';
        if (last != '\n' && last != '\r') w.Write("\n");

        foreach (var (num, (gen, obj)) in _objects)
        {
            if (obj == null) throw new InvalidOperationException(Loc.T($"Đối tượng {num} chưa được gán", $"Object {num} has not been assigned"));
            WrittenOffsets[num] = w.Position;
            w.Write($"{num} {gen} obj\n");
            if (obj is IPositionAware pa) pa.BaseOffset = w.Position;
            obj.WriteTo(w);
            w.Write("\nendobj\n");
        }

        var trailerBase = new PdfDictionary();
        foreach (var key in new[] { "Root", "Info", "Encrypt" })
            if (_file.Trailer[key] is { } v) trailerBase[key] = v;
        trailerBase["ID"] = BuildId();

        bool useStream = _file.LastXrefIsStream || _file.Reconstructed;
        if (!useStream)
        {
            long xrefPos = w.Position;
            w.Write("xref\n");
            var nums = WrittenOffsets.Keys.OrderBy(k => k).ToList();
            int i = 0;
            while (i < nums.Count)
            {
                int j = i;
                while (j + 1 < nums.Count && nums[j + 1] == nums[j] + 1) j++;
                w.Write($"{nums[i]} {j - i + 1}\n");
                for (int k = i; k <= j; k++)
                    w.Write($"{WrittenOffsets[nums[k]]:D10} {_objects[nums[k]].Gen:D5} n\r\n");
                i = j + 1;
            }
            var trailer = trailerBase;
            trailer["Size"] = new PdfNumber(Math.Max(_file.Size, _next));
            trailer["Prev"] = new PdfNumber(_file.LastXrefOffset);
            w.Write("trailer\n");
            trailer.WriteTo(w);
            w.Write($"\nstartxref\n{xrefPos}\n%%EOF\n");
        }
        else
        {
            int xrefNum = _next++;
            long xrefPos = w.Position;
            var entries = new SortedDictionary<int, (int Type, long F2, int F3)>();
            if (_file.Reconstructed)
            {
                // Original xref is unusable: emit a complete table (no /Prev).
                entries[0] = (0, 0, 65535);
                foreach (var (num, e) in _file.Xref)
                    if (e.Type == 1) entries[num] = (1, e.Offset, e.Gen);
                    else if (e.Type == 2) entries[num] = (2, e.StreamNum, e.IndexInStream);
            }
            foreach (var (num, off) in WrittenOffsets) entries[num] = (1, off, _objects[num].Gen);
            entries[xrefNum] = (1, xrefPos, 0);

            long maxF2 = entries.Values.Max(e => e.F2);
            int w2 = maxF2 <= 0xFFFFFFFFL ? 4 : 5;
            var data = new MemoryStream();
            var index = new PdfArray();
            var keys = entries.Keys.ToList();
            int i = 0;
            while (i < keys.Count)
            {
                int j = i;
                while (j + 1 < keys.Count && keys[j + 1] == keys[j] + 1) j++;
                index.Add(new PdfNumber(keys[i]));
                index.Add(new PdfNumber(j - i + 1));
                for (int k = i; k <= j; k++)
                {
                    var (t, f2, f3) = entries[keys[k]];
                    data.WriteByte((byte)t);
                    for (int b = w2 - 1; b >= 0; b--) data.WriteByte((byte)(f2 >> (8 * b)));
                    data.WriteByte((byte)(f3 >> 8));
                    data.WriteByte((byte)f3);
                }
                i = j + 1;
            }
            var dict = trailerBase;
            dict["Type"] = new PdfName("XRef");
            dict["Size"] = new PdfNumber(Math.Max(_file.Size, xrefNum + 1));
            dict["W"] = PdfArray.Of(1, w2, 2);
            dict["Index"] = index;
            if (!_file.Reconstructed && _file.LastXrefOffset >= 0) dict["Prev"] = new PdfNumber(_file.LastXrefOffset);
            w.Write($"{xrefNum} 0 obj\n");
            new PdfStream(dict, data.ToArray()).WriteTo(w);
            w.Write($"\nendobj\nstartxref\n{xrefPos}\n%%EOF\n");
        }
        return w.ToArray();
    }

    private PdfArray BuildId()
    {
        var fresh = RandomNumberGenerator.GetBytes(16);
        PdfObject first = new PdfString(fresh, true);
        if (_file.Resolve(_file.Trailer["ID"]) is PdfArray old && old.Count >= 1 && old[0] is PdfString s0)
            first = s0;
        return new PdfArray([first, new PdfString(RandomNumberGenerator.GetBytes(16), true)]);
    }
}

/// <summary>Objects that need to know where they were written inside the output buffer.</summary>
public interface IPositionAware
{
    long BaseOffset { get; set; }
}
