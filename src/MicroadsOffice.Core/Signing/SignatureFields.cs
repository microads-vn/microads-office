using System.Drawing;
using MicroadsOffice.Core.Pdf;

namespace MicroadsOffice.Core.Signing;

/// <summary>A signature form field found in the document (signed or empty).</summary>
public sealed class SignatureFieldInfo
{
    public string FullName { get; init; } = "";
    public PdfReference? FieldRef { get; init; }
    public PdfReference? WidgetRef { get; init; }
    public PdfDictionary? SignatureDict { get; init; }
    public int PageIndex { get; init; } = -1;
    /// <summary>Widget rectangle in page user space (x1,y1 lower-left; width/height).</summary>
    public RectangleF Rect { get; init; }
    public bool IsSigned => SignatureDict != null;
}

public static class SignatureFields
{
    public static List<SignatureFieldInfo> Find(PdfFile f)
    {
        var result = new List<SignatureFieldInfo>();
        var acro = f.ResolveDict(f.Catalog?["AcroForm"]);
        if (acro == null) return result;
        var pages = f.GetPageRefs();
        var pageIndex = new Dictionary<int, int>();
        for (int i = 0; i < pages.Count; i++) pageIndex.TryAdd(pages[i].Number, i);

        // Map widget object numbers -> page (for widgets without /P).
        Dictionary<int, int>? annotPage = null;
        int FindPageOfWidget(PdfReference? widget, PdfDictionary wd)
        {
            if (wd["P"] is PdfReference p && pageIndex.TryGetValue(p.Number, out int pi)) return pi;
            if (widget == null) return -1;
            if (annotPage == null)
            {
                annotPage = new Dictionary<int, int>();
                for (int i = 0; i < pages.Count; i++)
                {
                    try
                    {
                        if (f.Resolve(f.ResolveDict(pages[i])?["Annots"]) is PdfArray arr)
                            foreach (var a in arr.Items.OfType<PdfReference>()) annotPage.TryAdd(a.Number, i);
                    }
                    catch { }
                }
            }
            return annotPage.TryGetValue(widget.Number, out int idx) ? idx : -1;
        }

        var visited = new HashSet<int>();
        void Walk(PdfObject? node, string parentName, string? inheritedFt, int depth)
        {
            if (depth > 32) return;
            if (node is PdfReference nr && !visited.Add(nr.Number)) return;
            var d = f.ResolveDict(node);
            if (d == null) return;
            string name = d["T"] is PdfString t ? t.ToText() : "";
            string full = name.Length == 0 ? parentName : parentName.Length == 0 ? name : parentName + "." + name;
            string? ft = d.GetName("FT") ?? inheritedFt;
            var kids = f.Resolve(d["Kids"]) as PdfArray;
            bool kidsAreFields = kids != null && kids.Items.Any(k => f.ResolveDict(k)?["T"] != null);
            if (kids != null && kidsAreFields)
            {
                foreach (var k in kids.Items) Walk(k, full, ft, depth + 1);
                return;
            }
            if (ft != "Sig") return;
            var sig = f.ResolveDict(d["V"]);
            // Widget: the field itself (merged) or its first kid widget.
            PdfReference? widgetRef = node as PdfReference;
            PdfDictionary widget = d;
            if (kids != null && kids.Count > 0)
            {
                widgetRef = kids[0] as PdfReference;
                widget = f.ResolveDict(kids[0]) ?? d;
            }
            var rect = RectangleF.Empty;
            if (f.Resolve(widget["Rect"]) is PdfArray ra && ra.Count >= 4)
            {
                var v = ra.Items.Select(x => (float)((f.Resolve(x) as PdfNumber)?.Value ?? 0)).ToArray();
                rect = RectangleF.FromLTRB(Math.Min(v[0], v[2]), Math.Min(v[1], v[3]), Math.Max(v[0], v[2]), Math.Max(v[1], v[3]));
            }
            result.Add(new SignatureFieldInfo
            {
                FullName = full,
                FieldRef = node as PdfReference,
                WidgetRef = widgetRef,
                SignatureDict = sig,
                PageIndex = FindPageOfWidget(widgetRef, widget),
                Rect = rect
            });
        }
        if (f.Resolve(acro["Fields"]) is PdfArray fields)
            foreach (var fld in fields.Items) Walk(fld, "", null, 0);
        return result;
    }

    public static HashSet<string> AllFieldNames(PdfFile f)
    {
        var names = new HashSet<string>(StringComparer.Ordinal);
        var acro = f.ResolveDict(f.Catalog?["AcroForm"]);
        if (acro == null) return names;
        var visited = new HashSet<int>();
        void Walk(PdfObject? node, string parent, int depth)
        {
            if (depth > 32) return;
            if (node is PdfReference r && !visited.Add(r.Number)) return;
            var d = f.ResolveDict(node);
            if (d == null) return;
            string name = d["T"] is PdfString t ? t.ToText() : "";
            string full = name.Length == 0 ? parent : parent.Length == 0 ? name : parent + "." + name;
            if (full.Length > 0) names.Add(full);
            if (f.Resolve(d["Kids"]) is PdfArray kids) foreach (var k in kids.Items) Walk(k, full, depth + 1);
        }
        if (f.Resolve(acro["Fields"]) is PdfArray fields) foreach (var x in fields.Items) Walk(x, "", 0);
        return names;
    }
}
