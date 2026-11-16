using MicroadsOffice.Core.Office;
using MicroadsOffice.Core.Pdf;

namespace MicroadsOffice.Core.Signing;

/// <summary>
/// Compares a signed revision with the final file object by object and decides whether every change made
/// afterwards is a pure "add another signature" change (new signature dictionaries, signature widgets and
/// their appearances, AcroForm /Fields + /SigFlags, page /Annots additions, /V on empty signature fields).
/// Anything else (content streams, other annotations, metadata, deleted objects...) is reported as a modification.
/// </summary>
public static class RevisionDiff
{
    public static bool OnlySignatureChanges(byte[] full, long revisionLength, out string reason)
    {
        reason = "";
        PdfFile oldF, newF;
        try
        {
            oldF = PdfFile.Load(full[..(int)revisionLength]);
            newF = PdfFile.Load(full);
        }
        catch (Exception ex) { reason = Loc.T("Không phân tích được các phiên bản: ", "Could not analyse the revisions: ") + ex.Message; return false; }

        foreach (var (num, e) in newF.Xref)
        {
            var oldEntry = oldF.Xref.TryGetValue(num, out var oe) ? oe : default;
            bool oldExists = oldEntry.Type != 0;
            if (e.Type == 0)
            {
                if (oldExists) { reason = Loc.T($"Đối tượng {num} bị xóa sau khi ký.", $"Object {num} was deleted after signing."); return false; }
                continue;
            }
            PdfObject? n, o;
            try { n = newF.GetObject(num); o = oldExists ? oldF.GetObject(num) : null; }
            catch { reason = Loc.T($"Đối tượng {num} không đọc được.", $"Object {num} could not be read."); return false; }
            if (n == null) continue;
            if (n is PdfStream ns && ns.Dict.GetName("Type") is "XRef" or "ObjStm") continue;
            if (o == null)
            {
                if (!IsAllowedNew(newF, n)) { reason = Loc.T($"Đối tượng mới {num} ({Describe(n)}) không phải thành phần chữ ký.", $"New object {num} ({Describe(n)}) is not part of a signature."); return false; }
                continue;
            }
            if (Same(o, n)) continue;
            if (!IsAllowedModification(oldF, newF, o, n)) { reason = Loc.T($"Đối tượng {num} ({Describe(o)}) bị thay đổi sau khi ký.", $"Object {num} ({Describe(o)}) was changed after signing."); return false; }
        }
        return true;
    }

    private static string Describe(PdfObject o)
    {
        var d = o is PdfStream s ? s.Dict : o as PdfDictionary;
        return d?.GetName("Type") ?? d?.GetName("Subtype") ?? o.GetType().Name.Replace("Pdf", "");
    }

    private static bool Same(PdfObject a, PdfObject b)
    {
        if (a is PdfStream sa && b is PdfStream sb)
            return sa.Data.AsSpan().SequenceEqual(sb.Data) && Same(sa.Dict, sb.Dict);
        return a.ToBytes().AsSpan().SequenceEqual(b.ToBytes());
    }

    private static bool IsSigWidgetOrField(PdfFile f, PdfDictionary d)
    {
        if (d.GetName("FT") == "Sig") return true;
        var parent = f.ResolveDict(d["Parent"]);
        return d.GetName("Subtype") == "Widget" && parent?.GetName("FT") == "Sig";
    }

    private static bool IsAllowedNew(PdfFile f, PdfObject n)
    {
        switch (n)
        {
            case PdfStream s:
                // appearance streams and their images
                return s.Dict.GetName("Subtype") is "Form" or "Image";
            case PdfDictionary d:
                if (d.GetName("Type") == "Sig" || d.GetName("Type") == "DocTimeStamp") return true;
                if (IsSigWidgetOrField(f, d)) return true;
                if (d["Fields"] != null && d.Keys.All(k => k is "Fields" or "SigFlags" or "DA" or "DR" or "NeedAppearances")) return true; // new AcroForm
                if (d.GetName("Type") == "ExtGState") return true;
                if (d.Keys.All(k => k is "N" or "D" or "R")) return true; // appearance dictionary
                return false;
            case PdfArray a:
                return a.Items.All(i => i is PdfReference); // new /Annots or /Fields array
            default:
                return false;
        }
    }

    private static bool IsAllowedModification(PdfFile oldF, PdfFile newF, PdfObject o, PdfObject n)
    {
        if (o is PdfArray oa && n is PdfArray na) return IsAppendOnly(newF, oa, na);
        if (o is not PdfDictionary od || n is not PdfDictionary nd) return false;
        var changed = od.Keys.Union(nd.Keys).Where(k =>
        {
            var a = od[k]; var b = nd[k];
            return a == null || b == null || !Same(a, b);
        }).ToList();
        if (changed.Count == 0) return true;

        string? type = od.GetName("Type");
        if (type == "Catalog")
        {
            if (changed.Any(k => k != "AcroForm")) return false;
            var a = od["AcroForm"]; var b = nd["AcroForm"];
            if (b is PdfReference) return a == null || a is PdfReference ar && b is PdfReference br && ar.Number == br.Number || a is PdfDictionary;
            if (a is PdfDictionary ad && b is PdfDictionary bd) return IsAllowedModification(oldF, newF, ad, bd);
            return a == null && b is PdfDictionary;
        }
        if (type == "Page")
        {
            if (changed.Any(k => k != "Annots")) return false;
            var a = oldF.Resolve(od["Annots"]) as PdfArray ?? new PdfArray();
            var b = newF.Resolve(nd["Annots"]) as PdfArray;
            return b != null && IsAppendOnly(newF, a, b);
        }
        if (od["Fields"] != null || nd["Fields"] != null) // AcroForm
        {
            if (changed.Any(k => k is not ("Fields" or "SigFlags"))) return false;
            if (!changed.Contains("Fields")) return true;
            var a = oldF.Resolve(od["Fields"]) as PdfArray ?? new PdfArray();
            var b = newF.Resolve(nd["Fields"]) as PdfArray;
            return b != null && IsAppendOnly(newF, a, b);
        }
        // Signing an existing empty signature field: /V and the appearance may be set.
        if (IsSigWidgetOrField(oldF, od))
        {
            if (od["V"] != null && changed.Contains("V")) return false; // replacing an existing signature is not allowed
            return changed.All(k => k is "V" or "AP" or "AS");
        }
        return false;
    }

    /// <summary>The new array must start with all old items; appended items must be signature fields/widgets.</summary>
    private static bool IsAppendOnly(PdfFile newF, PdfArray a, PdfArray b)
    {
        if (b.Count < a.Count) return false;
        for (int i = 0; i < a.Count; i++) if (!Same(a[i], b[i])) return false;
        for (int i = a.Count; i < b.Count; i++)
        {
            var d = newF.ResolveDict(b[i]);
            if (d == null || !IsSigWidgetOrField(newF, d)) return false;
        }
        return true;
    }
}
