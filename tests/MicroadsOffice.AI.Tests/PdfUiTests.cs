using System.Text;
using MicroadsOffice.AI.UI;

namespace MicroadsOffice.AI.Tests;

/// <summary>The assistant in Microads PDF: read-only (answers stay in the pane, nothing is written into the PDF).</summary>
public static class PdfUiTests
{
    /// <summary>A one-page PDF with a text layer (Helvetica), built by hand so the test needs no sample file.</summary>
    public static string MakePdf(string path, string text)
    {
        var content = $"BT /F1 18 Tf 72 760 Td ({text}) Tj ET";
        var objs = new[]
        {
            "<</Type/Catalog/Pages 2 0 R>>",
            "<</Type/Pages/Kids[3 0 R]/Count 1>>",
            "<</Type/Page/Parent 2 0 R/MediaBox[0 0 595 842]/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>",
            $"<</Length {content.Length}>>\nstream\n{content}\nendstream",
            "<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>",
        };
        var sb = new StringBuilder("%PDF-1.4\n");
        var offsets = new List<int>();
        for (int i = 0; i < objs.Length; i++) { offsets.Add(sb.Length); sb.Append($"{i + 1} 0 obj\n{objs[i]}\nendobj\n"); }
        int xref = sb.Length;
        sb.Append($"xref\n0 {objs.Length + 1}\n0000000000 65535 f \n");
        foreach (var o in offsets) sb.Append($"{o:D10} 00000 n \n");
        sb.Append($"trailer\n<</Size {objs.Length + 1}/Root 1 0 R>>\nstartxref\n{xref}\n%%EOF\n");
        File.WriteAllBytes(path, Encoding.ASCII.GetBytes(sb.ToString()));
        return path;
    }

    public static void Run(Runner r)
    {
        var closer = Ui.StartCloser();
        var ai = new ScriptedProvider();
        AiService.Override = _ => ai;
        AiSettings.Update(x => { x.Provider = AiProviderKind.Compatible; x.CompatibleModel = "fake-model"; });
        MainForm? f = null;
        try
        {
            var pdf = MakePdf(Path.Combine(T.Out, "hop-dong-mau.pdf"), "Hop dong mau so 12 - thoi han 30 ngay");

            r.Test("P01 Home › Trợ lý AI opens a read-only assistant pane for the open PDF", () =>
            {
                f = (MainForm)PdfApp.CreateMainForm([pdf]);
                f.Show(); Ui.Pump(800);
                var btn = f.Controls.Find("ribbonHome", true).OfType<ToolStrip>().SelectMany(t => t.Items.Cast<ToolStripItem>()).FirstOrDefault(i => i.Name == "btnAiAssistant")
                          ?? throw new AssertException("btnAiAssistant not on the Home tab");
                T.True(btn.Image != null && btn.ToolTipText.Contains("PDF"), "icon + tooltip");
                btn.PerformClick(); Ui.Pump(100);
                T.True(f.AiPane is { Visible: true }, "pane visible");
                T.True(((ToolStripButton)btn).Checked, "checked");
                T.False(Ui.Find<CheckBox>(f.AiPane!, "chkAiAutoApply").Enabled, "no writing into PDFs");
                T.Eq(5, Ui.Find<FlowLayoutPanel>(f.AiPane!, "flowAiQuick").Controls.Count);
            });

            r.Test("P02 'Tóm tắt' sends the PDF text with page markers; the answer offers Copy only", () =>
            {
                ai.Replies.Enqueue("Hợp đồng mẫu số 12, thời hạn 30 ngày.\n```microads-document\n# Không được ghi\n```");
                Ui.Click(Ui.Find<Button>(f!.AiPane!, "btnAiQ_Summarize"));
                Ui.WaitRequest(f.AiPane!);
                var sent = ai.Requests.Last().Messages.Last().Text;
                T.True(sent.Contains("--- Page 1 ---") && sent.Contains("Hop dong mau so 12"), sent);
                T.True(sent.Contains("hop-dong-mau.pdf"), "file name");
                T.True(ai.Requests.Last().Instructions.Contains("read-only"), "told it is read-only");
                var ans = f.AiPane!.LastAnswer!;
                T.True(ans.FindAction("btnAiCopy") != null, "copy");
                T.True(ans.FindAction("btnAiInsert") == null && ans.FindAction("btnAiApply") == null && ans.FindAction("btnAiUndo") == null, "no write buttons");
            });

            r.Test("P03 'Giải thích đoạn chọn' asks for a selection first; Close hides the pane", () =>
            {
                int n = ai.Requests.Count;
                Ui.Click(Ui.Find<Button>(f!.AiPane!, "btnAiQ_ExplainSelection"));
                T.Eq(n, ai.Requests.Count);
                Ui.Click(Ui.Find<Button>(f.AiPane!, "btnAiClose"));
                T.False(f.AiPane!.Visible, "hidden");
            });
        }
        finally
        {
            AiService.Override = null;
            try { f?.Close(); Ui.Pump(300); } catch { }
            closer.Stop();
        }
    }
}
