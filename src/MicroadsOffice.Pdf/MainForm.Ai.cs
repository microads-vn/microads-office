using System.Text;
using MicroadsOffice.AI;
using MicroadsOffice.AI.Mcp;
using MicroadsOffice.AI.Providers;
using MicroadsOffice.AI.UI;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;

namespace MicroadsOffice;

/// <summary>AI assistant for PDF: read-only (summaries, questions, translation of the open document; answers stay in the pane).</summary>
public sealed partial class MainForm : IAiHost
{
    private Panel? _aiHost;
    private AiPane? _aiPane;
    private Splitter? _aiSplit;
    private ToolStripButton? _aiButton;

    public AiPane? AiPane => _aiPane;

    private void InitAi(Panel host, RibbonTab home)
    {
        _aiHost = host;
        home.Separator();
        _aiButton = home.Button("btnAiAssistant", Loc.T("Trợ lý\nAI", "AI\nassistant"), Icons.Info, (_, _) => ToggleAi(),
            Loc.T("Mở Trợ lý AI (ChatGPT, Claude…): tóm tắt, hỏi đáp, dịch tài liệu PDF đang mở", "Open the AI assistant (ChatGPT, Claude…): summarise, ask about and translate the open PDF"));
        _aiButton.Image = AiIcons.Sparkle(true);
        Cmd(_aiButton);
        AiBridge.Start(AppKind.Pdf, () => IsDisposed ? null : this, this);
    }

    public void ToggleAi() => ShowAi(_aiPane == null || !_aiPane.Visible);

    public void ShowAi(bool show)
    {
        if (show && _aiPane == null && _aiHost != null)
        {
            _aiPane = new AiPane(this);
            _aiPane.CloseRequested += (_, _) => ShowAi(false);
            _aiSplit = new Splitter { Dock = DockStyle.Right, Width = 3, BackColor = Theme.Border, Name = "aiSplitter" };
            _aiHost.Controls.Add(_aiSplit);
            _aiHost.Controls.Add(_aiPane);
            foreach (Control c in _aiHost.Controls.Cast<Control>().Where(c => c.Dock == DockStyle.Fill).ToList()) _aiHost.Controls.SetChildIndex(c, 0);
        }
        if (_aiPane == null) return;
        _aiPane.Visible = show;
        if (_aiSplit != null) _aiSplit.Visible = show;
        if (_aiButton != null) _aiButton.Checked = show;
    }

    AppKind IAiHost.Kind => AppKind.Pdf;
    Control IAiHost.Owner => this;
    string IAiHost.DocumentName => Current?.DisplayName ?? Loc.T("(chưa mở tệp PDF)", "(no PDF open)");
    bool IAiHost.CanEdit => false;
    string IAiHost.ActionInstructions => "";

    string? IAiHost.SelectionText
    {
        get { var t = Current?.Viewer.GetSelectedText(); return string.IsNullOrWhiteSpace(t) ? null : t; }
    }

    string IAiHost.GetDocumentContext(int maxChars)
    {
        var v = Current;
        if (v == null) return Loc.T("(Chưa mở tệp PDF nào.)", "(No PDF is open.)");
        var pdf = v.Pdf;
        var sb = new StringBuilder();
        sb.AppendLine($"Pages: {pdf.PageCount}; the user is viewing page {v.Viewer.CurrentPage + 1}.");
        for (int i = 0; i < pdf.PageCount; i++)
        {
            string text;
            try { text = pdf.GetPageText(i); } catch { text = ""; }
            sb.AppendLine($"--- Page {i + 1} ---");
            sb.AppendLine(text.Trim());
            if (sb.Length > maxChars) { sb.AppendLine($"[… pages {i + 2}-{pdf.PageCount} truncated …]"); break; }
        }
        if (sb.ToString().Count(char.IsLetterOrDigit) < 20 * Math.Max(1, pdf.PageCount))
            sb.AppendLine("(This PDF has little or no text layer — it may be a scan; say so if you cannot answer.)");
        return sb.ToString();
    }

    IReadOnlyList<AiQuickAction> IAiHost.QuickActions =>
    [
        new("Summarize", Loc.T("Tóm tắt", "Summarize"), Loc.T("Tóm tắt tài liệu PDF này: mục đích, các ý chính, số liệu và mốc thời gian quan trọng.", "Summarize this PDF: purpose, key points, important figures and dates.")),
        new("KeyTerms", Loc.T("Điểm cần lưu ý", "Key points to watch"), Loc.T("Liệt kê các nghĩa vụ, thời hạn, khoản tiền, điều kiện và rủi ro cần lưu ý trong tài liệu này (nếu là hợp đồng/văn bản pháp lý).", "List the obligations, deadlines, amounts, conditions and risks to watch in this document (if it is a contract/legal text).")),
        new("TranslatePage", Loc.T("Dịch trang đang xem", "Translate this page"), Loc.T("Dịch nội dung trang tôi đang xem sang tiếng Việt (nếu đã là tiếng Việt thì dịch sang tiếng Anh).", "Translate the page I am viewing into Vietnamese (or into English if it is already Vietnamese).")),
        new("ExplainSelection", Loc.T("Giải thích đoạn chọn", "Explain selection"), Loc.T("Giải thích đoạn tôi đang chọn bằng lời dễ hiểu.", "Explain the selected passage in plain language."), NeedsSelection: true),
        new("Tables", Loc.T("Trích bảng số liệu", "Extract tables"), Loc.T("Trích các bảng số liệu trong tài liệu thành bảng Markdown để tôi sao chép sang Excel.", "Extract the data tables in the document as Markdown tables I can copy into Excel.")),
    ];

    AiApplyResult IAiHost.Apply(AiActionBlock block) => throw new AiException(Loc.T("Tệp PDF chỉ đọc trong Trợ lý AI.", "PDFs are read-only for the AI assistant."));
    AiApplyResult IAiHost.InsertAnswer(string markdown, bool replaceSelection) => throw new AiException(Loc.T("Tệp PDF chỉ đọc trong Trợ lý AI.", "PDFs are read-only for the AI assistant."));
    void IAiHost.Undo() { }
    long IAiHost.EditSeq => 0;
}
