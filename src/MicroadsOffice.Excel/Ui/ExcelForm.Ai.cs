using MicroadsOffice.AI;
using MicroadsOffice.AI.Mcp;
using MicroadsOffice.AI.Providers;
using MicroadsOffice.AI.UI;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Excel.AI;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

/// <summary>AI assistant integration for Excel: side pane, ribbon button, IAiHost (reads cells, writes validated JSON actions).</summary>
public sealed partial class ExcelForm : IAiHost
{
    private AiPane? _aiPane;
    private Splitter? _aiSplit;

    public AiPane? AiPane => _aiPane;

    private void InitAi()
    {
        AiBridge.Start(AppKind.Excel, () => IsDisposed ? null : this, this);
    }

    public void ToggleAi() => ShowAi(_aiPane == null || !_aiPane.Visible);

    public void ShowAi(bool show)
    {
        if (show && _aiPane == null)
        {
            _aiPane = new AiPane(this);
            _aiPane.CloseRequested += (_, _) => ShowAi(false);
            _aiSplit = new Splitter { Dock = DockStyle.Right, Width = 3, BackColor = Theme.Border, Name = "aiSplitter" };
            _main.Controls.Add(_aiSplit);
            _main.Controls.Add(_aiPane);
            _main.Controls.SetChildIndex(Grid, 0);
        }
        if (_aiPane == null) return;
        _aiPane.Visible = show;
        if (_aiSplit != null) _aiSplit.Visible = show;
        Check("btnAiAssistant", show);
        if (!show) Grid.Focus();
    }

    // ------------------------------------------------------------------ IAiHost

    AppKind IAiHost.Kind => AppKind.Excel;
    Control IAiHost.Owner => this;
    string IAiHost.DocumentName => _doc.DisplayName;
    bool IAiHost.CanEdit => true;

    string? IAiHost.SelectionText
    {
        get
        {
            var sel = Grid.Selection;
            var s = _doc.Sheet;
            var sb = new System.Text.StringBuilder();
            bool any = false;
            for (int r = sel.R1; r <= Math.Min(sel.R2, sel.R1 + 500); r++)
            {
                var row = new List<string>();
                for (int c = sel.C1; c <= Math.Min(sel.C2, sel.C1 + 50); c++)
                {
                    var cell = s.GetCell(r, c);
                    var t = SheetOps.DisplayText(s, r, c);
                    if (cell?.Formula != null) t += " [=" + cell.Formula + "]";
                    if (t.Length > 0) any = true;
                    row.Add(t.Replace('\t', ' ').Replace('\n', ' '));
                }
                sb.AppendLine(string.Join("\t", row));
            }
            if (!any && sel.IsSingle) return null;
            return $"{s.Name}!{sel} (tab-separated, formulas in [=…]):\n" + sb.ToString().TrimEnd();
        }
    }

    string IAiHost.GetDocumentContext(int maxChars) => SheetActions.Describe(_doc.Book, Grid.Selection, Grid.Active, maxChars);

    string IAiHost.ActionInstructions => SheetActions.Guide + "\nSupported functions: " + SheetActions.FunctionList();

    IReadOnlyList<AiQuickAction> IAiHost.QuickActions =>
    [
        new("Explain", Loc.T("Giải thích công thức", "Explain formula"), Loc.T("Giải thích công thức ở ô đang chọn (hoặc các công thức trong vùng chọn) bằng lời dễ hiểu, kèm ví dụ. Không sửa trang tính.", "Explain the formula in the selected cell (or the formulas in the selection) in plain language with an example. Do not change the sheet.")),
        new("WriteFormula", Loc.T("Viết công thức", "Write a formula"), Loc.T("Hãy viết công thức cho ô đang chọn. Yêu cầu của tôi: ", "Write a formula for the selected cell. What I need: "), Tooltip: Loc.T("Điền yêu cầu vào ô hỏi rồi gửi", "Complete the request in the box, then send")),
        new("Analyze", Loc.T("Phân tích dữ liệu", "Analyze data"), Loc.T("Phân tích dữ liệu trong trang tính: xu hướng, giá trị nổi bật, bất thường, tổng hợp chính. Trả lời ngắn gọn bằng các ý, có số liệu cụ thể.", "Analyze the data in the sheet: trends, highlights, anomalies and key totals. Answer in short bullets with concrete numbers.")),
        new("Summary", Loc.T("Bảng tổng hợp", "Summary table"), Loc.T("Tạo một trang tính mới \"Tổng hợp\" tóm tắt dữ liệu hiện có (tổng, trung bình, đếm theo nhóm phù hợp) bằng công thức tham chiếu dữ liệu gốc, có tiêu đề in đậm và tự vừa độ rộng cột.", "Create a new sheet \"Summary\" that summarises the existing data (totals, averages, counts by a suitable group) using formulas that reference the source data, with bold headers and autofit columns.")),
        new("Chart", Loc.T("Vẽ biểu đồ", "Make a chart"), Loc.T("Chọn loại biểu đồ phù hợp nhất cho dữ liệu (hoặc vùng đang chọn) và chèn biểu đồ có tiêu đề rõ ràng.", "Choose the most suitable chart type for the data (or the selection) and insert a chart with a clear title.")),
        new("Clean", Loc.T("Làm sạch dữ liệu", "Clean data"), Loc.T("Kiểm tra vùng chọn (hoặc dữ liệu) để tìm khoảng trắng thừa, viết hoa không thống nhất, số lưu dạng chữ, ô trống, trùng lặp. Liệt kê vấn đề và sửa trực tiếp những gì sửa an toàn được.", "Check the selection (or the data) for extra spaces, inconsistent capitalisation, numbers stored as text, blanks and duplicates. List the issues and fix what can be fixed safely.")),
        new("Template", Loc.T("Tạo bảng mẫu", "Create a template"), Loc.T("Tạo bảng tính mẫu theo yêu cầu sau, có tiêu đề, định dạng số, công thức tổng và đường viền: ", "Create a spreadsheet template for the following, with headers, number formats, total formulas and borders: "), Tooltip: Loc.T("Ví dụ: bảng lương 10 nhân viên, sổ thu chi tháng", "e.g. payroll for 10 employees, monthly cash book")),
    ];

    AiApplyResult IAiHost.Apply(AiActionBlock block)
    {
        if (block.Kind != "sheet") throw new AiException(Loc.T("Khối nội dung không dành cho Excel.", "This content block is not for Excel."));
        return ApplySheetActions(block.Body);
    }

    /// <summary>Validates then applies a JSON action list as ONE undo step.</summary>
    public AiApplyResult ApplySheetActions(string json)
    {
        if (Grid.IsEditing) Grid.CommitEdit(0, 0);
        var ops = SheetActions.Parse(json, _doc.Book);
        var redo = _doc.SaveRedo(); bool wasModified = _doc.Modified;
        _doc.Checkpoint(Loc.T("Trợ lý AI", "AI assistant"));
        string summary;
        try
        {
            using var g = CreateGraphics();
            summary = SheetActions.Apply(ops, _doc.Book, (text, st) =>
                CellPainter.Measure(g, text.Split('\n').OrderByDescending(x => x.Length).First(), CellPainter.GetFont(st.FontName, st.FontSize, CellPainter.StyleOf(st)), true, 0).Width / Dpi.Factor);
        }
        catch (Exception e)
        {
            // all or nothing: put the workbook, its undo/redo history and the modified flag back exactly as they were
            _doc.Rollback(redo, wasModified);
            AfterAi();
            throw e as AiException ?? new AiException(Loc.T("Không ghi được vào trang tính: ", "Could not write to the sheet: ") + e.Message, inner: e);
        }
        AfterAi();
        return new AiApplyResult(summary);
    }

    AiApplyResult IAiHost.InsertAnswer(string markdown, bool replaceSelection)
    {
        if (Grid.IsEditing) Grid.CommitEdit(0, 0);
        var a = Grid.Active;
        var table = Markdown.Parse(markdown).OfType<MdTable>().FirstOrDefault();
        var formula = markdown.Split('\n').Select(l => l.Trim().Trim('`')).FirstOrDefault(l => l.StartsWith('='));
        if (table == null && formula != null)
        {
            try { SheetActions.CheckFormula(formula); }
            catch (FormatException e) { throw new AiException(e.Message); } // nothing touched yet
        }
        _doc.Checkpoint(Loc.T("Trợ lý AI", "AI assistant"));
        string summary;
        if (table != null)
        {
            for (int r = 0; r < table.Rows.Count; r++)
                for (int c = 0; c < table.Rows[r].Count; c++)
                    SheetOps.SetInput(_doc.Sheet, a.Row + r, a.Col + c, Markdown.Plain(table.Rows[r][c]));
            summary = Loc.F("Đã ghi bảng {0}×{1} từ ô {2}.", "Wrote a {0}×{1} table at {2}.", table.Rows.Count, table.Rows.Max(x => x.Count), a);
        }
        else if (formula != null)
        {
            SheetOps.SetInput(_doc.Sheet, a.Row, a.Col, formula);
            summary = Loc.F("Đã đặt công thức vào ô {0}.", "Put the formula in {0}.", a);
        }
        else
        {
            SheetOps.SetInput(_doc.Sheet, a.Row, a.Col, Markdown.ToPlainText(markdown));
            summary = Loc.F("Đã ghi vào ô {0}.", "Wrote into {0}.", a);
        }
        AfterAi();
        return new AiApplyResult(summary);
    }

    void IAiHost.Undo() => CmdUndo();

    long IAiHost.EditSeq => _doc.EditSeq;

    private void AfterAi()
    {
        Changed();
        Grid.Doc = _doc;
        Grid.RebindSelection();
        Tabs.Book = _doc.Book;
        UpdateAll();
        SwitchSheet(_doc.Book.ActiveSheet); // the assistant may have added / chosen another sheet
    }
}
