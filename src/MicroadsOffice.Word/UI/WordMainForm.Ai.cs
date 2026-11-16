using MicroadsOffice.AI;
using MicroadsOffice.AI.Mcp;
using MicroadsOffice.AI.Providers;
using MicroadsOffice.AI.UI;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;
using MicroadsOffice.Word.AI;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.UI;

/// <summary>AI assistant integration: the side pane, the ribbon button and the IAiHost implementation for Word.</summary>
sealed partial class WordMainForm : IAiHost
{
    private Panel? _aiContent;
    private AiPane? _aiPane;
    private Splitter? _aiSplit;
    private static WordMainForm? _lastActive;

    /// <summary>The assistant pane (created on first use).</summary>
    public AiPane? AiPane => _aiPane;

    private void InitAi(Panel content)
    {
        _aiContent = content;
        Activated += (_, _) => _lastActive = this;
        _lastActive ??= this;
        // Claude Desktop bridge: always listening, but every request is refused until the user allows it in Settings.
        AiBridge.Start(AppKind.Word, () => _lastActive is { IsDisposed: false } w ? w : Windows.FirstOrDefault(x => !x.IsDisposed), this);
    }

    /// <summary>Shows or hides the assistant pane (ribbon "Trợ lý AI").</summary>
    public void ToggleAi() => ShowAi(_aiPane == null || !_aiPane.Visible);

    public void ShowAi(bool show)
    {
        if (show && _aiPane == null && _aiContent != null)
        {
            _aiPane = new AiPane(this);
            _aiPane.CloseRequested += (_, _) => ShowAi(false);
            _aiSplit = new Splitter { Dock = DockStyle.Right, Width = 3, BackColor = Theme.Border, Name = "aiSplitter" };
            _aiContent.Controls.Add(_aiSplit);
            _aiContent.Controls.Add(_aiPane);
            _aiContent.Controls.SetChildIndex(_editorHost, 0); // the editor keeps the Fill slot
        }
        if (_aiPane == null) return;
        _aiPane.Visible = show;
        if (_aiSplit != null) _aiSplit.Visible = show;
        UpdateUi();
        if (!show) _ed.Focus();
    }

    // ------------------------------------------------------------------ IAiHost

    AppKind IAiHost.Kind => AppKind.Word;
    Control IAiHost.Owner => this;
    string IAiHost.DocumentName => DisplayName;
    string? IAiHost.SelectionText => _ed.HasSelection ? _ed.SelectedText : null;
    bool IAiHost.CanEdit => true;

    string IAiHost.GetDocumentContext(int maxChars) => WordMarkdown.Export(_ed.Document, maxChars);

    string IAiHost.ActionInstructions =>
        """
        Action block for Microads Word. Write the content as Markdown inside:
        ```microads-document mode=MODE
        ...content...
        ```
        MODE is one of:
        - insert: add the content after the paragraph where the cursor is (default for new content);
        - replace_selection: replace exactly the selected text (only when the user selected something and asks to change it; keep the same language and level of formatting as the selection);
        - replace_document: replace the whole document (only when the user explicitly asks to rewrite/redo the whole document, and only if the <document> you received was complete — never when it says it was truncated or contains [image]);
        - append: add at the end of the document.
        Supported Markdown: # / ## / ### headings (they become the document's Heading 1-3 styles), paragraphs, **bold**, *italic*, __underline__, bullet and numbered lists (indent 2 spaces per level), pipe tables (first row = header), > quotes, `code`, "\pagebreak" on its own line, and alignment blocks:
        :::center
        CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM
        **Độc lập - Tự do - Hạnh phúc**
        :::
        (also :::right and :::justify). Inside an alignment block each line is its own line. Do not use HTML. Do not wrap the whole document in a code block.
        Use only ONE action block per answer.
        """;

    IReadOnlyList<AiQuickAction> IAiHost.QuickActions =>
    [
        new("Summarize", Loc.T("Tóm tắt", "Summarize"), Loc.T("Tóm tắt nội dung chính của văn bản này thành các ý ngắn gọn (trả lời trong khung chat, không sửa văn bản).", "Summarize the key points of this document as short bullets (answer in the chat, do not change the document).")),
        new("Rewrite", Loc.T("Viết lại", "Rewrite"), Loc.T("Viết lại đoạn đang chọn cho rõ ràng, mạch lạc, văn phong trang trọng; giữ nguyên ý. Thay đúng vùng chọn.", "Rewrite the selected text to be clearer and more formal, keeping the meaning. Replace the selection."), NeedsSelection: true),
        new("Proofread", Loc.T("Sửa chính tả", "Proofread"), Loc.T("Kiểm tra chính tả, dấu câu và ngữ pháp. Nếu có vùng chọn: sửa và thay đúng vùng chọn, giữ nguyên ý. Nếu không có vùng chọn: KHÔNG ghi vào văn bản, chỉ liệt kê các lỗi kèm cách sửa trong khung chat.", "Check spelling, punctuation and grammar. With a selection: fix it and replace exactly the selection, keeping the meaning. Without a selection: do NOT write into the document; list the errors and fixes in the chat only.")),
        new("ToEnglish", Loc.T("Dịch sang tiếng Anh", "Translate to English"), Loc.T("Dịch vùng chọn sang tiếng Anh, văn phong tự nhiên. Thay đúng vùng chọn.", "Translate the selection into natural English. Replace the selection."), NeedsSelection: true),
        new("ToVietnamese", Loc.T("Dịch sang tiếng Việt", "Translate to Vietnamese"), Loc.T("Dịch vùng chọn sang tiếng Việt, văn phong tự nhiên. Thay đúng vùng chọn.", "Translate the selection into natural Vietnamese. Replace the selection."), NeedsSelection: true),
        new("Shorten", Loc.T("Rút gọn", "Shorten"), Loc.T("Rút gọn vùng chọn còn khoảng một nửa độ dài, giữ ý chính. Thay đúng vùng chọn.", "Shorten the selection to about half its length, keeping the key points. Replace the selection."), NeedsSelection: true),
        new("Outline", Loc.T("Lập dàn ý", "Outline"), Loc.T("Dựa trên văn bản hiện có (hoặc chủ đề tôi sẽ nêu), lập dàn ý chi tiết với các tiêu đề cấp 1-3 và chèn vào văn bản.", "Based on the current document (or the topic I give), create a detailed outline with level 1-3 headings and insert it into the document.")),
        new("OfficialLetter", Loc.T("Soạn công văn", "Official letter"), Loc.T("Soạn một công văn hành chính mẫu theo Nghị định 30/2020/NĐ-CP (quốc hiệu, tiêu ngữ, số ký hiệu, trích yếu, nội dung, nơi nhận, chữ ký), để trống các thông tin chưa biết, rồi chèn vào văn bản.", "Draft a sample Vietnamese official letter following Decree 30/2020/ND-CP (national motto, number, subject, body, recipients, signature), leaving unknown details blank, and insert it into the document.")),
    ];

    AiApplyResult IAiHost.Apply(AiActionBlock block)
    {
        if (block.Kind != "document") throw new AiException(Loc.T("Khối nội dung không dành cho Word.", "This content block is not for Word."));
        return WriteMarkdown(block.Body, block.Mode);
    }

    AiApplyResult IAiHost.InsertAnswer(string markdown, bool replaceSelection) => WriteMarkdown(markdown, replaceSelection ? "replace_selection" : "insert");

    void IAiHost.Undo() { if (_ed.CanUndo) { _ed.Undo(); _ed.Focus(); } }

    long IAiHost.EditSeq => _ed.EditSeq;

    /// <summary>
    /// Rewriting the whole document is only allowed when the model saw all of it and Markdown can carry all of it:
    /// otherwise text past the context limit, pictures and fields would be lost.
    /// </summary>
    private void CheckWholeRewrite(Document d)
    {
        bool hasObjects = d.Paragraphs.Any(p => p.Runs.Any(r => r.Object != null));
        bool tooLong = d.PlainText.Length > AiDefaults.MaxContextChars;
        if (hasObjects || tooLong)
            throw new AiException(hasObjects
                ? Loc.T("Văn bản có hình ảnh hoặc trường tự động nên trợ lý không viết lại toàn bộ (sẽ làm mất chúng). Hãy chọn phần cần sửa rồi yêu cầu lại.",
                        "The document has pictures or fields, so the assistant does not rewrite all of it (they would be lost). Select the part to change and ask again.")
                : Loc.T("Văn bản quá dài để trợ lý xem hết nên không viết lại toàn bộ. Hãy chọn phần cần sửa rồi yêu cầu lại.",
                        "The document is too long for the assistant to see completely, so it does not rewrite all of it. Select the part to change and ask again."), "unsafe_rewrite");
    }

    /// <summary>Writes assistant Markdown into the document as ONE undo step. Modes: insert, replace_selection, replace_document, append.</summary>
    public AiApplyResult WriteMarkdown(string markdown, string mode)
    {
        var d = _ed.Document;
        var blocks = WordMarkdown.ToBlocks(d, markdown);
        if (blocks.Count == 0) throw new AiException(Loc.T("Nội dung trống.", "The content is empty."));
        if (mode == "replace_selection" && !_ed.HasSelection) mode = "insert";
        if (mode == "replace_document") CheckWholeRewrite(d);
        string summary = mode switch
        {
            "replace_selection" => Loc.T("Đã thay vùng chọn.", "Replaced the selection."),
            "replace_document" => Loc.T("Đã viết lại toàn bộ văn bản.", "Rewrote the whole document."),
            "append" => Loc.T("Đã thêm vào cuối văn bản.", "Added at the end of the document."),
            _ => Loc.T("Đã chèn vào văn bản.", "Inserted into the document."),
        };
        _ed.Edit("ai", () =>
        {
            DocPos pos;
            switch (mode)
            {
                case "replace_document":
                    d.Blocks.Clear();
                    var normal = d.Styles[StyleIds.Normal];
                    d.Blocks.Add(new Paragraph("", normal.Char, normal.Para) { MarkProps = normal.Char });
                    d.StructureChanged();
                    pos = new DocPos(0, 0);
                    break;
                case "replace_selection":
                    pos = Editing.Delete(d, _ed.SelStart, _ed.SelEnd);
                    break;
                case "append":
                    {
                        if (d.Blocks.Count == 0 || d.Blocks[^1] is Table)
                        {
                            var n = d.Styles[StyleIds.Normal];
                            d.Blocks.Add(new Paragraph("", n.Char, n.Para) { MarkProps = n.Char });
                            d.StructureChanged();
                        }
                        int last = LastBodyParagraph(d);
                        pos = new DocPos(last, d.Paragraphs[last].Length);
                        if (d.Paragraphs[last].Length > 0) pos = Editing.SplitParagraph(d, pos);
                        break;
                    }
                default:
                    {
                        // new content goes after the paragraph that holds the cursor (or into it when it is empty)
                        var at = _ed.HasSelection ? _ed.SelEnd : _ed.Caret;
                        var p = d.Paragraphs[at.Para];
                        pos = p.Length == 0 ? at : Editing.SplitParagraph(d, new DocPos(at.Para, p.Length));
                        break;
                    }
            }
            var target = d.Paragraphs[pos.Para];
            if (target.Length == 0 && mode != "replace_selection" && blocks[0] is Paragraph first)
            {
                target.Props = first.Props; target.MarkProps = first.MarkProps; target.Touch();
            }
            _ed.SetCaretInternal(Editing.PasteBlocks(d, pos, blocks));
        });
        _ed.ScrollToCaret();
        return new AiApplyResult(summary);
    }

    private static int LastBodyParagraph(Document d)
    {
        for (int i = d.Paragraphs.Count - 1; i >= 0; i--) if (d.Paragraphs[i].Cell == null) return i;
        return d.Paragraphs.Count - 1;
    }
}
