using System.Net.Http.Json;
using System.Text.Json.Nodes;
using MicroadsOffice.AI.Providers;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Dialogs;
using MicroadsOffice.UI;

namespace MicroadsOffice.AI.UI;

/// <summary>
/// "Báo cáo nội dung không phù hợp" — lets the user report an AI answer to Microads (Microsoft Store policy 11.16).
/// Nothing is sent unless the user clicks Send; the question/answer are included only while the box is ticked.
/// </summary>
public sealed class AiReportDialog : DialogBase
{
    public static readonly (string Id, string Vi, string En)[] Reasons =
    [
        ("harmful", "Có hại hoặc nguy hiểm", "Harmful or dangerous"),
        ("offensive", "Xúc phạm, thù ghét hoặc không phù hợp", "Offensive, hateful or inappropriate"),
        ("wrong", "Sai sự thật, gây hiểu nhầm", "False or misleading"),
        ("privacy", "Lộ thông tin riêng tư", "Reveals private information"),
        ("other", "Khác", "Other"),
    ];

    readonly string _question, _answer, _app, _provider, _model;
    readonly FlowLayoutPanel _reasons;
    readonly TextBox _comment;
    readonly CheckBox _include;
    readonly Label _result;
    readonly Button _send;

    /// <summary>For tests: the running send.</summary>
    public Task? Pending { get; private set; }
    public bool Sent { get; private set; }

    public AiReportDialog(AppKind app, string question, string answer) : base(Loc.T("Báo cáo nội dung AI", "Report AI content"), 480, 430)
    {
        Name = "aiReportDialog";
        _question = question; _answer = answer; _app = app.ToString().ToLowerInvariant();
        var s = AiSettings.Load();
        _provider = s.Provider.ToString(); _model = s.ModelFor(s.Provider);

        var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
        int w = Dpi.Scale(440);
        flow.Controls.Add(new Label { Text = Loc.T("Câu trả lời này có vấn đề gì?", "What is wrong with this answer?"), Font = Theme.UiFontBold, AutoSize = true });
        _reasons = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Margin = new Padding(Dpi.Scale(6), 0, 0, Dpi.Scale(6)) };
        foreach (var (id, vi, en) in Reasons)
            _reasons.Controls.Add(new RadioButton { Text = Loc.T(vi, en), Tag = id, Name = "rbAiReport_" + id, AutoSize = true, Checked = id == "harmful" });
        flow.Controls.Add(_reasons);
        flow.Controls.Add(new Label { Text = Loc.T("Ghi chú (không bắt buộc)", "Comment (optional)"), AutoSize = true });
        _comment = new TextBox { Multiline = true, Width = w, Height = Dpi.Scale(60), Name = "txtAiReportComment", ScrollBars = ScrollBars.Vertical };
        flow.Controls.Add(_comment);
        _include = new CheckBox { Text = Loc.T("Gửi kèm câu hỏi và câu trả lời để Microads xem xét", "Include the question and the answer for Microads to review"), Checked = true, AutoSize = true, MaximumSize = new Size(w, 0), Name = "chkAiReportInclude", Margin = new Padding(0, Dpi.Scale(6), 0, 0) };
        flow.Controls.Add(_include);
        flow.Controls.Add(new Label
        {
            Text = Loc.T("Báo cáo được gửi tới Microads (microads.vn) để cải thiện an toàn của Trợ lý AI. Nếu để chọn ô trên, câu hỏi và toàn bộ câu trả lời này (kể cả nội dung trợ lý đã soạn vào văn bản) được gửi kèm. Không gửi khóa API, tài khoản hay phần nào khác của văn bản.",
                         "The report goes to Microads (microads.vn) to improve the assistant's safety. With the box ticked, this question and the full answer (including what the assistant wrote into the document) are included. No API key, account or other part of your document is sent."),
            AutoSize = true, MaximumSize = new Size(w, 0), ForeColor = Theme.SubtleInk, Margin = new Padding(0, Dpi.Scale(6), 0, 0),
        });
        _result = new Label { AutoSize = true, MaximumSize = new Size(w, 0), Name = "lblAiReportResult", Margin = new Padding(0, Dpi.Scale(6), 0, 0) };
        flow.Controls.Add(_result);
        Body.Controls.Add(flow);

        CancelBtn = AddButton(Loc.T("Đóng", "Close"), "btnCancel", DialogResult.Cancel);
        _send = AddButton(Loc.T("Gửi báo cáo", "Send report"), "btnAiReportSend", DialogResult.None, primary: true, (_, _) => Send());
        CancelButton = CancelBtn;
    }

    /// <summary>Where reports go; tests override with MICROADS_AI_REPORT_URL (https, or http on loopback).</summary>
    public static string Endpoint
    {
        get
        {
            var v = Environment.GetEnvironmentVariable("MICROADS_AI_REPORT_URL");
            if (!string.IsNullOrWhiteSpace(v) && Uri.TryCreate(v, UriKind.Absolute, out var u) && (u.Scheme == "https" || (u.Scheme == "http" && u.IsLoopback))) return v;
            return "https://microads.vn/office/report.php";
        }
    }

    string Reason => _reasons.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked)?.Tag as string ?? "other";

    void Send()
    {
        if (!_include.Checked && _comment.Text.Trim().Length == 0)
        {
            _result.ForeColor = Theme.WarnAmber;
            _result.Text = Loc.T("Hãy ghi chú vấn đề, hoặc chọn gửi kèm câu hỏi và câu trả lời.", "Describe the problem, or include the question and the answer.");
            return;
        }
        var body = new JsonObject
        {
            ["app"] = _app, ["version"] = Product.VersionText, ["lang"] = Loc.Lang, ["provider"] = _provider, ["model"] = _model,
            ["reason"] = Reason, ["comment"] = _comment.Text.Trim(),
            ["question"] = _include.Checked ? Trim(_question, 4000) : "", ["answer"] = _include.Checked ? Trim(_answer, 8000) : "",
        };
        _send.Enabled = false;
        _result.ForeColor = Theme.SubtleInk; _result.Text = Loc.T("Đang gửi…", "Sending…");
        Pending = Run();
        async Task Run()
        {
            try
            {
                using var http = AiProvider.NewHttp(TimeSpan.FromSeconds(30));
                using var resp = await UiThread.On(this, http.PostAsJsonAsync(Endpoint, body));
                if (IsDisposed) return;
                if (!resp.IsSuccessStatusCode) throw new HttpRequestException(((int)resp.StatusCode).ToString());
                Sent = true;
                _result.ForeColor = Theme.OkGreen;
                _result.Text = Loc.T("Đã gửi. Cảm ơn bạn — Microads sẽ xem xét báo cáo này.", "Sent. Thank you — Microads will review this report.");
                CancelBtn!.Text = Loc.T("Xong", "Done");
            }
            catch (Exception e) when (!IsDisposed)
            {
                _result.ForeColor = Theme.ErrRed;
                _result.Text = Loc.T("Chưa gửi được (", "Could not send (") + e.Message + Loc.T("). Kiểm tra Internet rồi thử lại.", "). Check the Internet connection and try again.");
                _send.Enabled = true;
            }
        }
    }

    static string Trim(string s, int max) => s.Length > max ? s[..max] : s;
}
