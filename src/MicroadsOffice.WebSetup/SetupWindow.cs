using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MicroadsOffice.WebSetup
{
    /// <summary>Small window: connecting → downloading (progress, speed) → checking the signature → setup started.</summary>
    internal sealed class SetupWindow : Form
    {
        private enum Stage { Connecting, Downloading, Verifying, Launching, Error }

        private static readonly Color Red = Color.FromArgb(232, 49, 47);
        private static readonly Color Green = Color.FromArgb(43, 168, 74);
        private static readonly Color Ink = Color.FromArgb(23, 38, 28);
        private static readonly Color Mute = Color.FromArgb(101, 112, 101);
        private static readonly Color ErrorInk = Color.FromArgb(176, 28, 28);

        private readonly Func<string> _setupArgs;
        private readonly Label _title, _subtitle, _status, _detail, _note;
        private readonly ProgressBar _bar;
        private readonly Button _lang, _website, _retry, _full, _cancel;
        private readonly PictureBox _logo;
        private Stage _stage = Stage.Connecting;
        private ReleaseInfo _release;
        private long _done, _total;
        private double _speed;
        private SetupFailure _failure;
        private CancellationTokenSource _cts;
        private string _work;
        private bool _setupRunning;

        public int ExitCode { get; private set; } = 2;

        public SetupWindow(Func<string> setupArgs)
        {
            _setupArgs = setupArgs;
            SuspendLayout();
            // Every control below sets its own Font: in .NET Framework WinForms a control that inherits the form's font is
            // scaled a second time under AutoScaleMode.Dpi (found by measuring the window at 125 %).
            Font = new Font("Segoe UI", 9.75f);
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            Name = "webSetupForm";
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(560, 318);
            BackColor = Color.White;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            _logo = new PictureBox { Name = "picLogo", Location = new Point(24, 22), Size = new Size(48, 48), SizeMode = PictureBoxSizeMode.Zoom, Image = LoadLogo() };
            _title = new Label { Name = "lblTitle", Text = "Microads Office", Location = new Point(84, 18), AutoSize = true, Font = new Font("Segoe UI Semibold", 16f), ForeColor = Ink };
            _subtitle = new Label { Name = "lblSubtitle", Location = new Point(87, 52), Size = new Size(340, 22), ForeColor = Mute, Font = new Font("Segoe UI", 9.75f) };
            _lang = LinkButton("btnLang", new Point(440, 10), new Size(104, 26));
            _lang.TextAlign = ContentAlignment.MiddleRight;
            _lang.Click += (s, e) => { Program.Lang = Program.Vi ? "en" : "vi"; ApplyTexts(); };

            var accent = new Panel { Name = "accent", Location = new Point(0, 88), Size = new Size(560, 4) };
            accent.Paint += (s, e) =>
            {
                using (var b = new LinearGradientBrush(accent.ClientRectangle, Red, Green, LinearGradientMode.Horizontal)) e.Graphics.FillRectangle(b, accent.ClientRectangle);
            };

            _status = new Label { Name = "lblStatus", Location = new Point(24, 108), Size = new Size(512, 26), Font = new Font("Segoe UI Semibold", 11f), ForeColor = Ink, AutoEllipsis = true };
            _bar = new ProgressBar { Name = "progressDownload", Location = new Point(24, 138), Size = new Size(512, 16), Style = ProgressBarStyle.Marquee, MarqueeAnimationSpeed = 30, Maximum = 1000 };
            _detail = new Label { Name = "lblDetail", Location = new Point(24, 160), Size = new Size(512, 40), ForeColor = Mute, Font = new Font("Segoe UI", 9.75f) };
            _note = new Label { Name = "lblNote", Location = new Point(24, 204), Size = new Size(512, 40), ForeColor = Mute, Font = new Font("Segoe UI", 8.75f) };

            var bottom = new Panel { Name = "bottomBar", Location = new Point(0, 258), Size = new Size(560, 60), BackColor = Color.FromArgb(244, 245, 242) };
            _website = LinkButton("btnWebsite", new Point(16, 17), new Size(220, 26));
            _website.TextAlign = ContentAlignment.MiddleLeft;
            _website.Click += (s, e) => Engine.OpenUrl(Engine.Website);
            _retry = PushButton("btnRetry", true);
            _full = PushButton("btnFullInstaller", false);
            _cancel = PushButton("btnCancel", false);
            _retry.Click += (s, e) => Start();
            _full.Click += (s, e) => Engine.OpenUrl(Engine.PageUrl);
            _cancel.Click += (s, e) => Close();
            bottom.Controls.AddRange(new Control[] { _website, _retry, _full, _cancel });

            Controls.AddRange(new Control[] { _logo, _title, _subtitle, _lang, accent, _status, _bar, _detail, _note, bottom });
            ResumeLayout(false);
            ApplyTexts();
        }

        /// <summary>Largest frame of the embedded suite icon (PNG frames: .NET Framework's Icon class cannot decode them).</summary>
        private static Image LoadLogo()
        {
            try
            {
                byte[] ico;
                using (var s = typeof(SetupWindow).Assembly.GetManifestResourceStream("office.ico"))
                using (var ms = new MemoryStream()) { s.CopyTo(ms); ico = ms.ToArray(); }
                int count = BitConverter.ToUInt16(ico, 4), best = -1, bestW = 0;
                for (int i = 0; i < count; i++)
                {
                    int w = ico[6 + 16 * i] == 0 ? 256 : ico[6 + 16 * i];
                    if (w > bestW) { bestW = w; best = i; }
                }
                int len = BitConverter.ToInt32(ico, 6 + 16 * best + 8), off = BitConverter.ToInt32(ico, 6 + 16 * best + 12);
                if (ico[off] == 0x89) return new Bitmap(new MemoryStream(ico, off, len));
                using (var icon = new Icon(new MemoryStream(ico), bestW, bestW)) return icon.ToBitmap();
            }
            catch { return null; }
        }

        private static Button LinkButton(string name, Point at, Size size)
        {
            var b = new Button
            {
                Name = name, Location = at, Size = size, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                ForeColor = Color.FromArgb(196, 32, 42), BackColor = Color.Transparent, Font = new Font("Segoe UI", 9f, FontStyle.Underline), TabStop = true
            };
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(240, 240, 236);
            b.FlatAppearance.MouseDownBackColor = Color.FromArgb(230, 230, 226);
            return b;
        }

        private static Button PushButton(string name, bool primary)
        {
            var b = new Button
            {
                Name = name, Size = new Size(110, 32), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                BackColor = primary ? Red : Color.White, ForeColor = primary ? Color.White : Ink,
                Font = primary ? new Font("Segoe UI Semibold", 9.75f) : new Font("Segoe UI", 9.75f)
            };
            b.FlatAppearance.BorderSize = primary ? 0 : 1;
            b.FlatAppearance.BorderColor = Color.FromArgb(200, 204, 198);
            return b;
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            LayoutButtons(); // sizes are final now (after automatic DPI scaling)
            Start();
        }

        // ------------------------------------------------------------------ texts (re-applied when the language changes)

        private void ApplyTexts()
        {
            var T = (Func<string, string, string>)Program.T;
            Text = T("Cài đặt Microads Office", "Microads Office Setup");
            _subtitle.Text = T("Word · Excel · PDF — luôn luôn miễn phí", "Word · Excel · PDF — always free");
            _lang.Text = Program.Vi ? "English" : "Tiếng Việt";
            _website.Text = T("Công ty TNHH Công nghệ Microads", "Microads Technology Co., Ltd.") + " · microads.vn";
            _note.Text = T("Trình cài đặt tải phiên bản mới nhất từ microads.vn và chỉ chạy bộ cài khi tệp có chữ ký số hợp lệ của Microads.",
                "This installer downloads the latest version from microads.vn and only runs it when the file carries a valid Microads digital signature.");
            _retry.Text = T("Thử lại", "Try again");
            _full.Text = T("Tải bộ cài đầy đủ", "Full installer");
            _cancel.Text = _stage == Stage.Error ? T("Đóng", "Close") : T("Hủy", "Cancel");
            _status.ForeColor = _stage == Stage.Error ? ErrorInk : Ink;
            var ver = _release?.Version ?? "";
            switch (_stage)
            {
                case Stage.Connecting:
                    _status.Text = T("Đang kết nối tới microads.vn…", "Connecting to microads.vn…");
                    _detail.Text = T("Đang tìm phiên bản mới nhất.", "Looking for the latest version.");
                    break;
                case Stage.Downloading:
                    _status.Text = T("Đang tải Microads Office " + ver + "…", "Downloading Microads Office " + ver + "…");
                    _detail.Text = ProgressText();
                    break;
                case Stage.Verifying:
                    _status.Text = T("Đang kiểm tra chữ ký số của Microads…", "Checking the Microads digital signature…");
                    _detail.Text = T("Tệp chỉ được chạy khi chữ ký hợp lệ.", "The file only runs if the signature is valid.");
                    break;
                case Stage.Launching:
                    _status.Text = T("Đang mở trình cài đặt Microads Office " + ver + "…", "Starting Microads Office " + ver + " setup…");
                    _detail.Text = T("Chữ ký số hợp lệ ✓", "Valid digital signature ✓");
                    break;
                case Stage.Error:
                    _status.Text = ErrorTitle(_failure);
                    _detail.Text = ErrorText(_failure);
                    break;
            }
            LayoutButtons();
        }

        private string ProgressText()
        {
            var ci = Program.Vi ? new CultureInfo("vi-VN") : CultureInfo.InvariantCulture;
            string Mb(long b) => (b / 1048576.0).ToString("0.0", ci);
            var s = Mb(_done) + " / " + Mb(_total) + " MB";
            if (_speed > 0)
            {
                s += " · " + (_speed / 1048576.0).ToString("0.0", ci) + Program.T(" MB/giây", " MB/s");
                var left = (_total - _done) / _speed;
                if (left >= 1) s += " · " + Program.T("còn khoảng ", "about ") + (left >= 60 ? Math.Ceiling(left / 60) + Program.T(" phút", " min left") : Math.Ceiling(left) + Program.T(" giây", " s left"));
            }
            return s;
        }

        private static string ErrorTitle(SetupFailure f)
        {
            switch (f?.Kind)
            {
                case FailureKind.Network: return Program.T("Không tải được Microads Office", "Could not download Microads Office");
                case FailureKind.NoRelease: return Program.T("Chưa có phiên bản để tải", "No version available yet");
                case FailureKind.Signature: return Program.T("Tệp tải về không hợp lệ — đã xóa để bảo vệ máy tính", "Invalid download — deleted to protect your computer");
                case FailureKind.Disk: return Program.T("Không ghi được tệp tạm", "Could not write the temporary file");
                case FailureKind.Launch: return Program.T("Không mở được trình cài đặt", "Could not start the setup");
                default: return Program.T("Đã xảy ra lỗi", "Something went wrong");
            }
        }

        private static string ErrorText(SetupFailure f)
        {
            string tip;
            switch (f?.Kind)
            {
                case FailureKind.Network: tip = Program.T("Hãy kiểm tra kết nối Internet rồi bấm Thử lại, hoặc tải bộ cài đầy đủ.", "Check your Internet connection and click Try again, or get the full installer."); break;
                case FailureKind.NoRelease: tip = Program.T("Vui lòng thử lại sau ít phút.", "Please try again in a few minutes."); break;
                case FailureKind.Signature: tip = Program.T("Tệp không có chữ ký số hợp lệ của Microads nên không được chạy. Bấm Thử lại để tải lại.", "The file does not carry a valid Microads signature and was not run. Click Try again to download it again."); break;
                case FailureKind.Disk: tip = Program.T("Hãy giải phóng dung lượng ổ đĩa rồi bấm Thử lại.", "Free some disk space and click Try again."); break;
                default: tip = Program.T("Bấm Thử lại, hoặc tải bộ cài đầy đủ tại microads.vn/office.", "Click Try again, or get the full installer from microads.vn/office."); break;
            }
            var msg = f?.Message;
            return string.IsNullOrEmpty(msg) ? tip : tip + "\n(" + (msg.Length > 110 ? msg.Substring(0, 110) + "…" : msg) + ")";
        }

        private void LayoutButtons()
        {
            bool err = _stage == Stage.Error;
            _retry.Visible = err;
            _full.Visible = err;
            _cancel.Visible = true;
            _cancel.Enabled = _stage != Stage.Launching;
            // Works in the current coordinates: before automatic DPI scaling the bar is 560 wide, afterwards 560 × DPI/96.
            float k = _cancel.Parent.ClientSize.Width / 560f;
            int scale(int v) => (int)Math.Round(v * k);
            int x = _cancel.Parent.ClientSize.Width - scale(16);
            foreach (var b in new[] { _cancel, _full, _retry })
            {
                if (!b.Visible) continue;
                b.Width = Math.Max(scale(96), TextRenderer.MeasureText(b.Text, b.Font).Width + scale(28));
                x -= b.Width;
                b.Location = new Point(x, (_cancel.Parent.ClientSize.Height - b.Height) / 2);
                x -= scale(8);
            }
            _website.Visible = !err;
            AcceptButton = err ? _retry : null;
            CancelButton = _cancel;
        }

        private void SetStage(Stage s)
        {
            _stage = s;
            _bar.Style = s == Stage.Downloading ? ProgressBarStyle.Continuous : ProgressBarStyle.Marquee;
            if (s == Stage.Error) { _bar.Style = ProgressBarStyle.Continuous; _bar.Value = 0; }
            ApplyTexts();
        }

        // ------------------------------------------------------------------ work

        private async void Start()
        {
            if (_cts != null) return; // already running
            _failure = null;
            _release = null;
            _done = _total = 0;
            _speed = 0;
            _cts = new CancellationTokenSource();
            var ct = _cts.Token;
            Engine.Cleanup(_work);
            _work = null;
            SetStage(Stage.Connecting);
            try
            {
                _release = await Task.Run(() => Engine.FetchLatest(ct));
                _work = Engine.CreateWorkFolder();
                _total = _release.Size;
                SetStage(Stage.Downloading);
                var sw = Stopwatch.StartNew();
                long lastUi = -1000, lastBytes = 0;
                double lastT = 0;
                var r = _release;
                var work = _work;
                var path = await Task.Run(() => Engine.Download(r, work, (done, total) =>
                {
                    var now = sw.ElapsedMilliseconds;
                    if (now - lastUi < 120 && done < total) return;
                    lastUi = now;
                    var t = now / 1000.0;
                    if (t - lastT >= 0.5) { var inst = (done - lastBytes) / (t - lastT); _speed = _speed <= 0 ? inst : _speed * 0.7 + inst * 0.3; lastBytes = done; lastT = t; }
                    try { BeginInvoke((Action)(() => OnProgress(done, total))); } catch { }
                }, ct));
                ct.ThrowIfCancellationRequested();
                SetStage(Stage.Verifying);
                var locked = await Task.Run(() => Engine.VerifyAndLock(path, r, ct));
                Process proc;
                try
                {
                    // The last cancellation check and the start run on the UI thread, back to back: a click on Hủy / ✕
                    // (handled on the same thread) is either seen here or comes after the setup is marked as running.
                    ct.ThrowIfCancellationRequested();
                    _setupRunning = true;
                    proc = Engine.StartSetup(path, _setupArgs());
                }
                catch { _setupRunning = false; throw; }
                finally { locked.Dispose(); }
                SetStage(Stage.Launching);
                Engine.Log("started setup " + r.Version);
                await Task.Run(() => WaitForWindow(proc));
                if (proc.HasExited && proc.ExitCode != 0)
                {
                    _setupRunning = false;
                    throw new SetupFailure(FailureKind.Launch, "setup exited with code " + proc.ExitCode);
                }
                Hide();
                await Task.Run(() => proc.WaitForExit());
                ExitCode = proc.ExitCode == 0 ? 0 : 1;
                await Task.Run(() => Engine.Cleanup(work));
                _work = null;
                _setupRunning = false;
                _cts = null;
                Close();
                return;
            }
            catch (Exception ex) when (ct.IsCancellationRequested)
            {
                Engine.Log("cancelled: " + ex.GetType().Name);
                ExitCode = 2;
            }
            catch (SetupFailure f) { Fail(f); }
            catch (Exception ex) { Fail(new SetupFailure(ex is IOException || ex is System.Net.WebException ? FailureKind.Network : FailureKind.Other, ex.Message, ex)); }
            _cts = null;
            if (ct.IsCancellationRequested) { Engine.Cleanup(_work); _work = null; }
        }

        private void Fail(SetupFailure f)
        {
            Engine.Log("failed (" + f.Kind + "): " + f.Message);
            _failure = f;
            ExitCode = f.Kind == FailureKind.Network || f.Kind == FailureKind.NoRelease ? 4 : f.Kind == FailureKind.Signature ? 5 : 6;
            Engine.Cleanup(_work);
            _work = null;
            if (!IsDisposed) { SetStage(Stage.Error); if (!Visible) Show(); Activate(); }
        }

        private void OnProgress(long done, long total)
        {
            if (_stage != Stage.Downloading || IsDisposed) return;
            _done = done; _total = total > 0 ? total : _total;
            _bar.Value = _total > 0 ? (int)Math.Min(1000, done * 1000 / _total) : 0;
            _detail.Text = ProgressText();
        }

        private static void WaitForWindow(Process p)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < 30000)
            {
                try { p.Refresh(); if (p.HasExited || p.MainWindowHandle != IntPtr.Zero) return; } catch { return; }
                Thread.Sleep(200);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_setupRunning && e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; return; } // the setup window is open
            if (_cts != null)
            {
                try { _cts.Cancel(); } catch { }
                ExitCode = 2;
            }
            base.OnFormClosing(e);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            if (!_setupRunning)
            {
                try { Hide(); } catch { } // the window disappears at once; the cancelled download stops and is deleted below
                Engine.Cleanup(_work);
            }
        }
    }
}
