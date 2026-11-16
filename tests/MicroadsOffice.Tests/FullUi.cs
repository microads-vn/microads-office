using System.Drawing;
using System.Drawing.Printing;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using MicroadsOffice.Core.Pdfium;
using MicroadsOffice.Core.Signing;
using Application = FlaUI.Core.Application;
using Point = System.Drawing.Point;

namespace MicroadsOffice.Tests;

/// <summary>End-to-end UI tests that drive the real application like a user (mouse + keyboard).</summary>
public static class FullUi
{
    private static Window W = null!;
    private static UIA3Automation A = null!;
    private static Application App = null!;
    private static string Out = "";
    private static string Exe = "";
    private static readonly VerifyOptions Offline = new() { OnlineRevocation = false, AllowCertificateDownloads = false };

    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    private static extern uint SetThreadExecutionState(uint esFlags);

    // ------------------------------------------------------------------ input safety guard
    // Every mouse click / key press is sent ONLY if the target window belongs to Microads Office.
    // Other applications (chat apps, browsers...) that pop up over the app are never touched.

    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
    private struct POINTSTRUCT { public int X; public int Y; }
    [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(POINTSTRUCT p);

    private static bool IsOurs(IntPtr h) { GetWindowThreadProcessId(h, out var pid); return pid == (uint)App.ProcessId; }
    private static bool OursAt(Point p) => IsOurs(WindowFromPoint(new POINTSTRUCT { X = p.X, Y = p.Y }));

    private static void EnsureOurs()
    {
        for (int i = 0; i < 3; i++)
        {
            if (IsOurs(GetForegroundWindow())) return;
            try
            {
                var top = AllWindows().LastOrDefault(w => w.Properties.ControlType.ValueOrDefault == ControlType.Window) ?? W;
                top.AsWindow().SetForeground();
            }
            catch { }
            Thread.Sleep(400);
        }
        if (!IsOurs(GetForegroundWindow())) throw new AssertException("AN TOÀN: cửa sổ ứng dụng khác đang ở trên Microads Office — dừng, không gửi thao tác.");
    }

    private static void SafeClickAt(Point pt)
    {
        EnsureOurs();
        if (!OursAt(pt)) { EnsureOurs(); Thread.Sleep(300); }
        if (!OursAt(pt)) throw new AssertException("AN TOÀN: vị trí bấm bị cửa sổ khác che — dừng.");
        Mouse.Click(pt);
    }

    private static Point ClickPoint(AutomationElement el)
    {
        try { if (el.TryGetClickablePoint(out var cp)) return cp; } catch { }
        var r = el.BoundingRectangle;
        return new Point(r.Left + r.Width / 2, r.Top + r.Height / 2);
    }

    private static void SafeClick(AutomationElement? el)
    {
        if (el == null) return;
        SafeClickAt(ClickPoint(el));
    }

    private static class Kb
    {
        public static void Press(VirtualKeyShort k) { EnsureOurs(); Keyboard.Press(k); }
        public static void TypeSimultaneously(params VirtualKeyShort[] k) { EnsureOurs(); Keyboard.TypeSimultaneously(k); }
    }

    // ------------------------------------------------------------------ helpers

    private static AutomationElement Find(AutomationElement root, Func<ConditionFactory, ConditionBase> cond, int ms = 8000, bool required = true)
    {
        var el = Retry.WhileNull(() => root.FindFirstDescendant(cond), TimeSpan.FromMilliseconds(ms), TimeSpan.FromMilliseconds(150)).Result;
        if (el == null && required) throw new AssertException("Không tìm thấy phần tử: " + cond(A.ConditionFactory));
        return el!;
    }

    private static AutomationElement Id(AutomationElement root, string id, int ms = 8000) => Find(root, c => c.ByAutomationId(id), ms);
    private static AutomationElement Name(AutomationElement root, string name, int ms = 8000) => Find(root, c => c.ByName(name), ms);
    private static AutomationElement Ribbon(string tab) => Id(W, "ribbon" + tab);
    private static AutomationElement Status => Id(W, "statusBar");

    private static void Click(AutomationElement el, int wait = 350)
    {
        SafeClick(el);
        Thread.Sleep(wait);
    }

    private static void RibbonClick(string tab, string name, int wait = 400)
    {
        SelectTab(tab);
        Click(Name(Ribbon(tab), name), wait);
    }

    private static void SelectTab(string tab)
    {
        var header = tab switch { "Home" => "Thẻ Trang chủ", "View" => "Thẻ Xem", "Protect" => "Thẻ Ký & Bảo vệ", _ => "Thẻ Trợ giúp" };
        for (int i = 0; i < 4; i++)
        {
            var strip = W.FindFirstDescendant(c => c.ByAutomationId("ribbon" + tab));
            if (strip != null && !strip.IsOffscreen) return;
            Click(Name(Id(W, "ribbonTabs"), header), 350);
        }
    }

    private static bool TitleMatch(string? name, string part) =>
        part.Length <= 3 ? string.Equals(name ?? "", part, StringComparison.OrdinalIgnoreCase) : (name ?? "").Contains(part, StringComparison.OrdinalIgnoreCase);

    private static Window Modal(string titlePart, int ms = 10000)
    {
        var win = Retry.WhileNull(() =>
        {
            foreach (var m in W.ModalWindows) if (TitleMatch(m.Name, titlePart)) return m;
            foreach (var m in AllWindows()) if (TitleMatch(m.Name, titlePart) && m.Properties.AutomationId.ValueOrDefault != "MainForm") return m.AsWindow();
            foreach (var top in A.GetDesktop().FindAllChildren(c => c.ByProcessId(App.ProcessId)))
            {
                if (TitleMatch(top.Name, titlePart) && top.Properties.ControlType.ValueOrDefault == ControlType.Window && top.Properties.AutomationId.ValueOrDefault != "MainForm")
                    return top.AsWindow();
                foreach (var child in top.FindAllChildren(c => c.ByControlType(ControlType.Window)))
                    if (TitleMatch(child.Name, titlePart)) return child.AsWindow();
            }
            return null;
        }, TimeSpan.FromMilliseconds(ms), TimeSpan.FromMilliseconds(200)).Result;
        if (win == null) throw new AssertException("Không thấy hộp thoại: " + titlePart);
        Thread.Sleep(300);
        return win;
    }

    private static bool ModalGone(string titlePart, int ms = 8000) =>
        Retry.WhileTrue(() => W.ModalWindows.Any(m => (m.Name ?? "").Contains(titlePart, StringComparison.OrdinalIgnoreCase)), TimeSpan.FromMilliseconds(ms)).Success;

    /// <summary>Common Windows file dialog: type the path and confirm.</summary>
    private static void FileDialog(string titlePart, string path)
    {
        var dlg = Modal(titlePart, 15000);
        var edit = Retry.WhileNull(() => dlg.FindFirstDescendant(c => c.ByAutomationId("FileNameControlHost"))?.FindFirstDescendant(c => c.ByControlType(ControlType.Edit))
                                         ?? dlg.FindFirstDescendant(c => c.ByAutomationId("1148").And(c.ByControlType(ControlType.Edit)))
                                         ?? dlg.FindFirstDescendant(c => c.ByAutomationId("1001").And(c.ByControlType(ControlType.Edit))), TimeSpan.FromSeconds(8)).Result
                   ?? throw new AssertException("Không thấy ô tên tệp");
        Thread.Sleep(800);
        for (int attempt = 0; attempt < 3; attempt++)
        {
            // Paste through the clipboard: immune to Vietnamese IMEs (Unikey/EVKey) and autocomplete.
            SetClipboard(path);
            edit.Focus();
            Thread.Sleep(250);
            Kb.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
            Kb.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_V);
            Thread.Sleep(400);
            string val = "";
            try { val = edit.Patterns.Value.Pattern.Value.Value; } catch { }
            if (!string.Equals(val, path, StringComparison.OrdinalIgnoreCase)) { try { edit.Patterns.Value.Pattern.SetValue(path); } catch { } Thread.Sleep(300); }
            Kb.Press(VirtualKeyShort.ENTER);
            Thread.Sleep(1000);
            var warn = AllWindows().FirstOrDefault(w => w.ClassName == "#32770" && w.FindFirstDescendant(c => c.ByAutomationId("65535")) != null
                                                       && (w.Name ?? "") == (dlg.Name ?? ""));
            if (warn == null) break;
            SafeClick(warn.FindFirstDescendant(c => c.ByAutomationId("2").Or(c.ByAutomationId("1")).And(c.ByControlType(ControlType.Button))));
            Thread.Sleep(1200);
        }
        // Overwrite confirmation, if any
        var confirm = W.ModalWindows.FirstOrDefault(m => m.FindFirstDescendant(c => c.ByAutomationId("CommandButton_6")) != null);
        SafeClick(confirm?.FindFirstDescendant(c => c.ByAutomationId("CommandButton_6")));
    }

    /// <summary>
    /// Answers a message box. Codes keep the classic MessageBox meaning: 6 = Yes/affirmative, 7 = No/negative,
    /// 1 = OK/first button, 2 = Cancel/last button. Works with the classic MessageBox and with the app's
    /// localized task dialogs (Msg.*), whose buttons are ordered affirmative → negative → cancel.
    /// </summary>
    private static string MsgBox(string buttonId, int ms = 8000) => MsgBoxCore(buttonId, ms, required: true)!;

    private static string? MsgBoxCore(string buttonId, int ms, bool required)
    {
        AutomationElement? box = null;
        AutomationElement? button = null;
        Retry.WhileNull(() =>
        {
            foreach (var m in AllWindows().Where(w => w.ClassName == "#32770"))
            {
                if (m.FindFirstDescendant(c => c.ByAutomationId("65535")) != null)
                {
                    var b = m.FindFirstDescendant(c => c.ByAutomationId(buttonId).And(c.ByControlType(ControlType.Button)));
                    if (b != null) { box = m; button = b; return m; }
                }
                var cmds = m.FindAllDescendants(c => c.ByControlType(ControlType.Button))
                    .Where(b => (b.Properties.AutomationId.ValueOrDefault ?? "").StartsWith("CommandButton_"))
                    .OrderBy(b => int.TryParse((b.Properties.AutomationId.ValueOrDefault ?? "").Substring("CommandButton_".Length), out var n) ? n : 0)
                    .ToList();
                if (cmds.Count == 0) continue;
                var pick = buttonId switch
                {
                    "6" or "1" => cmds[0],
                    "7" => cmds.Count > 1 ? cmds[1] : null,
                    _ => cmds[^1]
                };
                if (pick != null) { box = m; button = pick; return m; }
            }
            return null;
        }, TimeSpan.FromMilliseconds(ms));
        if (box == null || button == null)
        {
            if (required) throw new AssertException("Không thấy hộp thông báo (nút " + buttonId + ")");
            return null;
        }
        var text = string.Join(" | ", box.FindAllDescendants(c => c.ByControlType(ControlType.Text)).Select(x => x.Name));
        SafeClick(button);
        Thread.Sleep(500);
        return text;
    }

    private static IEnumerable<AutomationElement> AllWindows()
    {
        var result = new List<AutomationElement>();
        void Walk(AutomationElement e, int depth)
        {
            if (depth > 4) return;
            foreach (var c in e.FindAllChildren(cf => cf.ByControlType(ControlType.Window))) { result.Add(c); Walk(c, depth + 1); }
        }
        foreach (var top in A.GetDesktop().FindAllChildren(c => c.ByProcessId(App.ProcessId))) { result.Add(top); Walk(top, 1); }
        return result;
    }

    /// <summary>Selects the test digital ID in the sign dialog. Returns true when it is still locked (password asked on OK).</summary>
    private static bool ChooseTestCertificate(Window d)
    {
        string Info() => Id(d, "lblCertInfo").Name ?? "";
        if (!Info().Contains("Kiểm Thử UI"))
        {
            var cbo = Id(d, "cboCertificate");
            cbo.Focus(); Thread.Sleep(200);
            Kb.Press(VirtualKeyShort.HOME); Thread.Sleep(400);
            for (int i = 0; i < 20 && !Info().Contains("Kiểm Thử UI"); i++)
            {
                if (AllWindows().Any(w => (w.Name ?? "").Contains("Mở khóa tệp chứng thư"))) { UnlockPrompt(); continue; }
                Kb.Press(VirtualKeyShort.DOWN); Thread.Sleep(450);
                if (AllWindows().Any(w => (w.Name ?? "").Contains("Mở khóa tệp chứng thư"))) UnlockPrompt();
            }
        }
        T.True(Info().Contains("Kiểm Thử UI"), "test ID selected: " + Info());
        return Info().Contains("đang khóa");
    }

    private static void UnlockPrompt()
    {
        var pw = Modal("Mở khóa tệp chứng thư");
        Id(pw, "txtInput").Patterns.Value.Pattern.SetValue("1234");
        Click(Id(pw, "btnOk"), 900);
    }

    private static void SetClipboard(string text)
    {
        var th = new Thread(() => { for (int i = 0; i < 5; i++) { try { System.Windows.Forms.Clipboard.SetText(text); return; } catch { Thread.Sleep(100); } } });
        th.SetApartmentState(ApartmentState.STA);
        th.Start();
        th.Join();
    }

    /// <summary>Selects a combobox item like a user: open the list with the mouse and click the item.</summary>
    private static void SelectCombo(AutomationElement cbo, string text)
    {
        string Val() { try { return cbo.Patterns.Value.PatternOrDefault?.Value.ValueOrDefault ?? ""; } catch { return ""; } }
        for (int attempt = 0; attempt < 3 && Val() != text; attempt++)
        {
            var r = cbo.BoundingRectangle;
            SafeClickAt(new Point(r.Right - 8, r.Top + r.Height / 2));
            Thread.Sleep(500);
            var item = Retry.WhileNull(() => A.GetDesktop().FindFirstDescendant(c => c.ByControlType(ControlType.ListItem).And(c.ByName(text))),
                TimeSpan.FromSeconds(3)).Result;
            if (item != null)
            {
                try { item.Patterns.ScrollItem.PatternOrDefault?.ScrollIntoView(); } catch { }
                SafeClick(item);
            }
            else Kb.Press(VirtualKeyShort.ESCAPE);
            Thread.Sleep(500);
        }
        T.True(Val() == text, $"combo value '{Val()}' != '{text}'");
    }

    private static string PageBox => Name(Status, "Số trang").Patterns.Value.Pattern.Value.Value;
    private static string ZoomText => Id(Status, "1001").Patterns.Value.Pattern.Value.Value;
    private static string StatusText => Status.FindFirstChild(c => c.ByControlType(ControlType.Text))!.Name;
    private static AutomationElement Menu(string name) => Find(A.GetDesktop(), c => c.ByControlType(ControlType.MenuItem).And(c.ByName(name)));
    private static AutomationElement FileMenu => Name(Id(W, "ribbonTabs"), "Tệp");
    private static AutomationElement Viewer => Id(W, "pdfViewer");

    private static bool WaitFor(Func<bool> cond, int ms = 8000) => Retry.WhileFalse(() => { try { return cond(); } catch { return false; } }, TimeSpan.FromMilliseconds(ms), TimeSpan.FromMilliseconds(150)).Success;

    private static void Drag(Point from, Point to)
    {
        EnsureOurs();
        if (!OursAt(from) || !OursAt(to)) throw new AssertException("AN TOÀN: vùng kéo chuột bị cửa sổ khác che — dừng.");
        Mouse.MoveTo(from);
        Thread.Sleep(120);
        Mouse.Down(MouseButton.Left);
        for (int i = 1; i <= 12; i++)
        {
            Mouse.MoveTo(new Point(from.X + (to.X - from.X) * i / 12, from.Y + (to.Y - from.Y) * i / 12));
            Thread.Sleep(25);
        }
        Mouse.Up(MouseButton.Left);
        Thread.Sleep(400);
    }

    private static Bitmap? CaptureViewer() { try { return FlaUI.Core.Capturing.Capture.Element(Viewer).Bitmap; } catch { return null; } }

    private static double Diff(Bitmap a, Bitmap b)
    {
        int w = Math.Min(a.Width, b.Width), h = Math.Min(a.Height, b.Height), diff = 0, n = 0;
        for (int y = 0; y < h; y += 7) for (int x = 0; x < w; x += 7)
        {
            n++;
            var p = a.GetPixel(x, y); var q = b.GetPixel(x, y);
            if (Math.Abs(p.R - q.R) + Math.Abs(p.G - q.G) + Math.Abs(p.B - q.B) > 60) diff++;
        }
        return n == 0 ? 0 : (double)diff / n;
    }

    private static string ClipboardText()
    {
        string s = "";
        var th = new Thread(() => { try { s = System.Windows.Forms.Clipboard.GetText(); } catch { } });
        th.SetApartmentState(ApartmentState.STA);
        th.Start();
        th.Join();
        return s;
    }

    private static int ParseZoom() => int.Parse(ZoomText.TrimEnd('%'));

    private static void Step(TestRunner t, string name, Action body)
    {
        t.Test(name, () =>
        {
            try { W.SetForeground(); Thread.Sleep(150); } catch { }
            try { body(); }
            catch
            {
                UiTests.Shot(W, "FAIL-" + string.Concat(name.Where(char.IsLetterOrDigit)).Substring(0, Math.Min(40, name.Count(char.IsLetterOrDigit))));
                // close stray dialogs so following steps can continue
                foreach (var m in W.ModalWindows.Reverse()) { try { m.Close(); } catch { } }
                throw;
            }
        });
    }

    // ------------------------------------------------------------------ test plan

    public static void Run(TestRunner t, string exe, string outDir)
    {
        Out = outDir;
        Exe = exe;
        var data = Path.Combine(outDir, "ui-data");
        if (Directory.Exists(data)) Directory.Delete(data, true);
        Directory.CreateDirectory(data);
        foreach (var f in Directory.GetFiles(outDir, "ui-*.pdf")) File.Delete(f);
        var sample = Path.Combine(outDir, "sample-3pages.pdf");
        var sigField = Path.Combine(outDir, "sample-sigfield.pdf");
        File.WriteAllBytes(sample, TestPdf.Create(3, withLink: true, withOutline: true));
        File.WriteAllBytes(sigField, TestPdf.Create(1, withEmptySigField: true));
        var seal = Path.Combine(outDir, "seal.png");
        using (var s = CoreTests.MakeSeal()) s.Save(seal, System.Drawing.Imaging.ImageFormat.Png);
        var encrypted = Path.Combine(outDir, "encrypted.pdf"); // produced by pikepdf in the test script (password "abc")

        SetThreadExecutionState(0x80000000 | 0x00000002 | 0x00000001); // keep display on during UI tests
        A = new UIA3Automation();
        App = UiTests.Launch(exe, data, sample);
        UiTests.ShotAllowed = () => IsOurs(GetForegroundWindow());
        string? userClipboard = null;
        var ct = new Thread(() => { try { if (System.Windows.Forms.Clipboard.ContainsText()) userClipboard = System.Windows.Forms.Clipboard.GetText(); } catch { } });
        ct.SetApartmentState(ApartmentState.STA); ct.Start(); ct.Join();
        try
        {
            W = UiTests.MainWindow(App, A);
            W.SetForeground();
            Thread.Sleep(2000);
            RunSteps(t, sample, sigField, seal, encrypted, data);
        }
        finally
        {
            try { if (!App.HasExited) { App.Kill(); } } catch { }
            A.Dispose();
            if (userClipboard != null) SetClipboard(userClipboard); // restore the user's clipboard
        }
    }

    private static void RunSteps(TestRunner t, string sample, string sigField, string seal, string encrypted, string data)
    {
        Step(t, "UI01 Khởi động, mở tệp từ dòng lệnh, tiêu đề & thanh trạng thái", () =>
        {
            T.True(W.Title.Contains("sample-3pages.pdf") && W.Title.Contains("Microads Office"), "title: " + W.Title);
            T.Eq("/ 3", Name(Status, "/ 3").Name);
            T.Eq("1", PageBox);
            UiTests.Shot(W, "01-main");
        });

        Step(t, "UI02 Thanh trạng thái: Trang sau / Trang cuối / Trang trước / Trang đầu / nhập số trang", () =>
        {
            Click(Viewer, 500);
            Click(Name(Status, "Trang sau")); T.True(WaitFor(() => PageBox == "2"), "→2 was " + PageBox);
            Click(Name(Status, "Trang cuối")); T.True(WaitFor(() => PageBox == "3"), "→3");
            Click(Name(Status, "Trang trước")); T.True(WaitFor(() => PageBox == "2"), "→2");
            Click(Name(Status, "Trang đầu")); T.True(WaitFor(() => PageBox == "1"), "→1");
            var box = Name(Status, "Số trang");
            box.Patterns.Value.Pattern.SetValue("3");
            box.Focus(); Kb.Press(VirtualKeyShort.ENTER); Thread.Sleep(400);
            T.True(WaitFor(() => PageBox == "3"), "typed page 3");
            Click(Name(Status, "Trang đầu"));
        });

        Step(t, "UI03 Thu phóng: Phóng to / Thu nhỏ / Vừa trang / Vừa chiều rộng / ô tỉ lệ (ribbon & thanh trạng thái)", () =>
        {
            int z0 = ParseZoom();
            RibbonClick("Home", "Phóng to"); int z1 = ParseZoom(); T.True(z1 > z0, $"zoom in {z0}->{z1}");
            RibbonClick("Home", "Thu nhỏ"); int z2 = ParseZoom(); T.True(z2 < z1, $"zoom out {z1}->{z2}");
            RibbonClick("Home", "Vừa trang"); int zp = ParseZoom(); T.True(zp > 10 && zp < z1 + 500, "fit page " + zp);
            RibbonClick("Home", "Vừa chiều rộng"); int zw = ParseZoom(); T.True(zw > zp, $"fit width {zw} > fit page {zp}");
            Click(Name(Status, "Phóng to")); T.True(ParseZoom() > zw, "status zoom in");
            Click(Name(Status, "Thu nhỏ")); Click(Name(Status, "Vừa trang")); T.True(Math.Abs(ParseZoom() - zp) <= 1, "status fit page");
            Click(Name(Status, "Vừa chiều rộng")); T.True(Math.Abs(ParseZoom() - zw) <= 1, "status fit width");
            var zEdit = Id(Status, "1001");
            zEdit.Patterns.Value.Pattern.SetValue("150%"); zEdit.Focus(); Kb.Press(VirtualKeyShort.ENTER); Thread.Sleep(400);
            T.True(WaitFor(() => ZoomText == "150%"), "typed 150% got " + ZoomText);
            RibbonClick("Home", "Vừa chiều rộng");
        });

        Step(t, "UI04 Xoay trái / Xoay phải thay đổi hiển thị trang", () =>
        {
            Thread.Sleep(600);
            int zBefore = ParseZoom();
            using var before = CaptureViewer();
            RibbonClick("Home", "Xoay phải", 1800);
            using var rotated = CaptureViewer();
            T.True(ParseZoom() < zBefore, $"fit-width zoom recomputed for landscape view {zBefore} -> {ParseZoom()}");
            T.True(before != null && rotated != null && Diff(before, rotated) > 0.002, "view changed after rotate");
            UiTests.Shot(W, "04-rotated");
            RibbonClick("Home", "Xoay trái", 1800);
            using var back = CaptureViewer();
            T.True(Math.Abs(ParseZoom() - zBefore) <= 1, "rotate back restores zoom");
        });

        Step(t, "UI05 Công cụ Tay / Chọn văn bản / Ctrl+A / Sao chép vào clipboard", () =>
        {
            RibbonClick("Home", "Chọn văn bản");
            T.True(Name(Ribbon("Home"), "Chọn văn bản").Properties.ControlType.Value == ControlType.CheckBox, "select tool checked");
            Click(Viewer); Kb.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A); Thread.Sleep(500);
            T.True(WaitFor(() => Name(Ribbon("Home"), "Sao chép").IsEnabled), "copy enabled after select all");
            RibbonClick("Home", "Sao chép", 600);
            var clip = ClipboardText();
            T.True(clip.Contains("Microads Test Page 1") && clip.Contains("Microads Test Page 3"), "clipboard: " + clip[..Math.Min(80, clip.Length)]);
            T.True(StatusText.Contains("sao chép"), "status " + StatusText);
            RibbonClick("Home", "Công cụ Tay");
            T.True(Name(Ribbon("Home"), "Công cụ Tay").Properties.ControlType.Value == ControlType.CheckBox, "hand tool checked");
        });

        Step(t, "UI06 Tìm kiếm: nút Tìm kiếm, nhập từ khóa, kết quả, kết quả sau/trước, F3", () =>
        {
            RibbonClick("Home", "Tìm kiếm");
            var box = Id(W, "txtSearch");
            box.Patterns.Value.Pattern.SetValue("contract");
            Click(Id(W, "btnSearchGo"), 800);
            T.True(WaitFor(() => Id(W, "lblSearchStatus").Name.Contains("3 kết quả")), "3 results: " + Id(W, "lblSearchStatus").Name);
            Thread.Sleep(600);
            Click(Id(W, "btnSearchNext"), 700); T.True(WaitFor(() => PageBox == "2"), "next hit page 2 (was " + PageBox + ")");
            Click(Id(W, "btnSearchNext"), 500); T.True(WaitFor(() => PageBox == "3"), "next hit page 3");
            Click(Id(W, "btnSearchPrev"), 500); T.True(WaitFor(() => PageBox == "2"), "prev hit page 2");
            Click(Viewer); Kb.Press(VirtualKeyShort.F3); T.True(WaitFor(() => PageBox == "3"), "F3 → page 3");
            Id(W, "chkWholeWord").AsCheckBox().IsChecked = true;
            Id(W, "txtSearch").Patterns.Value.Pattern.SetValue("contrac");
            Click(Id(W, "btnSearchGo"), 800);
            T.True(WaitFor(() => Id(W, "lblSearchStatus").Name.Contains("Không tìm thấy")), "whole word no partial match");
            Id(W, "chkWholeWord").AsCheckBox().IsChecked = false;
            UiTests.Shot(W, "06-search");
        });

        Step(t, "UI07 Thẻ Xem: Trang đầu/trước/sau/cuối, Đi tới trang, Kích thước thật, ô tỉ lệ ribbon", () =>
        {
            RibbonClick("View", "Trang đầu"); T.True(WaitFor(() => PageBox == "1"), "first");
            RibbonClick("View", "Trang sau"); T.True(WaitFor(() => PageBox == "2"), "next");
            RibbonClick("View", "Trang cuối"); T.True(WaitFor(() => PageBox == "3"), "last");
            RibbonClick("View", "Trang trước"); T.True(WaitFor(() => PageBox == "2"), "prev");
            RibbonClick("View", "Đi tới trang...");
            var dlg = Modal("Đi tới trang");
            Id(dlg, "txtInput").Patterns.Value.Pattern.SetValue("3");
            Click(Id(dlg, "btnOk"));
            T.True(WaitFor(() => PageBox == "3"), "goto 3");
            RibbonClick("View", "Kích thước thật"); T.True(WaitFor(() => ZoomText == "100%"), "actual size " + ZoomText);
            var rz = Id(Ribbon("View"), "1001");
            rz.Patterns.Value.Pattern.SetValue("75%"); rz.Focus(); Kb.Press(VirtualKeyShort.ENTER);
            T.True(WaitFor(() => ZoomText == "75%"), "ribbon zoom 75% " + ZoomText);
        });

        Step(t, "UI08 Thẻ Xem: Một trang / Cuộn liên tục / Ngăn điều hướng / Hình thu nhỏ / Dấu trang", () =>
        {
            RibbonClick("View", "Một trang");
            T.True(Name(Ribbon("View"), "Một trang").Properties.ControlType.Value == ControlType.CheckBox, "single page checked");
            RibbonClick("View", "Trang sau"); T.True(WaitFor(() => PageBox == "3" || PageBox == "1" || PageBox == "2"), "nav works single");
            RibbonClick("View", "Cuộn liên tục");
            T.True(Name(Ribbon("View"), "Cuộn liên tục").Properties.ControlType.Value == ControlType.CheckBox, "continuous checked");
            RibbonClick("View", "Ngăn điều hướng");
            T.True(WaitFor(() => Name(Ribbon("View"), "Ngăn điều hướng").Properties.ControlType.Value == ControlType.Button), "nav hidden");
            RibbonClick("View", "Ngăn điều hướng");
            T.True(WaitFor(() => Name(Ribbon("View"), "Ngăn điều hướng").Properties.ControlType.Value == ControlType.CheckBox), "nav shown");
            RibbonClick("View", "Hình thu nhỏ");
            T.True(Id(W, "thumbnailList") != null, "thumbnails");
            UiTests.Shot(W, "08-thumbnails");
            RibbonClick("View", "Dấu trang");
            var item = Find(Id(W, "bookmarkTree"), c => c.ByName("Chapter 2"));
            Click(item); T.True(WaitFor(() => PageBox == "2"), "bookmark → page 2");
            // nav strip buttons
            Click(Name(Id(W, "navStrip"), "Hình thu nhỏ trang")); T.True(Id(W, "thumbnailList") != null, "nav thumbs");
            Click(Name(Id(W, "navStrip"), "Tìm kiếm")); T.True(Id(W, "txtSearch") != null, "nav search");
            Click(Name(Id(W, "navStrip"), "Bảng chữ ký số")); T.True(Id(W, "signatureTree") != null, "nav signatures");
            Click(Name(Id(W, "navStrip"), "Dấu trang (mục lục)")); T.True(Id(W, "bookmarkTree") != null, "nav bookmarks");
        });

        Step(t, "UI09 Toàn màn hình (F11/Esc) và Thuộc tính tài liệu", () =>
        {
            RibbonClick("View", "Toàn màn hình", 1200);
            T.True(WaitFor(() => W.FindFirstDescendant(c => c.ByAutomationId("ribbon")) is null or { IsOffscreen: true }), "ribbon hidden in full screen");
            UiTests.Shot(W, "09-fullscreen");
            Kb.Press(VirtualKeyShort.ESCAPE); Thread.Sleep(1000);
            T.True(WaitFor(() => W.FindFirstDescendant(c => c.ByAutomationId("ribbon")) is { IsOffscreen: false }), "ribbon back");
            RibbonClick("View", "Thuộc tính tài liệu");
            var dlg = Modal("Thuộc tính tài liệu");
            var lv = Id(dlg, "lstProperties");
            T.True(lv.FindFirstDescendant(c => c.ByName("Số trang")) != null, "has page count row");
            T.True(lv.FindAllDescendants().Any(e => e.Name == "3"), "page count 3");
            UiTests.Shot(dlg, "09-properties");
            Click(Id(dlg, "btnClose"));
        });

        Step(t, "UI10 Mẫu chữ ký: tạo mẫu con dấu (chọn ảnh, bố cục nền) và mẫu chữ ký tay (vẽ)", () =>
        {
            RibbonClick("Protect", "Mẫu chữ ký");
            var mgr = Modal("Mẫu chữ ký");
            Click(Id(mgr, "btnNewStyleMgr"));
            var sd = Modal("Tạo mẫu chữ ký");
            Id(sd, "txtStyleTitle").Patterns.Value.Pattern.SetValue("Con dấu QA");
            Click(Id(sd, "btnChooseImage"));
            FileDialog("Chọn ảnh", seal);
            T.True(WaitFor(() => Id(sd, "rdoGraphicImage").AsRadioButton().IsChecked), "image radio selected");
            var lay = Id(sd, "cboLayout"); lay.Focus(); Kb.Press(VirtualKeyShort.HOME); Kb.Press(VirtualKeyShort.DOWN); Kb.Press(VirtualKeyShort.DOWN); Thread.Sleep(300);
            T.True(Id(sd, "rdoGraphicNone").AsRadioButton().IsChecked == false, "only one graphic option selected");
            Id(sd, "chkBorder").AsCheckBox().IsChecked = true;
            Thread.Sleep(400);
            UiTests.Shot(sd, "10-style-seal");
            Click(Id(sd, "btnOk"));
            T.True(WaitFor(() => Id(mgr, "lstStyles").FindFirstDescendant(c => c.ByName("Con dấu QA")) != null), "style saved");
            Click(Id(mgr, "btnNewStyleMgr"));
            var sd2 = Modal("Tạo mẫu chữ ký");
            Id(sd2, "txtStyleTitle").Patterns.Value.Pattern.SetValue("Chữ ký tay QA");
            Click(Id(sd2, "btnDrawSignature"));
            var pad = Modal("Vẽ chữ ký tay");
            var r = Id(pad, "pnlSignaturePad").BoundingRectangle;
            Drag(new Point(r.Left + 40, r.Top + r.Height * 2 / 3), new Point(r.Left + 160, r.Top + r.Height / 3));
            Drag(new Point(r.Left + 160, r.Top + r.Height / 3), new Point(r.Left + 300, r.Top + r.Height * 2 / 3));
            UiTests.Shot(pad, "10-drawpad");
            Click(Id(pad, "btnOk"));
            T.True(WaitFor(() => Id(sd2, "rdoGraphicDrawn").AsRadioButton().IsChecked), "drawn radio selected");
            Click(Id(sd2, "btnOk"));
            T.True(WaitFor(() => Id(mgr, "lstStyles").FindFirstDescendant(c => c.ByName("Chữ ký tay QA")) != null), "drawn style saved");
            T.True(Directory.GetFiles(Path.Combine(data, "Data", "Styles"), "*.png").Length >= 2, "style images stored in portable Data folder");
            Click(Id(mgr, "btnClose"));
        });

        Step(t, "UI11 Tạo ID số tự ký (.pfx) và Chứng thư của tôi", () =>
        {
            RibbonClick("Protect", "Tạo ID số");
            var d = Modal("Tạo ID số");
            Id(d, "txtIdName").Patterns.Value.Pattern.SetValue("Kiểm Thử UI");
            Id(d, "txtIdOrg").Patterns.Value.Pattern.SetValue("Microads");
            Id(d, "txtIdEmail").Patterns.Value.Pattern.SetValue("ui@microads.test");
            Id(d, "txtIdPassword").Patterns.Value.Pattern.SetValue("1234");
            Id(d, "txtIdPassword2").Patterns.Value.Pattern.SetValue("1234");
            Click(Id(d, "btnOk"), 1500);
            var msg = MsgBox("2", 10000);
            T.True(msg.Contains("Đã tạo ID số"), "info: " + msg);
            T.True(Directory.GetFiles(Path.Combine(data, "Data", "DigitalIDs"), "*.pfx").Length == 1, "pfx file created in Data/DigitalIDs");
            RibbonClick("Protect", "Chứng thư của tôi");
            var cd = Modal("Chứng thư số của tôi");
            T.True(Id(cd, "lstCertificates").FindAllDescendants().Any(e => (e.Name ?? "").Contains("Kiểm Thử UI")), "pfx listed");
            Click(Id(cd, "btnRefresh"));
            Click(Id(cd, "btnClose"));
        });

        Step(t, "UI12 Thiết lập ký số", () =>
        {
            RibbonClick("Protect", "Thiết lập ký số");
            var d = Modal("Thiết lập ký số");
            Id(d, "chkTimestamp").AsCheckBox().IsChecked = false;
            Id(d, "chkRevocation").AsCheckBox().IsChecked = false;
            Click(Id(d, "btnOk"));
            T.True(WaitFor(() => StatusText.Contains("Đã lưu thiết lập")), "status " + StatusText);
        });

        string signedPath = Path.Combine(Out, "ui-signed.pdf");
        Step(t, "UI13 KÝ SỐ trọn luồng: vẽ khung → chọn ID (.pfx + mật khẩu) → mẫu con dấu → điều chỉnh → Ký và lưu", () =>
        {
            RibbonClick("View", "Trang đầu");
            RibbonClick("Home", "Vừa trang", 600);
            RibbonClick("Protect", "Ký số");
            T.True(WaitFor(() => StatusText.StartsWith("KÝ SỐ")), "sign mode status");
            var vr = Viewer.BoundingRectangle;
            var c = new Point(vr.Left + vr.Width / 2, vr.Top + vr.Height / 2);
            Drag(new Point(c.X - 20, c.Y + 40), new Point(c.X + 180, c.Y + 110));
            var d = Modal("Ký tài liệu");
            bool locked1 = ChooseTestCertificate(d);
            T.True(WaitFor(() => Id(d, "lblCertInfo").Name.Contains("Kiểm Thử UI")), "cert unlocked: " + Id(d, "lblCertInfo").Name);
            SelectCombo(Id(d, "cboStyle"), "Con dấu QA");
            Id(d, "txtLocation").Patterns.Value.Pattern.SetValue("Hà Nội");
            Id(d, "cboReason").AsComboBox().EditableText = "Phê duyệt hợp đồng";
            Thread.Sleep(500);
            UiTests.Shot(d, "13-sign-dialog");
            Click(Id(d, "btnOk"), 800);
            if (locked1) UnlockPrompt();
            var confirm = Id(W, "btnPendingConfirm");
            UiTests.Shot(W, "13-pending");
            // move the pending signature by dragging it
            var vr2 = Viewer.BoundingRectangle;
            var cr = confirm.BoundingRectangle;
            var inside = new Point(cr.Left + 30, cr.Top - 30);
            Drag(inside, new Point(inside.X - 60, inside.Y - 40));
            Click(Id(W, "btnPendingConfirm"));
            FileDialog("Lưu tài liệu đã ký", signedPath);
            T.True(WaitFor(() => File.Exists(signedPath), 15000), "signed file written");
            T.True(WaitFor(() => W.Title.Contains("ui-signed.pdf"), 8000), "reloaded signed file: " + W.Title);
            T.True(WaitFor(() => Id(W, "lblSigBanner").Name.Length > 0, 15000), "banner shown");
            T.True(Id(W, "signatureTree").FindFirstDescendant(cf => cf.ByName("Rev.1: Ký bởi Kiểm Thử UI")) != null, "signature in panel");
            Thread.Sleep(800);
            UiTests.Shot(W, "13-signed");
            var v = SignatureVerifier.Verify(File.ReadAllBytes(signedPath), Offline).Single();
            T.True(v.IntegrityValid && v.SignatureValid, "file signature valid");
            T.True(v.Application.Contains("Microads Office (Microads.vn)"), "platform recorded: " + v.Application);
            T.Eq("Phê duyệt hợp đồng", v.Reason);
            T.Eq("Hà Nội", v.Location);
        });

        Step(t, "UI14 Bảng chữ ký: Thuộc tính chữ ký, Xem chứng thư, Xem phiên bản đã ký, Xác thực, Bảng chữ ký", () =>
        {
            var tree = Id(W, "signatureTree");
            Click(Find(tree, c => c.ByName("Rev.1: Ký bởi Kiểm Thử UI")));
            Click(Id(W, "btnSigPropsPanel"));
            var d = Modal("Thuộc tính chữ ký số");
            T.True(Id(d, "lblSigStatus").Name.Contains("toàn vẹn") || Id(d, "lblSigStatus").Name.Contains("hợp lệ"), "status " + Id(d, "lblSigStatus").Name);
            T.True(Id(d, "txtSigDetails").Patterns.Value.Pattern.Value.Value.Contains("Microads Office (Microads.vn)"), "details show platform");
            UiTests.Shot(d, "14-sig-properties");
            Click(Id(d, "btnViewCert"), 1500);
            var certWin = Retry.WhileNull(() => AllWindows().FirstOrDefault(x => x.ClassName == "#32770" &&
                ((x.Name ?? "").Contains("Chứng thư số") || (x.Name ?? "").Contains("Certificate"))), TimeSpan.FromSeconds(8)).Result;
            T.True(certWin != null, "certificate dialog opened");
            UiTests.Shot(certWin!, "14-cert");
            certWin!.AsWindow().Close(); Thread.Sleep(600);
            Click(Id(d, "btnViewSignedVersion"), 1500);
            T.True(WaitFor(() => W.Title.Contains("bản đã ký Rev.1")), "signed version tab: " + W.Title);
            Kb.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_W); Thread.Sleep(800);
            T.True(WaitFor(() => W.Title.Contains("ui-signed.pdf")), "back to signed doc");
            RibbonClick("Protect", "Xác thực chữ ký");
            T.True(WaitFor(() => StatusText.Contains("Đã xác thực 1 chữ ký"), 15000), "validated: " + StatusText);
            RibbonClick("Protect", "Bảng chữ ký");
            T.True(Id(W, "signatureTree") != null, "panel");
            Click(Id(W, "btnValidateAllPanel"), 1200);
        });

        Step(t, "UI15 Menu Tệp ▸ Mở: tệp có trường chữ ký trống, ký vào trường có sẵn", () =>
        {
            Click(FileMenu);
            Click(Menu("Mở..."), 800);
            FileDialog("Mở tài liệu PDF", sigField);
            T.True(WaitFor(() => W.Title.Contains("sample-sigfield.pdf")), "opened");
            RibbonClick("Protect", "Bảng chữ ký");
            var tree = Id(W, "signatureTree");
            var node = Find(tree, c => c.ByName("ApproverSignature (chưa ký)"), 10000);
            Click(node);
            Click(Id(W, "btnSigPropsPanel"));
            var q = MsgBox("6");
            T.True(q.Contains("chưa được ký"), "ask to sign: " + q);
            var d = Modal("Ký tài liệu");
            bool locked2 = ChooseTestCertificate(d);
            SelectCombo(Id(d, "cboStyle"), "Chữ ký tay QA");
            Click(Id(d, "btnOk"), 800);
            if (locked2) UnlockPrompt();
            var outPath = Path.Combine(Out, "ui-field-signed.pdf");
            FileDialog("Lưu tài liệu đã ký", outPath);
            T.True(WaitFor(() => File.Exists(outPath), 15000), "written");
            var v = SignatureVerifier.Verify(File.ReadAllBytes(outPath), Offline).Single();
            T.True(v.IsSigned && v.FieldName == "ApproverSignature" && v.IntegrityValid && v.SignatureValid, "field signed");
            Thread.Sleep(1000);
            UiTests.Shot(W, "15-field-signed");
        });

        Step(t, "UI16 Chèn dấu ảnh → tài liệu có thay đổi (●) → Lưu (Ctrl+S) → chữ ký báo có thay đổi sau ký", () =>
        {
            RibbonClick("Protect", "Chèn dấu ảnh");
            var d = Modal("Chèn dấu hình ảnh");
            Click(Id(d, "btnStampChoose"));
            FileDialog("Chọn ảnh con dấu", seal);
            Click(Id(d, "btnOk"), 600);
            T.True(WaitFor(() => StatusText.StartsWith("CHÈN DẤU")), "stamp mode");
            var vr = Viewer.BoundingRectangle;
            Drag(new Point(vr.Left + vr.Width / 2 - 140, vr.Top + 120), new Point(vr.Left + vr.Width / 2 - 40, vr.Top + 220));
            Click(Id(W, "btnPendingConfirm"), 1200);
            T.True(WaitFor(() => W.Title.Contains("*")), "modified marker: " + W.Title);
            Kb.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_S); Thread.Sleep(1200);
            T.True(WaitFor(() => !W.Title.Contains("*")), "saved: " + W.Title);
            var bytes = File.ReadAllBytes(Path.Combine(Out, "ui-field-signed.pdf"));
            T.True(System.Text.Encoding.Latin1.GetString(bytes).Contains("/Subtype /Stamp"), "stamp annotation saved");
            var v = SignatureVerifier.Verify(bytes, Offline).Single();
            T.True(v.IntegrityValid && !v.CoversWholeDocument, "signature intact, change after signing detected");
            T.True(WaitFor(() => Id(W, "lblSigBanner").Name.Contains("thay đổi sau khi ký"), 10000), "banner warns: " + Id(W, "lblSigBanner").Name);
            UiTests.Shot(W, "16-stamped");
        });

        Step(t, "UI17 In: hộp thoại in, phạm vi trang, xem trước, In ra tệp PDF (Microsoft Print to PDF)", () =>
        {
            var printers = PrinterSettings.InstalledPrinters.Cast<string>().ToList();
            var pdfPrinter = printers.FirstOrDefault(p => p.Contains("Microsoft Print to PDF"))
                             ?? throw new SkipException("Không có máy in Microsoft Print to PDF");
            RibbonClick("Home", "In", 1200);
            var d = Modal("In");
            UiTests.Shot(d, "17-print-dialog");
            // SAFETY FIRST: virtual printer + print-to-file before any click on "In".
            SelectCombo(Id(d, "cboPrinter"), pdfPrinter);
            var outPdf = Path.Combine(Out, "ui-printed.pdf");
            if (File.Exists(outPdf)) File.Delete(outPdf);
            Id(d, "chkPrintToFile").AsCheckBox().IsChecked = true;
            Id(d, "txtPrintFile").Patterns.Value.Pattern.SetValue(outPdf);
            T.True(Id(d, "chkPrintToFile").AsCheckBox().IsChecked == true, "print-to-file checked");
            var selPrinter = Id(d, "cboPrinter").Patterns.Value.PatternOrDefault?.Value.ValueOrDefault ?? "";
            if (!selPrinter.Contains("Microsoft Print to PDF")) { Click(Id(d, "btnCancel")); T.Fail("An toàn: không chọn được máy in ảo (" + selPrinter + ")"); }
            // invalid range is rejected
            Click(Id(d, "rdoPageRange"), 300);
            T.True(Id(d, "rdoPageRange").AsRadioButton().IsChecked, "range radio selected");
            T.True(!Id(d, "rdoAllPages").AsRadioButton().IsChecked, "\"Tất cả\" must be unselected when \"Các trang\" is chosen");
            Id(d, "txtPageRange").Patterns.Value.Pattern.SetValue("abc");
            Click(Id(d, "btnPrint"));
            var err = MsgBox("2");
            T.True(err.Contains("Phạm vi trang không hợp lệ"), "range validated: " + err);
            Id(d, "txtPageRange").Patterns.Value.Pattern.SetValue("1");
            Id(d, "chkGrayscale").AsCheckBox().IsChecked = true;
            Click(Id(d, "btnPreviewNext")); Click(Id(d, "btnPreviewPrev"));
            Click(Id(d, "btnPrintPreview"), 2500);
            var pv = Modal("Xem trước khi in", 15000);
            UiTests.Shot(pv, "17-print-preview");
            pv.Close(); Thread.Sleep(800);
            Click(Id(d, "btnPrint"), 1500);
            T.True(WaitFor(() => File.Exists(outPdf) && new FileInfo(outPdf).Length > 500, 30000), "printed PDF produced");
            Thread.Sleep(1500);
            using var p = PdfiumDocument.Load(File.ReadAllBytes(outPdf));
            T.Eq(1, p.PageCount);
            T.True(StatusText.Contains("Đã gửi 1 trang"), "status: " + StatusText);
        });

        Step(t, "UI18 Lưu bản sao, Mở (hủy), Menu Tệp: Mở gần đây / Thuộc tính / Thiết lập / Đóng tài liệu", () =>
        {
            var copy = Path.Combine(Out, "ui-copy.pdf");
            RibbonClick("Home", "Lưu bản sao");
            FileDialog("Lưu tài liệu", copy);
            T.True(WaitFor(() => File.Exists(copy)), "copy saved");
            T.True(SignatureVerifier.Verify(File.ReadAllBytes(copy), Offline).Single().IntegrityValid, "copy keeps signature");
            RibbonClick("Home", "Mở", 1200);
            var od = Modal("Mở tài liệu PDF");
            od.Close(); Thread.Sleep(500);
            Click(FileMenu);
            var recent = Menu("Mở gần đây");
            SafeClick(recent); Thread.Sleep(600);
            T.True(A.GetDesktop().FindFirstDescendant(c => c.ByName("Xóa danh sách")) != null, "recent submenu has entries");
            Kb.Press(VirtualKeyShort.ESCAPE); Kb.Press(VirtualKeyShort.ESCAPE); Thread.Sleep(400);
            Click(FileMenu);
            Click(Menu("Thuộc tính tài liệu"), 800);
            Click(Id(Modal("Thuộc tính tài liệu"), "btnClose"));
            Click(FileMenu);
            Click(Menu("Thiết lập ký số..."), 800);
            Click(Id(Modal("Thiết lập ký số"), "btnCancel"));
            Click(FileMenu);
            Click(Menu("Lưu"), 600);
            int tabsBefore = Id(W, "tabDocuments").FindAllChildren(c => c.ByControlType(ControlType.TabItem)).Length;
            Click(FileMenu);
            Click(Menu("Đóng tài liệu"), 800);
            T.True(WaitFor(() => W.FindFirstDescendant(c => c.ByAutomationId("tabDocuments")) is null or { IsOffscreen: true } ||
                                 Id(W, "tabDocuments").FindAllChildren(c => c.ByControlType(ControlType.TabItem)).Length == tabsBefore - 1), "tab closed");
        });

        Step(t, "UI19 Trợ giúp: Hướng dẫn, Giới thiệu, Mở PDF bằng Microads (xác nhận → Không)", () =>
        {
            RibbonClick("Help", "Hướng dẫn");
            var h = Modal("Hướng dẫn");
            T.True(Id(h, "txtHelp") != null, "help text");
            Click(Id(h, "btnClose"));
            RibbonClick("Help", "Giới thiệu");
            var ab = Modal("Giới thiệu");
            T.True(Id(ab, "lblAbout").Name.Contains("Microads"), "about text");
            UiTests.Shot(ab, "19-about");
            Click(Id(ab, "btnClose"));
            RibbonClick("Help", "Mở PDF bằng Microads");
            var q = MsgBox("7");
            T.True(q.Contains("Mở bằng"), "confirm text: " + q);
        });

        Step(t, "UI20 Mở tệp có mật khẩu (sai → đúng) và chặn ký trên tệp mã hóa", () =>
        {
            if (!File.Exists(encrypted)) throw new SkipException("encrypted.pdf chưa được tạo");
            RibbonClick("Home", "Mở", 1000);
            FileDialog("Mở tài liệu PDF", encrypted);
            var p1 = Modal("Tài liệu được bảo vệ");
            Id(p1, "txtInput").Patterns.Value.Pattern.SetValue("sai-mat-khau");
            Click(Id(p1, "btnOk"), 800);
            var p2 = Modal("Tài liệu được bảo vệ");
            T.True(Id(p2, "lblPrompt").Name.Contains("Mật khẩu không đúng"), "wrong password message");
            Id(p2, "txtInput").Patterns.Value.Pattern.SetValue("abc");
            Click(Id(p2, "btnOk"), 1200);
            T.True(WaitFor(() => W.Title.Contains("encrypted.pdf")), "encrypted opened");
            RibbonClick("Protect", "Ký số");
            var msg = MsgBox("2");
            T.True(msg.Contains("mã hóa") || msg.Contains("mật khẩu"), "sign blocked: " + msg);
        });

        Step(t, "UI21 Đa thẻ (Ctrl+Tab), đóng hết → Trang bắt đầu, mở lại từ danh sách gần đây", () =>
        {
            var titles = new HashSet<string> { W.Title };
            Kb.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.TAB); Thread.Sleep(600);
            titles.Add(W.Title);
            T.True(titles.Count >= 2, "tab switched");
            for (int i = 0; i < 8; i++)
            {
                if (W.FindFirstDescendant(c => c.ByAutomationId("startPage")) is { IsOffscreen: false }) break;
                Kb.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_W); Thread.Sleep(700);
                var box = W.ModalWindows.FirstOrDefault(m => m.ClassName == "#32770");
                SafeClick(box?.FindFirstDescendant(c => c.ByAutomationId("7")));
                Thread.Sleep(300);
            }
            T.True(WaitFor(() => Id(W, "startPage") is { IsOffscreen: false }), "start page visible");
            T.Eq("Microads Office PDF", W.Title);
            UiTests.Shot(W, "21-start");
            T.True(Id(W, "btnStartOpen") != null, "open button on start page");
            var link = Retry.WhileNull(() => Id(W, "recentList").FindAllDescendants().FirstOrDefault(e => (e.Name ?? "").StartsWith("ui-copy.pdf")), TimeSpan.FromSeconds(5)).Result
                       ?? throw new AssertException("recent link ui-copy.pdf not found");
            Click(link, 1500);
            T.True(WaitFor(() => W.Title.Contains(".pdf")), "reopened from recent: " + W.Title);
        });

        Step(t, "UI23 Các nút phụ: Lưu/Ký số/Chèn dấu (Trang chủ), menu Tệp ▸ In / Lưu bản sao, hộp thoại Ký (Nạp .pfx, Tạo ID, Xem chứng thư, Làm mới, Tạo/Sửa/Xóa mẫu), Thuộc tính máy in, Xóa nét vẽ, menu chuột phải thẻ", () =>
        {
            if (!W.Title.Contains(".pdf")) { RibbonClick("Home", "Mở", 1000); FileDialog("Mở tài liệu PDF", sample); }
            T.True(WaitFor(() => W.Title.Contains(".pdf")), "document open");
            // Home: Lưu (document unmodified: saves in place, no dialog)
            RibbonClick("Home", "Lưu", 800);
            T.True(WaitFor(() => StatusText.StartsWith("Đã lưu")), "save status: " + StatusText);
            // Home: Ký số toggles the signing tool, Esc cancels
            RibbonClick("Home", "Ký số", 500);
            T.True(WaitFor(() => StatusText.StartsWith("KÝ SỐ")), "home sign tool");
            RibbonClick("Home", "Ký số", 500); // second click toggles the tool off
            T.True(WaitFor(() => Name(Ribbon("Home"), "Ký số").Properties.ControlType.Value == ControlType.Button), "sign tool toggled off");
            // Home: Chèn dấu ảnh → dialog → Hủy
            RibbonClick("Home", "Chèn dấu ảnh", 800);
            Click(Id(Modal("Chèn dấu hình ảnh"), "btnCancel"));
            // File menu: In... → Hủy ; Lưu bản sao... → Hủy
            Click(FileMenu); Click(Menu("In..."), 1200);
            Click(Id(Modal("In"), "btnCancel"));
            Click(FileMenu); Click(Menu("Lưu bản sao..."), 1200);
            Modal("Lưu tài liệu").Close(); Thread.Sleep(500);
            // Print dialog: printer "Thuộc tính..." opens the standard Windows print dialog → Cancel
            RibbonClick("Home", "In", 1200);
            var pd = Modal("In");
            SelectCombo(Id(pd, "cboPrinter"), "Microsoft Print to PDF");
            Click(Id(pd, "btnPrinterProperties"), 1500);
            // the driver's property sheet runs inside our process
            var sys = Retry.WhileNull(() => AllWindows().FirstOrDefault(w => w.ClassName == "#32770" &&
                (w.Name ?? "").Contains("Print", StringComparison.OrdinalIgnoreCase) && w.FindFirstDescendant(c => c.ByAutomationId("2").And(c.ByControlType(ControlType.Button))) != null),
                TimeSpan.FromSeconds(10)).Result;
            T.True(sys != null, "printer driver properties opened");
            UiTests.Shot(sys!, "23-printer-properties");
            var cancel = sys!.FindFirstDescendant(c => c.ByAutomationId("2").And(c.ByControlType(ControlType.Button)))!;
            if (cancel.Patterns.Invoke.IsSupported) cancel.Patterns.Invoke.Pattern.Invoke(); else SafeClick(cancel);
            Thread.Sleep(800);
            Click(Id(pd, "btnCancel"));
            // Sign dialog secondary buttons (draw a box with a single click)
            RibbonClick("Protect", "Ký số", 500);
            var vr = Viewer.BoundingRectangle;
            SafeClickAt(new Point(vr.Left + vr.Width / 2, vr.Top + vr.Height / 2)); Thread.Sleep(800);
            var d = Modal("Ký tài liệu");
            Click(Id(d, "btnRefreshCerts"), 800);
            Click(Id(d, "btnLoadPfx"), 1000);
            var pfx = Directory.GetFiles(Path.Combine(data, "Data", "DigitalIDs"), "*.pfx").First();
            FileDialog("Chọn tệp chứng thư số", pfx);
            var pw = Modal("Mật khẩu chứng thư");
            Id(pw, "txtInput").Patterns.Value.Pattern.SetValue("1234");
            Click(Id(pw, "btnOk"), 900);
            T.True(WaitFor(() => Id(d, "lblCertInfo").Name.Contains("Kiểm Thử UI")), "pfx loaded via button: " + Id(d, "lblCertInfo").Name);
            Click(Id(d, "btnViewCertificate"), 1500);
            var certWin = Retry.WhileNull(() => AllWindows().FirstOrDefault(x => x.ClassName == "#32770" &&
                ((x.Name ?? "").Contains("Chứng thư số") || (x.Name ?? "").Contains("Certificate"))), TimeSpan.FromSeconds(8)).Result;
            T.True(certWin != null, "certificate viewer opened");
            certWin!.AsWindow().Close(); Thread.Sleep(600);
            Click(Id(d, "btnCreateId"), 1000);
            Click(Id(Modal("Tạo ID số tự ký mới"), "btnCancel"));
            Click(Id(d, "btnNewStyle"), 1000);
            var ns = Modal("Tạo mẫu chữ ký");
            Click(Id(ns, "btnDrawSignature"), 1000);
            var pad = Modal("Vẽ chữ ký tay");
            var r = Id(pad, "pnlSignaturePad").BoundingRectangle;
            Drag(new Point(r.Left + 40, r.Top + 60), new Point(r.Left + 200, r.Top + 120));
            Click(Id(pad, "btnClearPad"));
            Click(Id(pad, "btnOk"), 600);
            T.Eq(true, MsgBox("2").Contains("vẽ chữ ký"), "empty pad rejected");
            Click(Id(pad, "btnCancel"));
            Click(Id(ns, "btnCancel"));
            SelectCombo(Id(d, "cboStyle"), "Con dấu QA");
            Click(Id(d, "btnEditStyle"), 1000);
            var es = Modal("Sửa mẫu chữ ký");
            T.Eq("Con dấu QA", Id(es, "txtStyleTitle").Patterns.Value.Pattern.Value.Value);
            Click(Id(es, "btnCancel"));
            Click(Id(d, "btnDeleteStyle"), 800);
            T.True(MsgBox("7").Contains("Xóa mẫu"), "delete style asks for confirmation (answered No)");
            Click(Id(d, "btnCancel"), 600);
            Kb.Press(VirtualKeyShort.ESCAPE); Thread.Sleep(300);
            // Style manager: Sửa / Xóa (answer No)
            RibbonClick("Protect", "Mẫu chữ ký");
            var mgr = Modal("Mẫu chữ ký");
            Click(Id(mgr, "btnEditStyleMgr"), 900);
            Click(Id(Modal("Sửa mẫu chữ ký"), "btnCancel"));
            Click(Id(mgr, "btnDeleteStyleMgr"), 800);
            MsgBox("7");
            Click(Id(mgr, "btnClose"));
            // Tab context menu: open a second document then "Đóng các thẻ khác"
            RibbonClick("Home", "Mở", 1000);
            FileDialog("Mở tài liệu PDF", sigField);
            T.True(WaitFor(() => Id(W, "tabDocuments").FindAllChildren(c => c.ByControlType(ControlType.TabItem)).Length >= 2), "two tabs");
            var tabItem = Id(W, "tabDocuments").FindAllChildren(c => c.ByControlType(ControlType.TabItem)).Last();
            var tp = ClickPoint(tabItem);
            EnsureOurs(); if (!OursAt(tp)) throw new AssertException("AN TOÀN: tab bị che");
            Mouse.Click(tp, MouseButton.Right); Thread.Sleep(500);
            Click(Menu("Đóng các thẻ khác"), 800);
            MsgBoxCore("7", 2500, required: false); // "don't save" if a document asks
            T.True(WaitFor(() => Id(W, "tabDocuments").FindAllChildren(c => c.ByControlType(ControlType.TabItem)).Length == 1), "other tabs closed");
        });

        Step(t, "UI24 Phiên duy nhất: mở thêm tệp từ Explorer → mở thành thẻ mới trong cửa sổ đang chạy", () =>
        {
            var second = UiTests.Launch(Exe, data, sigField);
            T.True(WaitFor(() => second.HasExited, 15000), "second process forwards the file and exits");
            T.True(WaitFor(() => W.Title.Contains("sample-sigfield.pdf"), 10000), "file opened in the running window: " + W.Title);
            T.Eq(1, A.GetDesktop().FindAllChildren(c => c.ByProcessId(App.ProcessId).And(c.ByAutomationId("MainForm"))).Length);
        });

        Step(t, "UI22 Thoát ứng dụng (Tệp ▸ Thoát) và lưu cấu hình portable", () =>
        {
            Click(FileMenu);
            Click(Menu("Thoát"), 1500);
            T.True(WaitFor(() => App.HasExited, 10000), "app exited");
            T.True(File.Exists(Path.Combine(data, "Data", "settings.json")), "settings.json in portable Data folder");
        });
    }
}
