using MicroadsOffice.Core;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;

namespace MicroadsOffice.Tests;

/// <summary>Renders the PDF app windows to PNG in both languages (no mouse / keyboard input).</summary>
public static class RenderProbe
{
    public static void Run(string outDir)
    {
        Exception? error = null;
        var t = new Thread(() => { try { RunSta(outDir); } catch (Exception ex) { error = ex; } });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();
        if (error != null) throw error;
    }

    private static void RunSta(string outDir)
    {
        Directory.CreateDirectory(outDir);
        AppSettings.BaseDirectory = Path.Combine(outDir, "data");
        Directory.CreateDirectory(AppSettings.DataDirectory);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        foreach (var lang in new[] { Loc.Vietnamese, Loc.English })
        {
            OfficeShell.Init(AppKind.Pdf);
            Loc.Lang = lang;
            using var f = (Form)PdfApp.CreateMainForm([]);
            f.StartPosition = FormStartPosition.Manual;
            f.Location = new Point(-4000, 0);
            f.Size = new Size(1280, 800);
            f.ShowInTaskbar = false;
            f.Show();
            Pump();
            Save(f, Path.Combine(outDir, $"pdf-start-{lang}.png"));
            foreach (var tab in new[] { "tabHelp", "tabProtect", "tabView" })
            {
                var strip = Find<ToolStrip>(f, "ribbonTabs");
                strip?.Items[tab]?.PerformClick();
                Pump();
                Save(f, Path.Combine(outDir, $"pdf-{tab}-{lang}.png"));
            }
            using (var about = new AboutOfficeDialog())
            {
                about.StartPosition = FormStartPosition.Manual; about.Location = new Point(-4000, 0); about.ShowInTaskbar = false;
                about.Show(); Pump(); Save(about, Path.Combine(outDir, $"pdf-about-{lang}.png")); about.Close();
            }
            using (var help = new MicroadsOffice.Dialogs.HelpDialog())
            {
                help.StartPosition = FormStartPosition.Manual; help.Location = new Point(-4000, 0); help.ShowInTaskbar = false;
                help.Show(); Pump(); Save(help, Path.Combine(outDir, $"pdf-help-{lang}.png")); help.Close();
            }
            f.Hide();
        }
        Console.WriteLine("rendered to " + outDir);
    }

    private static void Pump() { for (int i = 0; i < 20; i++) { Application.DoEvents(); Thread.Sleep(15); } }

    private static void Save(Control c, string path)
    {
        using var b = new Bitmap(c.Width, c.Height);
        c.DrawToBitmap(b, new Rectangle(0, 0, c.Width, c.Height));
        b.Save(path);
    }

    private static T? Find<T>(Control root, string name) where T : Control
        => root.Controls.Find(name, true).OfType<T>().FirstOrDefault();
}
