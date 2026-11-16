using System.Reflection;
using MicroadsOffice.Core;
using MicroadsOffice.Dialogs;

namespace MicroadsOffice.Tests;

public static class PreviewProbe
{
    public static void Run(string dataDir, string outPng, string styleTitle)
    {
        AppSettings.BaseDirectory = dataDir;
        var th = new Thread(() =>
        {
            var s = AppSettings.Load();
            using var d = new SignDialog(s, new System.Drawing.SizeF(210, 70));
            var cbo = (ComboBox)typeof(SignDialog).GetField("_style", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(d)!;
            for (int i = 0; i < cbo.Items.Count; i++) if (((MicroadsOffice.Core.Signing.SignatureStyle)cbo.Items[i]!).Title == styleTitle) cbo.SelectedIndex = i;
            var pic = (PictureBox)typeof(SignDialog).GetField("_preview", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(d)!;
            Console.WriteLine("style image loaded: " + (d.StyleImage != null));
            pic.Image?.Save(outPng);
        });
        th.SetApartmentState(ApartmentState.STA); th.Start(); th.Join();
    }
}
