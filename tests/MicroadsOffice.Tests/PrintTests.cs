using System.Drawing;
using MicroadsOffice.Printing;

namespace MicroadsOffice.Tests;

public static class PrintTests
{
    public static void Run(TestRunner t)
    {
        t.Test("Print: page range parsing (1-3, 5, 8-, reverse, invalid)", () =>
        {
            T.Eq("0,1,2,4", string.Join(",", PdfPrinter.ParseRange("1-3, 5", 10)!));
            T.Eq("7,8,9", string.Join(",", PdfPrinter.ParseRange("8-", 10)!));
            T.Eq("0,1", string.Join(",", PdfPrinter.ParseRange("-2", 10)!));
            T.Eq("2,1,0", string.Join(",", PdfPrinter.ParseRange("3-1", 10)!));
            T.True(PdfPrinter.ParseRange("abc", 10) == null, "letters rejected");
            T.True(PdfPrinter.ParseRange("0", 10) == null, "page 0 rejected");
            T.True(PdfPrinter.ParseRange("11", 10) == null, "out of range rejected");
            T.True(PdfPrinter.ParseRange("", 10) == null, "empty rejected");
        });

        t.Test("Print: layout fit / actual size / shrink, centring and auto-rotate", () =>
        {
            var a4 = new SizeF(595, 842);                     // points
            var printable = new RectangleF(0, 0, 787, 1119);    // ~A4 in 1/100 inch minus margins
            var fit = PdfPrinter.Layout(a4, printable, new PrintOptions { Scaling = PrintScaling.Fit }, false, out bool rot);
            T.True(!rot, "portrait page on portrait paper: no rotation");
            T.True(Math.Abs(fit.Width - 787) < 1 || Math.Abs(fit.Height - 1119) < 1, "fit touches an edge");
            var actual = PdfPrinter.Layout(a4, printable, new PrintOptions { Scaling = PrintScaling.ActualSize }, false, out _);
            T.True(Math.Abs(actual.Width - 595 / 72f * 100) < 0.5, "actual size 100%");
            T.True(Math.Abs(actual.X - (printable.Width - actual.Width) / 2) < 0.5, "centred");
            var big = new SizeF(1190, 1684);                   // A2 page
            var shrink = PdfPrinter.Layout(big, printable, new PrintOptions { Scaling = PrintScaling.ShrinkOversized }, false, out _);
            T.True(shrink.Width <= printable.Width + 0.5 && shrink.Height <= printable.Height + 0.5, "oversized page shrunk");
            var small = PdfPrinter.Layout(new SizeF(200, 200), printable, new PrintOptions { Scaling = PrintScaling.ShrinkOversized }, false, out _);
            T.True(Math.Abs(small.Width - 200 / 72f * 100) < 0.5, "small page not enlarged");
            var land = PdfPrinter.Layout(new SizeF(842, 595), printable, new PrintOptions { Scaling = PrintScaling.Fit, AutoRotate = true }, false, out bool rot2);
            T.True(rot2 && land.Height > land.Width, "landscape page rotated onto portrait paper");
        });
    }
}
