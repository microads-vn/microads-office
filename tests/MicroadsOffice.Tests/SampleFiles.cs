namespace MicroadsOffice.Tests;

public static class SampleFiles
{
    public static void Write(string dir)
    {
        Directory.CreateDirectory(dir);
        File.WriteAllBytes(Path.Combine(dir, "sample-3pages.pdf"), TestPdf.Create(3, withLink: true, withOutline: true));
        File.WriteAllBytes(Path.Combine(dir, "sample-rotated.pdf"), TestPdf.Create(2, rotations: [90, 0]));
        File.WriteAllBytes(Path.Combine(dir, "sample-sigfield.pdf"), TestPdf.Create(1, withEmptySigField: true));
        using var seal = CoreTests.MakeSeal();
        seal.Save(Path.Combine(dir, "seal.png"), System.Drawing.Imaging.ImageFormat.Png);
        Console.WriteLine("samples written to " + dir);
    }
}
