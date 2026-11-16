using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace MicroadsOffice.Core.Office;

/// <summary>
/// Detects the Microsoft Store (MSIX) edition. The Store signs the package with Microsoft's certificate
/// (no SmartScreen / browser warnings) and delivers updates itself, so this edition never downloads or
/// replaces its own executable and never runs the built-in installer.
/// </summary>
public static class StorePackage
{
    private static bool? _packaged;
    private static string? _familyName;

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetCurrentPackageFamilyName(ref uint length, StringBuilder? name);

    private const int AppModelErrorNoPackage = 15700;

    /// <summary>True when running from an installed MSIX package (Microsoft Store edition).</summary>
    public static bool IsPackaged
    {
        get
        {
            if (_packaged.HasValue) return _packaged.Value;
            try
            {
                uint len = 0;
                int rc = GetCurrentPackageFamilyName(ref len, null);
                if (rc == AppModelErrorNoPackage) { _packaged = false; return false; }
                var sb = new StringBuilder((int)len);
                rc = GetCurrentPackageFamilyName(ref len, sb);
                _familyName = rc == 0 ? sb.ToString() : null;
                _packaged = rc == 0;
            }
            catch { _packaged = false; } // Windows 7/8: no package support
            return _packaged.Value;
        }
    }

    public static string? FamilyName => IsPackaged ? _familyName : null;

    /// <summary>Store page of this app (also lists available updates).</summary>
    public static string StoreUri => FamilyName != null ? "ms-windows-store://pdp/?PFN=" + FamilyName : "ms-windows-store://downloadsandupdates";

    public static void OpenStorePage()
    {
        try { Process.Start(new ProcessStartInfo(StoreUri) { UseShellExecute = true }); } catch { }
    }

    /// <summary>Windows "Default apps" settings (packaged apps cannot claim file types themselves).</summary>
    public static void OpenDefaultAppsSettings()
    {
        try { Process.Start(new ProcessStartInfo("ms-settings:defaultapps") { UseShellExecute = true }); } catch { }
    }
}
