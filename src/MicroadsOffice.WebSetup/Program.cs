using System;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace MicroadsOffice.WebSetup
{
    /// <summary>
    ///   MicroadsOffice_Setup.exe                       → window: download the latest Microads Office, verify, start its setup
    ///   MicroadsOffice_Setup.exe --silent [--dir PATH] → no window (IT / mass deployment); exit code = result
    ///   --lang=vi|en                                   → language of this window and of the setup
    ///   --dir PATH must be a full path (D:\Apps\Microads Office, \\server\share\...); %VARIABLES% are expanded.
    /// Exit codes: 0 installed, 1 the setup was closed without installing or reported an error, 2 cancelled,
    ///             3 unsupported Windows, 4 network / nothing published, 5 signature or release check failed,
    ///             6 other error (disk, invalid --dir, setup could not start). Details: %TEMP%\MicroadsOffice_Setup.log
    /// </summary>
    internal static class Program
    {
        public static string Lang = DetectLanguage();
        public static bool Vi => Lang != "en";
        public static string T(string vi, string en) => Vi ? vi : en;

        [STAThread]
        private static int Main(string[] args)
        {
            bool silent = false;
            string dirArg = null;
            for (int i = 0; i < args.Length; i++)
            {
                var a = args[i];
                if (a.Equals("--silent", StringComparison.OrdinalIgnoreCase) || a.Equals("/S", StringComparison.Ordinal)) silent = true;
                else if (a.StartsWith("--lang=", StringComparison.OrdinalIgnoreCase)) Lang = a.Substring(7).ToLowerInvariant() == "en" ? "en" : "vi";
                else if (a.StartsWith("--dir=", StringComparison.OrdinalIgnoreCase)) dirArg = a.Substring(6);
                else if (a.Equals("--dir", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length) dirArg = args[++i];
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (!Environment.Is64BitOperatingSystem || Environment.OSVersion.Version.Major < 10)
            {
                Engine.Log("unsupported Windows " + Environment.OSVersion.VersionString);
                if (!silent)
                {
                    var r = MessageBox.Show(T("Microads Office cần Windows 10 hoặc Windows 11 bản 64-bit.\n\nMở trang microads.vn/office để xem hướng dẫn?",
                            "Microads Office needs 64-bit Windows 10 or Windows 11.\n\nOpen microads.vn/office for more information?"),
                        "Microads Office", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (r == DialogResult.Yes) Engine.OpenUrl(Engine.PageUrl);
                }
                return 3;
            }

            string dir = null;
            if (dirArg != null && (dir = NormalizeDirectory(dirArg)) == null)
            {
                // A relative folder would be resolved against the temporary download folder, which is deleted afterwards.
                Engine.Log("invalid --dir \"" + dirArg + "\": a full path such as D:\\Apps\\Microads Office is required");
                if (!silent)
                    MessageBox.Show(T("Thư mục cài đặt không hợp lệ: \"" + dirArg + "\".\nHãy dùng đường dẫn đầy đủ, ví dụ D:\\Apps\\Microads Office.",
                            "Invalid install folder: \"" + dirArg + "\".\nUse a full path, for example D:\\Apps\\Microads Office."),
                        "Microads Office", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 6;
            }

            if (silent) return RunSilent(SetupArguments(true, dir));

            using (var f = new SetupWindow(() => SetupArguments(false, dir)))
            {
                Application.Run(f);
                return f.ExitCode;
            }
        }

        /// <summary>Full path of the install folder, or null when it is not an absolute path (X:\... or \\server\share\...).</summary>
        internal static string NormalizeDirectory(string d)
        {
            d = Environment.ExpandEnvironmentVariables((d ?? "").Trim().Trim('"').Trim());
            if (d.Length == 0 || d.IndexOf('"') >= 0 || d.IndexOf('%') >= 0) return null;
            bool drive = d.Length >= 3 && char.IsLetter(d[0]) && d[1] == ':' && (d[2] == '\\' || d[2] == '/');
            bool unc = d.StartsWith(@"\\", StringComparison.Ordinal) && d.Length > 2 && d[2] != '?' && d[2] != '.';
            if (!drive && !unc) return null;
            try { d = Path.GetFullPath(d); } catch { return null; }
            var root = Path.GetPathRoot(d);
            return string.Equals(d, root, StringComparison.OrdinalIgnoreCase) ? d : d.TrimEnd('\\', '/');
        }

        /// <summary>
        /// Arguments for the downloaded Microads Office setup (built when it starts: the language may have been switched).
        /// CONTRACT — this program never changes, so every future MicroadsOffice.exe must keep accepting:
        ///   --setup  --lang=vi|en  --silent  --dir "PATH"  --web-setup (marker: started by the web installer; ignored so far)
        /// </summary>
        internal static string SetupArguments(bool silent, string dir)
            => "--setup --web-setup --lang=" + Lang + (silent ? " --silent" : "") + (dir != null ? " --dir " + Quote(dir) : "");

        /// <summary>Windows command-line quoting: a trailing backslash is doubled so that "D:\\" means D:\ (not an escaped quote).</summary>
        private static string Quote(string s) => "\"" + (s.EndsWith("\\", StringComparison.Ordinal) ? s + "\\" : s) + "\"";

        private static int RunSilent(string setupArgs)
        {
            string work = null;
            try
            {
                var r = Engine.FetchLatest(CancellationToken.None);
                work = Engine.CreateWorkFolder();
                var path = Engine.Download(r, work, null, CancellationToken.None);
                System.Diagnostics.Process p;
                using (Engine.VerifyAndLock(path, r, CancellationToken.None)) p = Engine.StartSetup(path, setupArgs);
                p.WaitForExit();
                Engine.Log("silent install of " + r.Version + " finished with code " + p.ExitCode);
                return p.ExitCode == 0 ? 0 : 1;
            }
            catch (SetupFailure ex)
            {
                Engine.Log("silent install failed (" + ex.Kind + "): " + ex.Message);
                return ex.Kind == FailureKind.Network || ex.Kind == FailureKind.NoRelease ? 4 : ex.Kind == FailureKind.Signature ? 5 : 6;
            }
            catch (Exception ex)
            {
                Engine.Log("silent install failed: " + ex);
                return 6;
            }
            finally { Engine.Cleanup(work); }
        }

        /// <summary>Vietnamese for Vietnamese Windows, Vietnamese regional format or the Vietnam time zone; English otherwise.</summary>
        private static string DetectLanguage()
        {
            try
            {
                if (CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "vi" || CultureInfo.CurrentCulture.TwoLetterISOLanguageName == "vi") return "vi";
                if (TimeZoneInfo.Local.Id == "SE Asia Standard Time" && TimeZoneInfo.Local.BaseUtcOffset == TimeSpan.FromHours(7)) return "vi";
            }
            catch { }
            return "en";
        }
    }
}
