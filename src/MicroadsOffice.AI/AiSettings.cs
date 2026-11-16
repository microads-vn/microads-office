using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.AI;

/// <summary>Which service answers the assistant. "None" = the assistant is not connected and never touches the network.</summary>
public enum AiProviderKind { None, ChatGpt, OpenAiKey, ClaudeKey, Compatible }

/// <summary>
/// Non-secret assistant settings, shared by Word, Excel and PDF (stored in office.json under "ai").
/// Secrets (API keys, ChatGPT tokens) never live here: see <see cref="SecretStore"/>.
/// </summary>
public sealed class AiSettings
{
    public AiProviderKind Provider { get; set; } = AiProviderKind.None;
    public string ChatGptModel { get; set; } = "";
    public string OpenAiModel { get; set; } = "";
    public string ClaudeModel { get; set; } = AiDefaults.ClaudeModel;
    public string CompatibleModel { get; set; } = "";
    public string CompatibleBaseUrl { get; set; } = "http://127.0.0.1:11434/v1";
    /// <summary>Send the open document / selection with each question (the user can untick it in the pane).</summary>
    public bool SendDocument { get; set; } = true;
    /// <summary>The ChatGPT "You're using your ChatGPT plan" notice was shown once (OpenAI UI guideline).</summary>
    public bool ChatGptWelcomeShown { get; set; }
    /// <summary>Claude Desktop (MCP) may read and edit documents open in Microads Office.</summary>
    public bool AllowClaudeDesktop { get; set; }
    /// <summary>Extra standing instructions the user wants applied to every request (like "formal tone").</summary>
    public string Instructions { get; set; } = "";

    const string Key = "ai";
    static readonly JsonSerializerOptions Json = new() { WriteIndented = false };

    public static AiSettings Load()
    {
        try
        {
            OfficeConfig.Reload();
            var raw = OfficeConfig.GetAppData(Key);
            if (!string.IsNullOrWhiteSpace(raw)) return JsonSerializer.Deserialize<AiSettings>(raw, Json) ?? new AiSettings();
        }
        catch { }
        return new AiSettings();
    }

    public void Save() => OfficeConfig.SetAppData(Key, JsonSerializer.Serialize(this, Json));

    public static void Update(Action<AiSettings> change) { var s = Load(); change(s); s.Save(); Changed?.Invoke(); }

    /// <summary>Raised in this process after <see cref="Update"/> (panes refresh their status line).</summary>
    public static event Action? Changed;
    public static void RaiseChanged() => Changed?.Invoke();

    public string ModelFor(AiProviderKind k) => k switch
    {
        AiProviderKind.ChatGpt => ChatGptModel,
        AiProviderKind.OpenAiKey => OpenAiModel,
        AiProviderKind.ClaudeKey => ClaudeModel,
        AiProviderKind.Compatible => CompatibleModel,
        _ => "",
    };

    public void SetModel(AiProviderKind k, string model)
    {
        switch (k)
        {
            case AiProviderKind.ChatGpt: ChatGptModel = model; break;
            case AiProviderKind.OpenAiKey: OpenAiModel = model; break;
            case AiProviderKind.ClaudeKey: ClaudeModel = model; break;
            case AiProviderKind.Compatible: CompatibleModel = model; break;
        }
    }

    public static string ProviderName(AiProviderKind k) => k switch
    {
        AiProviderKind.ChatGpt => Loc.T("ChatGPT (tài khoản)", "ChatGPT (account)"),
        AiProviderKind.OpenAiKey => Loc.T("OpenAI (khóa API)", "OpenAI (API key)"),
        AiProviderKind.ClaudeKey => Loc.T("Claude (khóa API)", "Claude (API key)"),
        AiProviderKind.Compatible => Loc.T("Máy chủ tương thích OpenAI", "OpenAI-compatible server"),
        _ => Loc.T("Chưa kết nối", "Not connected"),
    };
}

public static class AiDefaults
{
    public const string ClaudeModel = "claude-opus-5-5";
    public const string OpenAiApiBase = "https://api.openai.com/v1";
    public const string AnthropicConsole = "https://platform.claude.com/settings/keys";
    public const string OpenAiKeysPage = "https://platform.openai.com/api-keys";
    public const string ChatGptUsage = "https://chatgpt.com/settings/usage";
    public const string ClaudeDesktopDownload = "https://claude.ai/download";
    /// <summary>Maximum characters of document text sent with one request.</summary>
    public const int MaxContextChars = 120_000;
}

/// <summary>Folder for assistant files: <c>Data\ai</c> next to office.json.</summary>
public static class AiPaths
{
    public static string Dir
    {
        get { var d = Path.Combine(AppSettings.DataDirectory, "ai"); Directory.CreateDirectory(d); return d; }
    }
    public static string File(string name) => Path.Combine(Dir, name);
}

/// <summary>
/// Secrets encrypted with Windows DPAPI for the current Windows user (CryptProtectData). The file is useless on another
/// computer or for another Windows account. Written atomically (temp file + move).
/// </summary>
public static class SecretStore
{
    static readonly byte[] Entropy = Encoding.UTF8.GetBytes("MicroadsOffice.AI.v1");

    public static void Write(string name, string? value)
    {
        var path = AiPaths.File(name + ".bin");
        if (string.IsNullOrEmpty(value)) { try { if (System.IO.File.Exists(path)) System.IO.File.Delete(path); } catch { } return; }
        var data = Protect(Encoding.UTF8.GetBytes(value));
        var tmp = path + ".tmp";
        System.IO.File.WriteAllBytes(tmp, data);
        System.IO.File.Move(tmp, path, true);
    }

    public static string? Read(string name)
    {
        var path = AiPaths.File(name + ".bin");
        try { return System.IO.File.Exists(path) ? Encoding.UTF8.GetString(Unprotect(System.IO.File.ReadAllBytes(path))) : null; }
        catch { return null; } // another Windows user / corrupted: treated as "not set"
    }

    public static bool Has(string name) => !string.IsNullOrEmpty(Read(name));

    public static byte[] Protect(byte[] plain) => Crypt(plain, true);
    public static byte[] Unprotect(byte[] cipher) => Crypt(cipher, false);

    static byte[] Crypt(byte[] input, bool protect)
    {
        var inBlob = new DataBlob(); var entropy = new DataBlob(); var outBlob = new DataBlob();
        var hIn = GCHandle.Alloc(input, GCHandleType.Pinned);
        var hEnt = GCHandle.Alloc(Entropy, GCHandleType.Pinned);
        try
        {
            inBlob.cbData = input.Length; inBlob.pbData = hIn.AddrOfPinnedObject();
            entropy.cbData = Entropy.Length; entropy.pbData = hEnt.AddrOfPinnedObject();
            const int CRYPTPROTECT_UI_FORBIDDEN = 0x1;
            bool ok = protect
                ? CryptProtectData(ref inBlob, "Microads Office AI", ref entropy, IntPtr.Zero, IntPtr.Zero, CRYPTPROTECT_UI_FORBIDDEN, ref outBlob)
                : CryptUnprotectData(ref inBlob, IntPtr.Zero, ref entropy, IntPtr.Zero, IntPtr.Zero, CRYPTPROTECT_UI_FORBIDDEN, ref outBlob);
            if (!ok) throw new CryptographicException(Marshal.GetLastWin32Error());
            var result = new byte[outBlob.cbData];
            Marshal.Copy(outBlob.pbData, result, 0, outBlob.cbData);
            return result;
        }
        finally
        {
            hIn.Free(); hEnt.Free();
            if (outBlob.pbData != IntPtr.Zero) { ZeroAndFree(outBlob); }
        }
    }

    static void ZeroAndFree(DataBlob b)
    {
        for (int i = 0; i < b.cbData; i++) Marshal.WriteByte(b.pbData, i, 0);
        LocalFree(b.pbData);
    }

    [StructLayout(LayoutKind.Sequential)]
    struct DataBlob { public int cbData; public IntPtr pbData; }

    [DllImport("crypt32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern bool CryptProtectData(ref DataBlob pDataIn, string szDataDescr, ref DataBlob pOptionalEntropy, IntPtr pvReserved, IntPtr pPromptStruct, int dwFlags, ref DataBlob pDataOut);

    [DllImport("crypt32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern bool CryptUnprotectData(ref DataBlob pDataIn, IntPtr ppszDataDescr, ref DataBlob pOptionalEntropy, IntPtr pvReserved, IntPtr pPromptStruct, int dwFlags, ref DataBlob pDataOut);

    [DllImport("kernel32.dll")]
    static extern IntPtr LocalFree(IntPtr hMem);
}

/// <summary>Names of the secret files.</summary>
public static class SecretNames
{
    public const string OpenAiKey = "openai-key";
    public const string ClaudeKey = "claude-key";
    public const string CompatibleKey = "compatible-key";
    public const string ChatGpt = "chatgpt";
}
