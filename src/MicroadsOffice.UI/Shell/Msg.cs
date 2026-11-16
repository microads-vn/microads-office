using MicroadsOffice.Core.Office;

namespace MicroadsOffice.UI;

public enum SaveChoice { Save, DontSave, Cancel }

/// <summary>
/// Message boxes whose buttons follow the app language (the standard MessageBox uses the Windows
/// language, so an English Windows would show "Yes/No" inside the Vietnamese app).
/// </summary>
public static class Msg
{
    public static string Title => Product.AppName(OfficeShell.App);

    private static TaskDialogButton Show(IWin32Window? owner, string text, TaskDialogIcon icon, string? heading, params TaskDialogButton[] buttons)
    {
        var page = new TaskDialogPage
        {
            Caption = Title,
            Heading = heading,
            Text = text,
            Icon = icon,
            AllowCancel = true,
            SizeToContent = true
        };
        foreach (var b in buttons) page.Buttons.Add(b);
        page.DefaultButton = buttons.FirstOrDefault();
        try
        {
            return owner is IWin32Window w && w.Handle != IntPtr.Zero ? TaskDialog.ShowDialog(w, page) : TaskDialog.ShowDialog(page);
        }
        catch
        {
            // Fallback (should never happen on Windows 10/11)
            var r = MessageBox.Show(owner, (heading != null ? heading + "\n\n" : "") + text, Title, buttons.Length > 1 ? MessageBoxButtons.OKCancel : MessageBoxButtons.OK);
            return r == DialogResult.OK ? buttons[0] : buttons.Last();
        }
    }

    public static void Info(IWin32Window? owner, string text, string? heading = null)
        => Show(owner, text, TaskDialogIcon.Information, heading, new TaskDialogButton(Loc.T("Đóng", "Close")));

    public static void Warn(IWin32Window? owner, string text, string? heading = null)
        => Show(owner, text, TaskDialogIcon.Warning, heading, new TaskDialogButton(Loc.T("Đóng", "Close")));

    public static void Error(IWin32Window? owner, string text, string? heading = null)
        => Show(owner, text, TaskDialogIcon.Error, heading, new TaskDialogButton(Loc.T("Đóng", "Close")));

    /// <summary>Yes/No question with custom button captions.</summary>
    public static bool Confirm(IWin32Window? owner, string text, string? yes = null, string? no = null, string? heading = null, bool warning = false)
    {
        var y = new TaskDialogButton(yes ?? Loc.T("Đồng ý", "Yes"));
        var n = new TaskDialogButton(no ?? Loc.T("Không", "No"));
        return Show(owner, text, warning ? TaskDialogIcon.Warning : TaskDialogIcon.None, heading, y, n) == y;
    }

    /// <summary>"Save changes to X?" — Save / Don't save / Cancel.</summary>
    public static SaveChoice AskSave(IWin32Window? owner, string documentName)
    {
        var save = new TaskDialogButton(Loc.T("Lưu", "Save"));
        var dont = new TaskDialogButton(Loc.T("Không lưu", "Don't save"));
        var cancel = new TaskDialogButton(Loc.T("Hủy", "Cancel"));
        var r = Show(owner, Loc.F("Bạn có muốn lưu các thay đổi trong \"{0}\" không?", "Do you want to save the changes to \"{0}\"?", documentName),
            TaskDialogIcon.Warning, null, save, dont, cancel);
        return r == save ? SaveChoice.Save : r == dont ? SaveChoice.DontSave : SaveChoice.Cancel;
    }
}
