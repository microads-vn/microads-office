using MicroadsOffice.UI;

namespace MicroadsOffice.Dialogs;

public class DialogBase : Form
{
    protected readonly FlowLayoutPanel ButtonBar;
    protected readonly Panel Body;
    protected Button? OkButton;
    protected Button? CancelBtn;

    public static Icon? AppIcon { get; set; }

    public DialogBase(string title, int width, int height)
    {
        Text = title;
        Font = Theme.UiFont;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleMode = AutoScaleMode.None;
        ClientSize = new Size(Dpi.Scale(width), Dpi.Scale(height));
        BackColor = Color.White;
        if (AppIcon != null) Icon = AppIcon;
        ButtonBar = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, Height = Dpi.Scale(48),
            Padding = new Padding(Dpi.Scale(10), Dpi.Scale(8), Dpi.Scale(10), Dpi.Scale(8)), BackColor = Theme.PanelBack
        };
        Body = new Panel { Dock = DockStyle.Fill, Padding = new Padding(Dpi.Scale(14)) };
        Controls.Add(Body);
        Controls.Add(ButtonBar);
    }

    protected Button AddButton(string text, string name, DialogResult result, bool primary = false, EventHandler? onClick = null)
    {
        var b = new Button
        {
            Text = text, Name = name, MinimumSize = new Size(Dpi.Scale(92), Dpi.Scale(30)), AutoSize = true,
            FlatStyle = FlatStyle.Flat, Margin = new Padding(Dpi.Scale(4), 0, 0, 0), Cursor = Cursors.Hand
        };
        if (primary)
        {
            b.BackColor = Theme.Accent; b.ForeColor = Color.White; b.FlatAppearance.BorderSize = 0; b.Font = Theme.UiFontBold;
        }
        else
        {
            b.BackColor = Color.White; b.FlatAppearance.BorderColor = Theme.Border;
        }
        if (onClick != null) b.Click += onClick;
        else b.DialogResult = result;
        ButtonBar.Controls.Add(b);
        return b;
    }

    protected void AddOkCancel(string okText = "OK", string cancelText = "Hủy", EventHandler? onOk = null)
    {
        CancelBtn = AddButton(cancelText, "btnCancel", DialogResult.Cancel);
        OkButton = AddButton(okText, "btnOk", DialogResult.OK, primary: true, onOk);
        AcceptButton = OkButton;
        CancelButton = CancelBtn;
    }

    protected static Label L(string text, bool bold = false) => new()
    {
        Text = text, AutoSize = true, Anchor = AnchorStyles.Left, Font = bold ? Theme.UiFontBold : Theme.UiFont,
        Margin = new Padding(0, Dpi.Scale(6), Dpi.Scale(6), Dpi.Scale(2))
    };

    protected static TableLayoutPanel Grid(int cols = 2)
    {
        var t = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = cols, AutoScroll = true };
        t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        for (int i = 1; i < cols; i++) t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / (cols - 1)));
        return t;
    }

    public static void Error(IWin32Window? owner, string message, string title = "Microads Office")
        => MessageBox.Show(owner, message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);

    public static void Info(IWin32Window? owner, string message, string title = "Microads Office")
        => MessageBox.Show(owner, message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
}

/// <summary>Single text input (also used for passwords).</summary>
public sealed class InputDialog : DialogBase
{
    private readonly TextBox _box;
    public string Value => _box.Text;

    public InputDialog(string title, string prompt, string initial = "", bool password = false) : base(title, 420, 160)
    {
        var lbl = new Label { Text = prompt, Dock = DockStyle.Top, Height = Dpi.Scale(40), Name = "lblPrompt" };
        _box = new TextBox { Text = initial, Dock = DockStyle.Top, UseSystemPasswordChar = password, Name = "txtInput", Font = new Font("Segoe UI", 10f) };
        Body.Controls.Add(_box);
        Body.Controls.Add(lbl);
        AddOkCancel();
        Shown += (_, _) => { _box.Focus(); _box.SelectAll(); };
    }

    public static string? Ask(IWin32Window owner, string title, string prompt, string initial = "", bool password = false)
    {
        using var d = new InputDialog(title, prompt, initial, password);
        return d.ShowDialog(owner) == DialogResult.OK ? d.Value : null;
    }
}
