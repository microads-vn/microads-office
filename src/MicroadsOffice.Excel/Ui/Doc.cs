using MicroadsOffice.Core.Office;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

/// <summary>An open workbook with its file path, dirty flag and undo/redo history.</summary>
public sealed class WorkbookDoc
{
    private readonly Stack<(Workbook Book, string Label)> _undo = new();
    private readonly Stack<(Workbook Book, string Label)> _redo = new();
    public const int MaxUndo = 60;

    public Workbook Book { get; private set; }
    public string? FilePath { get; set; }
    public bool Modified { get; set; }
    public string? UntitledName { get; set; }
    /// <summary>The document is a fresh blank workbook nobody touched (replaced when a file is opened).</summary>
    public bool Pristine => FilePath == null && !Modified;
    public DateTime LastAutoRecover;
    public string? RecoverPath;

    public WorkbookDoc(Workbook book, string? path)
    {
        Book = book;
        FilePath = path;
    }

    public string DisplayName => FilePath != null ? Path.GetFileName(FilePath) : UntitledName ?? Loc.T("Sổ làm việc1", "Book1");

    public Sheet Sheet => Book.Active;

    /// <summary>Records the state before a change (call before every modification).</summary>
    /// <summary>Increases on every change, undo and redo (the AI pane uses it to undo only its own change).</summary>
    public long EditSeq { get; private set; }

    public void Checkpoint(string label)
    {
        EditSeq++;
        _undo.Push((Book.Clone(), label));
        if (_undo.Count > MaxUndo)
        {
            var keep = _undo.Take(MaxUndo).Reverse().ToList();
            _undo.Clear();
            foreach (var k in keep) _undo.Push(k);
        }
        _redo.Clear();
        Modified = true;
    }

    public bool CanUndo => _undo.Count > 0;
    public bool CanRedo => _redo.Count > 0;
    public string? UndoLabel => _undo.Count > 0 ? _undo.Peek().Label : null;
    public string? RedoLabel => _redo.Count > 0 ? _redo.Peek().Label : null;

    public bool Undo()
    {
        if (_undo.Count == 0) return false;
        var (book, label) = _undo.Pop();
        EditSeq++;
        _redo.Push((Book, label));
        Book = book;
        Modified = true;
        Recalc();
        return true;
    }

    public bool Redo()
    {
        if (_redo.Count == 0) return false;
        var (book, label) = _redo.Pop();
        EditSeq++;
        _undo.Push((Book, label));
        Book = book;
        Modified = true;
        Recalc();
        return true;
    }

    /// <summary>State to restore if a change started with <see cref="Checkpoint"/> has to be abandoned.</summary>
    public (Workbook Book, string Label)[] SaveRedo() => _redo.ToArray();

    /// <summary>
    /// Abandons the change started by the last <see cref="Checkpoint"/>: the workbook, the redo history and the modified flag go back
    /// exactly as they were (nothing half-applied ends up in Undo or Redo).
    /// </summary>
    public void Rollback((Workbook Book, string Label)[] redo, bool modified)
    {
        if (_undo.Count == 0) return;
        Book = _undo.Pop().Book;
        _redo.Clear();
        foreach (var r in redo.Reverse()) _redo.Push(r);
        Modified = modified;
        EditSeq++;
        Recalc();
    }

    public void Recalc()
    {
        if (!Book.ManualCalc) Book.Recalculate();
    }

    public void ReplaceBook(Workbook wb)
    {
        Book = wb;
        _undo.Clear();
        _redo.Clear();
    }
}

/// <summary>
/// Indirection for modal UI so that the in-process tests can answer dialogs automatically.
/// </summary>
public static class Ux
{
    /// <summary>When set, modal dialogs are passed here instead of being shown (tests fill fields and return a result).</summary>
    public static Func<Form, DialogResult>? DialogHook;
    /// <summary>When set, Open/Save file pickers return this callback's answer.</summary>
    public static Func<bool, string, string?>? FileHook; // (save, filter) → path
    /// <summary>When set, confirmations return this value instead of asking.</summary>
    public static Func<string, bool>? ConfirmHook;
    public static Func<string, SaveChoice>? SaveHook;
    public static Action<string>? MessageHook;

    public static DialogResult Show(Form dlg, IWin32Window? owner)
    {
        if (DialogHook != null) { try { return DialogHook(dlg); } finally { } }
        return dlg.ShowDialog(owner);
    }

    public static string? OpenFile(IWin32Window owner, string filter, string? title = null)
    {
        if (FileHook != null) return FileHook(false, filter);
        using var d = new OpenFileDialog { Filter = filter, Title = title ?? Loc.T("Mở", "Open") };
        return d.ShowDialog(owner) == DialogResult.OK ? d.FileName : null;
    }

    public static string? SaveFile(IWin32Window owner, string filter, string fileName, int filterIndex = 1, string? title = null)
    {
        if (FileHook != null) return FileHook(true, filter + "|" + fileName);
        using var d = new SaveFileDialog { Filter = filter, FileName = fileName, FilterIndex = filterIndex, Title = title ?? Loc.T("Lưu thành", "Save As"), OverwritePrompt = true, AddExtension = true };
        return d.ShowDialog(owner) == DialogResult.OK ? d.FileName : null;
    }

    public static bool Confirm(IWin32Window? owner, string text, string? yes = null, string? no = null, bool warning = false)
        => ConfirmHook != null ? ConfirmHook(text) : Msg.Confirm(owner, text, yes, no, null, warning);

    public static SaveChoice AskSave(IWin32Window? owner, string name) => SaveHook != null ? SaveHook(name) : Msg.AskSave(owner, name);

    public static void Info(IWin32Window? owner, string text) { if (MessageHook != null) MessageHook(text); else Msg.Info(owner, text); }
    public static void Warn(IWin32Window? owner, string text) { if (MessageHook != null) MessageHook(text); else Msg.Warn(owner, text); }
    public static void Error(IWin32Window? owner, string text) { if (MessageHook != null) MessageHook(text); else Msg.Error(owner, text); }
}
