using System.Runtime.CompilerServices;

namespace MicroadsOffice.AI.UI;

/// <summary>
/// Brings the result of background work back to a control's UI thread, whatever SynchronizationContext is current.
/// A plain "await" resumes on the UI thread only when a WinForms SynchronizationContext was current at the await; when it
/// was missing (seen after nested modal loops) the rest of the method ran on a thread-pool thread, and setting
/// Button.Enabled there hung. Code after <c>await UiThread.On(this, …)</c> always runs through Control.BeginInvoke on the
/// control's thread (or right away if the control was disposed meanwhile — the caller checks IsDisposed).
/// </summary>
public static class UiThread
{
    public static UiAwaitable<T> On<T>(Control c, Task<T> work) => new(c, work);

    public static UiAwaitable<bool> On(Control c, Task work) => new(c, Wrap(work));

    static async Task<bool> Wrap(Task t) { await t.ConfigureAwait(false); return true; }
}

public readonly struct UiAwaitable<T> : ICriticalNotifyCompletion
{
    readonly Control _c; readonly Task<T> _work;
    public UiAwaitable(Control c, Task<T> work) { _c = c; _work = work; }

    public UiAwaitable<T> GetAwaiter() => this;

    /// <summary>Already finished and already on the right thread: no hop needed.</summary>
    public bool IsCompleted
    {
        get { try { return _work.IsCompleted && (_c.IsDisposed || !_c.InvokeRequired); } catch { return false; } }
    }

    public T GetResult() => _work.GetAwaiter().GetResult();

    public void OnCompleted(Action continuation) => Schedule(continuation);
    public void UnsafeOnCompleted(Action continuation) => Schedule(continuation);

    void Schedule(Action continuation)
    {
        var c = _c;
        _work.ContinueWith(_ =>
        {
            try
            {
                if (!c.IsDisposed && c.IsHandleCreated) { c.BeginInvoke(continuation); return; }
            }
            catch (InvalidOperationException) { } // handle destroyed between the check and BeginInvoke
            continuation();
        }, CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.Default);
    }
}
