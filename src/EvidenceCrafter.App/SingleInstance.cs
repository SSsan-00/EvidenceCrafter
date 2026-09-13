using System.ComponentModel;
using System.Runtime.InteropServices;

namespace EvidenceCrafter.App;

internal sealed class SingleInstance : IDisposable
{
  // Local scopes kernel objects to the Windows session, independent of the EXE path/name.
  private const string ApplicationId = @"Local\EvidenceCrafter.7258DA22-0C34-4BD9-A9F8-64C4E0C649A7";
  private readonly Mutex mutex;
  private readonly EventWaitHandle activation;
  private readonly string windowMarker;

  internal SingleInstance(string name = ApplicationId)
  {
    windowMarker = name + ".Window";
    // Create the event before claiming ownership: a second launch can signal during startup.
    activation = new EventWaitHandle(false, EventResetMode.AutoReset, name + ".Activate");
    mutex = new Mutex(false, name + ".Mutex");
    try { IsPrimary = mutex.WaitOne(0); }
    catch (AbandonedMutexException) { IsPrimary = true; }
  }

  internal bool IsPrimary { get; }

  internal bool TakeActivationRequest() => activation.WaitOne(0);

  internal void ActivateExistingInstance()
  {
    NativeMethods.EnumWindows((window, _) =>
    {
      if (NativeMethods.GetProp(window, windowMarker) == 0) return true;
      NativeMethods.GetWindowThreadProcessId(window, out var processId);
      NativeMethods.AllowSetForegroundWindow(processId);
      return false;
    }, 0);
    activation.Set();
  }

  internal void Run(Form mainForm)
  {
    void MarkWindow(object? sender, EventArgs args)
    {
      if (!NativeMethods.SetProp(mainForm.Handle, windowMarker, 1))
        throw new Win32Exception(Marshal.GetLastWin32Error());
    }

    mainForm.HandleCreated += MarkWindow;
    if (mainForm.IsHandleCreated) MarkWindow(mainForm, EventArgs.Empty);
    using var timer = new System.Windows.Forms.Timer { Interval = 100 };
    timer.Tick += (_, _) =>
    {
      if (TakeActivationRequest() && !mainForm.IsDisposed && mainForm.IsHandleCreated)
        ActivateWindow(mainForm.Handle);
    };
    timer.Start();
    try { Application.Run(mainForm); }
    finally
    {
      timer.Stop();
      mainForm.HandleCreated -= MarkWindow;
      if (mainForm.IsHandleCreated) NativeMethods.RemoveProp(mainForm.Handle, windowMarker);
    }
  }

  private static void ActivateWindow(nint mainWindow)
  {
    if (NativeMethods.IsIconic(mainWindow)) NativeMethods.ShowWindow(mainWindow, 9); // SW_RESTORE

    // Follow native ownership, including common dialogs and nested modal editors.
    var target = mainWindow;
    while (true)
    {
      var popup = NativeMethods.GetLastActivePopup(target);
      if (popup == target) break;
      target = popup;
      if (NativeMethods.IsWindowVisible(target) && NativeMethods.IsWindowEnabled(target)) break;
    }

    // Never redirect focus to the disabled owner of a modal dialog.
    if (!NativeMethods.IsWindowVisible(target) || !NativeMethods.IsWindowEnabled(target)) return;
    if (NativeMethods.IsIconic(target)) NativeMethods.ShowWindow(target, 9);
    NativeMethods.SetForegroundWindow(target);
  }

  public void Dispose()
  {
    // Main owns the mutex on its STA thread until the entire UI lifetime has ended.
    if (IsPrimary) mutex.ReleaseMutex();
    mutex.Dispose();
    activation.Dispose();
  }

  private static class NativeMethods
  {
    internal delegate bool EnumWindowsCallback(nint window, nint parameter);
    [DllImport("user32.dll")] internal static extern bool EnumWindows(EnumWindowsCallback callback, nint parameter);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern nint GetProp(nint window, string name);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] internal static extern bool SetProp(nint window, string name, nint value);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern nint RemoveProp(nint window, string name);
    [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(nint window, out uint processId);
    [DllImport("user32.dll")] internal static extern bool AllowSetForegroundWindow(uint processId);
    [DllImport("user32.dll")] internal static extern bool IsIconic(nint window);
    [DllImport("user32.dll")] internal static extern bool ShowWindow(nint window, int command);
    [DllImport("user32.dll")] internal static extern nint GetLastActivePopup(nint window);
    [DllImport("user32.dll")] internal static extern bool IsWindowVisible(nint window);
    [DllImport("user32.dll")] internal static extern bool IsWindowEnabled(nint window);
    [DllImport("user32.dll")] internal static extern bool SetForegroundWindow(nint window);
  }
}
