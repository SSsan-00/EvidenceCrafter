using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EvidenceCrafter.Core.Models;

namespace EvidenceCrafter.Excel;

/// <summary>
/// Reconnects to the selected Workbook through the ROT, reveals the verified placement cell,
/// and selects the corresponding Excel cell without changing Workbook content.
/// </summary>
public sealed class ExcelPlacementFocusService : IPlacementFocusService
{
  public static bool BringToForeground(WorkbookIdentity workbook)
  {
    var window = workbook.ExcelWindowHandle;
    if (window == 0 || workbook.WindowSessionToken == 0 ||
        NativeMethods.GetProp(workbook.ExcelDocumentWindowHandle, WorkbookSessionTokenRegistry.WindowPropertyName) != workbook.WindowSessionToken)
      return false;
    NativeMethods.GetWindowThreadProcessId(window, out var processId);
    if (processId != workbook.ProcessId) return false;

    // The workbook identity is bound to this top-level Excel window.  Maximizing
    // that handle keeps other Excel windows at their current size.
    NativeMethods.ShowWindow(window, NativeMethods.MaximizeWindowCommand);
    return NativeMethods.SetForegroundWindow(window);
  }
  public FocusResult FocusPlacedImage(
    WorkbookIdentity workbook,
    string worksheetName,
    CellReference focusCell)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    if (focusCell.Row is < 1 or > ExcelWorksheetLimits.MaximumRow ||
      focusCell.Column is < 1 or > ExcelWorksheetLimits.MaximumColumn)
    {
      throw new ArgumentOutOfRangeException(nameof(focusCell), "Excel row or column exceeds the worksheet limits.");
    }

    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return new FocusResult(false, "Excel focus must run on an STA thread.");
    }

    if (string.IsNullOrWhiteSpace(workbook.RotMonikerDisplayName) ||
      workbook.WindowSessionToken == IntPtr.Zero)
    {
      return new FocusResult(false, "The Workbook does not have a verifiable open session; refresh and select it again.");
    }

    IRunningObjectTable? runningObjectTable = null;
    IEnumMoniker? monikerEnumerator = null;
    IBindCtx? bindContext = null;
    try
    {
      Marshal.ThrowExceptionForHR(NativeMethods.GetRunningObjectTable(0, out runningObjectTable));
      runningObjectTable.EnumRunning(out monikerEnumerator);
      Marshal.ThrowExceptionForHR(NativeMethods.CreateBindCtx(0, out bindContext));
      var monikers = new IMoniker[1];
      while (monikerEnumerator.Next(1, monikers, IntPtr.Zero) == 0)
      {
        object? runningObject = null;
        var targetMoniker = false;
        try
        {
          if (!MonikerMatches(monikers[0], bindContext, workbook))
          {
            continue;
          }

          targetMoniker = true;

          runningObjectTable.GetObject(monikers[0], out runningObject);
          if (runningObject is null)
          {
            continue;
          }

          var result = TryFocusRunningObject(runningObject, workbook, worksheetName, focusCell);
          if (result is not null)
          {
            return result;
          }
        }
        catch (Exception exception) when (IsAutomationFailure(exception))
        {
          if (targetMoniker)
          {
            return new FocusResult(
              false,
              $"The selected Workbook could not be focused ({exception.GetType().Name}, 0x{GetAutomationHResult(exception):X8}).");
          }

          // A stale or busy ROT entry must not prevent checking the remaining Excel objects.
        }
        finally
        {
          ComRelease.Release(runningObject);
          ComRelease.Release(monikers[0]);
          monikers[0] = null!;
        }
      }

      return new FocusResult(false, "The selected Workbook is no longer available in its Excel instance.");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return new FocusResult(false, $"Excel focus failed (0x{GetAutomationHResult(exception):X8}).");
    }
    finally
    {
      ComRelease.Release(bindContext);
      ComRelease.Release(monikerEnumerator);
      ComRelease.Release(runningObjectTable);
    }
  }

  private static bool MonikerMatches(
    IMoniker moniker,
    IBindCtx bindContext,
    WorkbookIdentity identity)
  {
    moniker.GetDisplayName(bindContext, null, out var displayName);
    return string.Equals(displayName, identity.RotMonikerDisplayName, StringComparison.Ordinal);
  }

  private static FocusResult? TryFocusRunningObject(
    object runningObject,
    WorkbookIdentity identity,
    string worksheetName,
    CellReference focusCell)
  {
    if (TryGetProperty(runningObject, "Workbooks", out var workbooks))
    {
      try
      {
        if (!ApplicationMatches(runningObject, identity))
        {
          return null;
        }

        return TryFocusFromCollection(runningObject, workbooks, identity, worksheetName, focusCell);
      }
      finally
      {
        ComRelease.Release(workbooks);
      }
    }

    if (!TryGetProperty(runningObject, "Worksheets", out var worksheets))
    {
      return null;
    }

    ComRelease.Release(worksheets);
    if (!TryGetProperty(runningObject, "Application", out var application) || application is null)
    {
      return null;
    }

    try
    {
      if (!ApplicationMatches(application, identity) || !WorkbookMatches(runningObject, identity))
      {
        return null;
      }

      return FocusWorkbook(application, runningObject, identity, worksheetName, focusCell);
    }
    finally
    {
      ComRelease.Release(application);
    }
  }

  private static FocusResult? TryFocusFromCollection(
    object application,
    object? workbooks,
    WorkbookIdentity identity,
    string worksheetName,
    CellReference focusCell)
  {
    if (workbooks is null)
    {
      return null;
    }

    var count = Convert.ToInt32(GetRequiredProperty(workbooks, "Count"), CultureInfo.InvariantCulture);
    for (var index = 1; index <= count; index++)
    {
      object? candidate = null;
      try
      {
        candidate = InvokeProperty(workbooks, "Item", index);
        if (candidate is not null && WorkbookMatches(candidate, identity))
        {
          return FocusWorkbook(application, candidate, identity, worksheetName, focusCell);
        }
      }
      finally
      {
        ComRelease.Release(candidate);
      }
    }

    return null;
  }

  private static FocusResult FocusWorkbook(
    object application,
    object workbook,
    WorkbookIdentity identity,
    string worksheetName,
    CellReference focusCell)
  {
    object? worksheets = null;
    object? worksheet = null;
    object? cells = null;
    object? targetCell = null;
    object? activeWindow = null;
    try
    {
      if (!WorkbookWindowMatchesIdentity(workbook, identity, out var targetWindowHandle))
      {
        return new FocusResult(false, "The selected Workbook was closed or reopened; select it again before focusing.");
      }

      var eventsWereEnabled = Convert.ToBoolean(
        GetRequiredProperty(application, "EnableEvents"),
        CultureInfo.InvariantCulture);

      SetProperty(application, "EnableEvents", false);
      try
      {
        InvokeMethod(workbook, "Activate");
        worksheets = GetRequiredProperty(workbook, "Worksheets");
        worksheet = InvokeProperty(worksheets, "Item", worksheetName) ??
          throw new InvalidOperationException($"Worksheet was not found: {worksheetName}");
        InvokeMethod(worksheet, "Activate");
        cells = GetRequiredProperty(worksheet, "Cells");
        targetCell = InvokeProperty(cells, "Item", focusCell.Row, focusCell.Column) ??
          throw new InvalidOperationException("The placement focus cell could not be resolved.");
        InvokeMethod(application, "Goto", targetCell, true);
        activeWindow = GetRequiredProperty(application, "ActiveWindow");
        SetProperty(activeWindow, "ScrollRow", Math.Max(1, focusCell.Row - 3));
        SetProperty(activeWindow, "ScrollColumn", 1);

        if (!WorkbookWindowMatchesIdentity(workbook, identity, out targetWindowHandle))
        {
          return new FocusResult(false, "The target Workbook window changed during focus; select it again.");
        }

        return new FocusResult(true, $"Focused {worksheetName}!R{focusCell.Row}C{focusCell.Column}.");
      }
      finally
      {
        SetProperty(application, "EnableEvents", eventsWereEnabled);
      }
    }
    finally
    {
      ComRelease.Release(activeWindow);
      ComRelease.Release(targetCell);
      ComRelease.Release(cells);
      ComRelease.Release(worksheet);
      ComRelease.Release(worksheets);
    }
  }

  private static bool WorkbookWindowMatchesIdentity(
    object workbook,
    WorkbookIdentity identity,
    out nint windowHandle)
  {
    if (identity.WindowSessionToken == IntPtr.Zero ||
      !TryGetWorkbookWindowIdentity(
        workbook,
        out windowHandle,
        out var documentWindowHandle,
        out var processId))
    {
      windowHandle = IntPtr.Zero;
      return false;
    }

    return windowHandle == identity.ExcelWindowHandle &&
      documentWindowHandle == identity.ExcelDocumentWindowHandle &&
      processId == identity.ProcessId &&
      NativeMethods.GetProp(documentWindowHandle, WorkbookSessionTokenRegistry.WindowPropertyName) ==
        identity.WindowSessionToken;
  }

  private static bool ApplicationMatches(object application, WorkbookIdentity identity)
  {
    var windowHandle = new IntPtr(Convert.ToInt64(
      GetRequiredProperty(application, "Hwnd"),
      CultureInfo.InvariantCulture));
    if (windowHandle == IntPtr.Zero)
    {
      return false;
    }

    var threadId = NativeMethods.GetWindowThreadProcessId(windowHandle, out var processId);
    return threadId != 0 && processId != 0 && processId == identity.ProcessId;
  }

  private static bool WorkbookMatches(object workbook, WorkbookIdentity identity)
  {
    var name = Convert.ToString(GetRequiredProperty(workbook, "Name"), CultureInfo.CurrentCulture);
    var fullName = Convert.ToString(GetRequiredProperty(workbook, "FullName"), CultureInfo.CurrentCulture);
    return string.Equals(name, identity.Name, StringComparison.OrdinalIgnoreCase) &&
      string.Equals(fullName, identity.FullPath, StringComparison.OrdinalIgnoreCase);
  }

  private static bool TryGetWorkbookWindowIdentity(
    object workbook,
    out nint windowHandle,
    out nint documentWindowHandle,
    out uint processId)
  {
    object? windows = null;
    object? window = null;
    windowHandle = IntPtr.Zero;
    documentWindowHandle = IntPtr.Zero;
    processId = 0;
    try
    {
      windows = GetRequiredProperty(workbook, "Windows");
      var count = Convert.ToInt32(GetRequiredProperty(windows, "Count"), CultureInfo.InvariantCulture);
      if (count < 1)
      {
        return false;
      }

      window = InvokeProperty(windows, "Item", 1);
      if (window is null)
      {
        return false;
      }

      windowHandle = new IntPtr(Convert.ToInt64(
        GetRequiredProperty(window, "Hwnd"),
        CultureInfo.InvariantCulture));
      var desktopWindow = NativeMethods.FindWindowEx(windowHandle, IntPtr.Zero, "XLDESK", null);
      documentWindowHandle = desktopWindow == IntPtr.Zero
        ? IntPtr.Zero
        : NativeMethods.FindWindowEx(desktopWindow, IntPtr.Zero, "EXCEL7", null);
      var threadId = windowHandle == IntPtr.Zero
        ? 0
        : NativeMethods.GetWindowThreadProcessId(windowHandle, out processId);
      return threadId != 0 && processId != 0 && documentWindowHandle != IntPtr.Zero;
    }
    finally
    {
      ComRelease.Release(window);
      ComRelease.Release(windows);
    }
  }

  private static object GetRequiredProperty(object target, string propertyName) =>
    InvokeProperty(target, propertyName) ??
    throw new InvalidOperationException($"COM property returned null: {propertyName}");

  private static bool TryGetProperty(object target, string propertyName, out object? value)
  {
    try
    {
      value = InvokeProperty(target, propertyName);
      return true;
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      value = null;
      return false;
    }
  }

  private static object? InvokeProperty(object target, string propertyName, params object[] arguments) =>
    target.GetType().InvokeMember(
      propertyName,
      BindingFlags.GetProperty,
      binder: null,
      target,
      arguments,
      CultureInfo.CurrentCulture);

  private static void InvokeMethod(object target, string methodName, params object[] arguments) =>
    _ = target.GetType().InvokeMember(
      methodName,
      BindingFlags.InvokeMethod,
      binder: null,
      target,
      arguments,
      CultureInfo.CurrentCulture);

  private static void SetProperty(object target, string propertyName, object value) =>
    _ = target.GetType().InvokeMember(
      propertyName,
      BindingFlags.SetProperty,
      binder: null,
      target,
      [value],
      CultureInfo.CurrentCulture);

  private static bool IsAutomationFailure(Exception exception) =>
    exception is COMException or TargetInvocationException or MissingMemberException or InvalidOperationException;

  private static int GetAutomationHResult(Exception exception) =>
    exception is TargetInvocationException { InnerException: not null } invocationException
      ? invocationException.InnerException!.HResult
      : exception.HResult;

  private static class NativeMethods
  {
    internal const int MaximizeWindowCommand = 3; // SW_MAXIMIZE

    [DllImport("user32.dll")]
    internal static extern bool SetForegroundWindow(nint window);
    [DllImport("user32.dll")]
    internal static extern bool ShowWindow(nint window, int command);
    [DllImport("ole32.dll")]
    internal static extern int GetRunningObjectTable(
      int reserved,
      [MarshalAs(UnmanagedType.Interface)] out IRunningObjectTable runningObjectTable);

    [DllImport("ole32.dll")]
    internal static extern int CreateBindCtx(
      int reserved,
      [MarshalAs(UnmanagedType.Interface)] out IBindCtx bindContext);

    [DllImport("user32.dll")]
    internal static extern uint GetWindowThreadProcessId(nint windowHandle, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern nint FindWindowEx(
      nint parentWindow,
      nint childAfter,
      string className,
      string? windowName);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern nint GetProp(nint windowHandle, string propertyName);

  }
}
