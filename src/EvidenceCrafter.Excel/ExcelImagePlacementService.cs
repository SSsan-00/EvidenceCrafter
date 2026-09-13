using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Excel;

/// <summary>
/// Places one image at the user's active cell in a verified Workbook.
/// This first usable slice does not insert or delete rows; it moves the managed Shape with its rows.
/// </summary>
public sealed class ExcelImagePlacementService
{
  private const int MsoFalse = 0;
  private const int MsoTrue = -1;
  private const int XlMove = 2;
  private readonly IPlacementFocusService focusService;
  private readonly ImageSizingService imageSizingService = new();

  public ExcelImagePlacementService(IPlacementFocusService? focusService = null)
  {
    this.focusService = focusService ?? new ExcelPlacementFocusService();
  }

  public ImagePlacementResult PlaceImage(
    WorkbookIdentity workbook,
    string worksheetName,
    CellReference? requestedCell,
    EvidenceSide side,
    string imagePath,
    ImageDimensions imageDimensions,
    double? availableWidthPoints = null,
    double horizontalMarginPoints = 6,
    double? scaleOverride = null)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ArgumentException.ThrowIfNullOrWhiteSpace(imagePath);
    if (requestedCell is not null && !IsValidCell(requestedCell.Value))
    {
      throw new ArgumentOutOfRangeException(nameof(requestedCell));
    }

    if (availableWidthPoints is not null &&
      (!double.IsFinite(availableWidthPoints.Value) || availableWidthPoints.Value <= 0))
    {
      throw new ArgumentOutOfRangeException(nameof(availableWidthPoints));
    }
    if (!double.IsFinite(horizontalMarginPoints) || horizontalMarginPoints < 0)
    {
      throw new ArgumentOutOfRangeException(nameof(horizontalMarginPoints));
    }

    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return ImagePlacementResult.Failed("Excel placement must run on an STA thread.");
    }

    if (workbook.IsReadOnly)
    {
      return ImagePlacementResult.Failed("The selected Workbook is read-only.");
    }

    if (string.IsNullOrWhiteSpace(workbook.RotMonikerDisplayName))
    {
      return ImagePlacementResult.Failed("The Workbook does not have a verifiable open session; refresh and select it again.");
    }

    if (workbook.WindowSessionToken == IntPtr.Zero)
    {
      return ImagePlacementResult.Failed("Workbook connection is unavailable; refresh before placing an image.");
    }

    if (!File.Exists(imagePath))
    {
      return ImagePlacementResult.Failed("The temporary image file no longer exists.");
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
        try
        {
          monikers[0].GetDisplayName(bindContext, null, out var displayName);
          if (!string.Equals(displayName, workbook.RotMonikerDisplayName, StringComparison.Ordinal))
          {
            continue;
          }

          runningObjectTable.GetObject(monikers[0], out runningObject);
          if (runningObject is null)
          {
            return ImagePlacementResult.Failed("The selected Workbook is no longer available in its Excel instance.");
          }

          var result = TryPlaceRunningObject(
            runningObject,
            workbook,
            worksheetName,
            requestedCell,
            side,
            imagePath,
            imageDimensions,
            availableWidthPoints,
            horizontalMarginPoints,
            scaleOverride);
          return result ?? ImagePlacementResult.Failed("The selected Workbook could not be matched in its Excel instance.");
        }
        catch (Exception exception) when (IsAutomationFailure(exception))
        {
          return ImagePlacementResult.Failed(
            $"Excel placement failed (0x{GetAutomationHResult(exception):X8}).");
        }
        finally
        {
          ComRelease.Release(runningObject);
          ComRelease.Release(monikers[0]);
          monikers[0] = null!;
        }
      }

      return ImagePlacementResult.Failed("The selected Workbook is no longer available in the Running Object Table.");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return ImagePlacementResult.Failed($"Excel placement failed (0x{GetAutomationHResult(exception):X8}).");
    }
    finally
    {
      ComRelease.Release(bindContext);
      ComRelease.Release(monikerEnumerator);
      ComRelease.Release(runningObjectTable);
    }
  }

  /// <summary>Deletes one verified EvidenceCrafter Shape by name for application-level Undo.</summary>
  public ImageDeletionResult DeletePlacedImage(
    WorkbookIdentity workbook,
    string worksheetName,
    string shapeName)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);
    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return ImageDeletionResult.Failed("Excel image deletion must run on an STA thread.");
    }

    if (workbook.IsReadOnly ||
      string.IsNullOrWhiteSpace(workbook.RotMonikerDisplayName) ||
      workbook.WindowSessionToken == IntPtr.Zero)
    {
      return ImageDeletionResult.Failed("Workbook session is not writable or verifiable; refresh before Undo.");
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
        try
        {
          monikers[0].GetDisplayName(bindContext, null, out var displayName);
          if (!string.Equals(displayName, workbook.RotMonikerDisplayName, StringComparison.Ordinal))
          {
            continue;
          }

          runningObjectTable.GetObject(monikers[0], out runningObject);
          return runningObject is null
            ? ImageDeletionResult.Failed("The selected Workbook is no longer available.")
            : TryDeleteRunningObject(runningObject, workbook, worksheetName, shapeName) ??
              ImageDeletionResult.Failed("The selected Workbook could not be matched.");
        }
        finally
        {
          ComRelease.Release(runningObject);
          ComRelease.Release(monikers[0]);
          monikers[0] = null!;
        }
      }

      return ImageDeletionResult.Failed("The selected Workbook is no longer available in the Running Object Table.");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return ImageDeletionResult.Failed($"Excel image deletion failed (0x{GetAutomationHResult(exception):X8}).");
    }
    finally
    {
      ComRelease.Release(bindContext);
      ComRelease.Release(monikerEnumerator);
      ComRelease.Release(runningObjectTable);
    }
  }

  private ImageDeletionResult? TryDeleteRunningObject(
    object runningObject,
    WorkbookIdentity identity,
    string worksheetName,
    string shapeName)
  {
    if (TryGetProperty(runningObject, "Workbooks", out var workbooks))
    {
      try
      {
        if (!ApplicationMatches(runningObject, identity))
        {
          return null;
        }

        var count = Convert.ToInt32(GetRequiredProperty(workbooks!, "Count"), CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
          object? candidate = null;
          try
          {
            candidate = InvokeProperty(workbooks!, "Item", index);
            if (candidate is not null && WorkbookMatches(candidate, identity))
            {
              return DeleteShapeFromWorkbook(runningObject, candidate, identity, worksheetName, shapeName);
            }
          }
          finally
          {
            ComRelease.Release(candidate);
          }
        }

        return null;
      }
      finally
      {
        ComRelease.Release(workbooks);
      }
    }

    if (!TryGetProperty(runningObject, "Application", out var application) || application is null)
    {
      return null;
    }

    try
    {
      return ApplicationMatches(application, identity) && WorkbookMatches(runningObject, identity)
        ? DeleteShapeFromWorkbook(application, runningObject, identity, worksheetName, shapeName)
        : null;
    }
    finally
    {
      ComRelease.Release(application);
    }
  }

  private static ImageDeletionResult DeleteShapeFromWorkbook(
    object application,
    object workbook,
    WorkbookIdentity identity,
    string worksheetName,
    string shapeName)
  {
    if (!WorkbookWindowMatchesIdentity(workbook, identity) || IsWorkbookReadOnly(workbook))
    {
      return ImageDeletionResult.Failed("The Workbook changed session or became read-only; Undo was not applied.");
    }

    var eventsWereEnabled = Convert.ToBoolean(GetRequiredProperty(application, "EnableEvents"), CultureInfo.InvariantCulture);

    object? worksheet = null;
    object? shapes = null;
    object? shape = null;
    try
    {
      SetProperty(application, "EnableEvents", false);
      worksheet = ResolveWorksheet(application, workbook, worksheetName, out var resolvedWorksheetName);
      if (IsWorksheetProtected(worksheet))
      {
        return ImageDeletionResult.Failed($"シート {resolvedWorksheetName} は保護されています。Undoできません。");
      }

      shapes = GetRequiredProperty(worksheet, "Shapes");
      shape = InvokeMethod(shapes, "Item", shapeName);
      var alternativeText = shape is null
        ? null
        : Convert.ToString(GetRequiredProperty(shape, "AlternativeText"), CultureInfo.InvariantCulture);
      if (shape is null ||
        alternativeText is null ||
        !alternativeText.StartsWith("CraftEvidence:v1|", StringComparison.Ordinal))
      {
        return ImageDeletionResult.Failed("Undo対象のEvidenceCrafter画像が見つかりません。");
      }

      if (!WorkbookWindowMatchesIdentity(workbook, identity) || IsWorkbookReadOnly(workbook))
      {
        return ImageDeletionResult.Failed("The Workbook changed before Undo; the image was not deleted.");
      }

      InvokeMethod(shape, "Delete");
      return new ImageDeletionResult(true, resolvedWorksheetName, "画像配置をUndoしました（未保存）。");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return ImageDeletionResult.Failed($"画像配置のUndoに失敗しました (0x{GetAutomationHResult(exception):X8})。");
    }
    finally
    {
      try
      {
        SetProperty(application, "EnableEvents", eventsWereEnabled);
      }
      catch (Exception exception) when (IsAutomationFailure(exception))
      {
      }

      ComRelease.Release(shape);
      ComRelease.Release(shapes);
      ComRelease.Release(worksheet);
    }
  }

  private ImagePlacementResult? TryPlaceRunningObject(
    object runningObject,
    WorkbookIdentity identity,
    string worksheetName,
    CellReference? requestedCell,
    EvidenceSide side,
    string imagePath,
    ImageDimensions imageDimensions,
    double? availableWidthPoints,
    double horizontalMarginPoints,
    double? scaleOverride)
  {
    if (TryGetProperty(runningObject, "Workbooks", out var workbooks))
    {
      try
      {
        if (!ApplicationMatches(runningObject, identity))
        {
          return null;
        }

        var count = Convert.ToInt32(GetRequiredProperty(workbooks!, "Count"), CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
          object? candidate = null;
          try
          {
            candidate = InvokeProperty(workbooks!, "Item", index);
            if (candidate is not null && WorkbookMatches(candidate, identity))
            {
              return PlaceWorkbook(
                runningObject,
                candidate,
                identity,
                worksheetName,
                requestedCell,
                side,
                imagePath,
                imageDimensions,
                availableWidthPoints,
                horizontalMarginPoints,
                scaleOverride);
            }
          }
          finally
          {
            ComRelease.Release(candidate);
          }
        }

        return null;
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
      return ApplicationMatches(application, identity) && WorkbookMatches(runningObject, identity)
        ? PlaceWorkbook(
          application,
          runningObject,
          identity,
          worksheetName,
          requestedCell,
          side,
          imagePath,
          imageDimensions,
          availableWidthPoints,
          horizontalMarginPoints,
          scaleOverride)
        : null;
    }
    finally
    {
      ComRelease.Release(application);
    }
  }

  private ImagePlacementResult PlaceWorkbook(
    object application,
    object workbook,
    WorkbookIdentity identity,
    string worksheetName,
    CellReference? requestedCell,
    EvidenceSide side,
    string imagePath,
    ImageDimensions imageDimensions,
    double? availableWidthPoints,
    double horizontalMarginPoints,
    double? scaleOverride)
  {
    if (!WorkbookWindowMatchesIdentity(workbook, identity))
    {
      return ImagePlacementResult.Failed("The selected Workbook was closed or reopened; refresh and select it again.");
    }

    var eventsWereEnabled = Convert.ToBoolean(
      GetRequiredProperty(application, "EnableEvents"),
      CultureInfo.InvariantCulture);

    object? worksheet = null;
    object? activeCell = null;
    object? shapes = null;
    object? targetCell = null;
    object? shape = null;
    var shapeName = $"EST_IMG_{Guid.NewGuid():N}";
    var focusCell = requestedCell;
    var resolvedWorksheetName = worksheetName;
    try
    {
      SetProperty(application, "EnableEvents", false);
      InvokeMethod(workbook, "Activate");
      worksheet = ResolveWorksheet(application, workbook, worksheetName, out resolvedWorksheetName);
      if (IsWorksheetProtected(worksheet))
      {
        return ImagePlacementResult.Failed(
          $"シート {resolvedWorksheetName} は保護されています。保護を解除してから配置してください。");
      }

      InvokeMethod(worksheet, "Activate");
      if (focusCell is null)
      {
        activeCell = GetRequiredProperty(application, "ActiveCell");
        focusCell = ReadCellReference(activeCell);
      }

      targetCell = GetRequiredProperty(worksheet, "Cells", focusCell.Value.Row, focusCell.Value.Column);
      var cellLeft = ReadDoubleProperty(targetCell, "Left");
      var cellTop = ReadDoubleProperty(targetCell, "Top");
      var cellWidth = ReadDoubleProperty(targetCell, "Width");
      var availableWidth = availableWidthPoints ?? Math.Max(cellWidth * 5.0 - (horizontalMarginPoints * 2), 24.0);
      var fittedImage = scaleOverride is { } scale
        ? imageSizingService.AtScale(imageDimensions, availableWidth, scale)
        : imageSizingService.FitToWidth(imageDimensions, availableWidth);

      // Revalidate the session and the live Workbook state immediately before
      // mutating Excel.  The selected identity is only a snapshot; the
      // Workbook can be reopened, become read-only, or lose its close-session
      // token while the user is confirming the preview.
      if (!WorkbookWindowMatchesIdentity(workbook, identity))
      {
        return ImagePlacementResult.Failed("The selected Workbook was closed or reopened; select it again before placing an image.");
      }

      if (IsWorkbookReadOnly(workbook))
      {
        return ImagePlacementResult.Failed("The selected Workbook became read-only before the image could be placed.");
      }

      shapes = GetRequiredProperty(worksheet, "Shapes");
      shape = InvokeMethod(
        shapes,
        "AddPicture",
        imagePath,
        MsoFalse,
        MsoTrue,
        cellLeft + horizontalMarginPoints,
        cellTop + PlacementPlanner.VerticalInsetPoints,
        fittedImage.WidthPoints,
        fittedImage.HeightPoints);
      if (shape is null)
      {
        return ImagePlacementResult.Failed("Excel did not return the inserted image Shape.");
      }

      SetProperty(shape, "Name", shapeName);
      var metadata = new ManagedShapeMetadata(1, side, focusCell.Value)
      {
        SourceDimensions = imageDimensions,
        AppliedScale = fittedImage.Scale,
      };
      var alternativeText = metadata.Serialize();
      SetProperty(shape, "AlternativeText", alternativeText);
      SetProperty(shape, "LockAspectRatio", MsoTrue);
      SetProperty(shape, "Placement", XlMove);

      var insertedName = Convert.ToString(GetRequiredProperty(shape, "Name"), CultureInfo.InvariantCulture);
      if (!string.Equals(insertedName, shapeName, StringComparison.Ordinal))
      {
        TryDeleteShape(shape);
        return ImagePlacementResult.Failed("The inserted image Shape could not be verified.");
      }

      SetProperty(application, "EnableEvents", eventsWereEnabled);
      var focus = focusService.FocusPlacedImage(identity, resolvedWorksheetName, focusCell.Value);
      return new ImagePlacementResult(
        true,
        focus.Succeeded,
        shapeName,
        resolvedWorksheetName,
        focusCell.Value,
        new ManagedShapeTarget(
          shapeName,
          resolvedWorksheetName,
          alternativeText,
          metadata,
          focusCell.Value,
          ReadDoubleProperty(shape, "Left"),
          ReadDoubleProperty(shape, "Top"),
          ReadDoubleProperty(shape, "Width"),
          ReadDoubleProperty(shape, "Height")),
        focus.Succeeded
          ? $"画像を{resolvedWorksheetName}!R{focusCell.Value.Row}C{focusCell.Value.Column}へ配置しました。"
          : $"画像を配置しましたが、対象セルの選択に失敗しました: {focus.Message}");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      TryDeleteShape(shape);
      return ImagePlacementResult.Failed($"Excelへの画像配置に失敗しました (0x{GetAutomationHResult(exception):X8}).");
    }
    finally
    {
      try
      {
        SetProperty(application, "EnableEvents", eventsWereEnabled);
      }
      catch (Exception exception) when (IsAutomationFailure(exception))
      {
        // The Workbook may have disconnected after the Shape was inserted.
      }

      ComRelease.Release(shape);
      ComRelease.Release(targetCell);
      ComRelease.Release(activeCell);
      ComRelease.Release(shapes);
      ComRelease.Release(worksheet);
    }
  }

  private static object ResolveWorksheet(
    object application,
    object workbook,
    string worksheetName,
    out string resolvedName)
  {
    if (string.Equals(worksheetName, "ActiveSheet", StringComparison.OrdinalIgnoreCase))
    {
      var activeSheet = GetRequiredProperty(application, "ActiveSheet");
      try
      {
        resolvedName = Convert.ToString(GetRequiredProperty(activeSheet, "Name"), CultureInfo.CurrentCulture) ??
          throw new InvalidOperationException("The active Excel sheet has no name.");
        return activeSheet;
      }
      catch
      {
        ComRelease.Release(activeSheet);
        throw;
      }
    }

    var worksheets = GetRequiredProperty(workbook, "Worksheets");
    try
    {
      var worksheet = InvokeProperty(worksheets, "Item", worksheetName) ??
        throw new InvalidOperationException($"Worksheet was not found: {worksheetName}");
      try
      {
        resolvedName = Convert.ToString(GetRequiredProperty(worksheet, "Name"), CultureInfo.CurrentCulture) ??
          throw new InvalidOperationException("The selected Excel sheet has no name.");
        return worksheet;
      }
      catch
      {
        ComRelease.Release(worksheet);
        throw;
      }
    }
    finally
    {
      ComRelease.Release(worksheets);
    }
  }

  private static CellReference ReadCellReference(object cell) =>
    new(
      Convert.ToInt32(GetRequiredProperty(cell, "Row"), CultureInfo.InvariantCulture),
      Convert.ToInt32(GetRequiredProperty(cell, "Column"), CultureInfo.InvariantCulture));

  private static bool IsWorksheetProtected(object worksheet) =>
    Convert.ToBoolean(GetRequiredProperty(worksheet, "ProtectContents"), CultureInfo.InvariantCulture) ||
    Convert.ToBoolean(GetRequiredProperty(worksheet, "ProtectDrawingObjects"), CultureInfo.InvariantCulture) ||
    Convert.ToBoolean(GetRequiredProperty(worksheet, "ProtectScenarios"), CultureInfo.InvariantCulture);

  private static bool IsWorkbookReadOnly(object workbook) =>
    Convert.ToBoolean(GetRequiredProperty(workbook, "ReadOnly"), CultureInfo.InvariantCulture);

  private static bool IsValidCell(CellReference cell) =>
    cell.Row is >= 1 and <= ExcelWorksheetLimits.MaximumRow &&
    cell.Column is >= 1 and <= ExcelWorksheetLimits.MaximumColumn;

  private static bool ApplicationMatches(object application, WorkbookIdentity identity)
  {
    var windowHandle = new IntPtr(Convert.ToInt64(
      GetRequiredProperty(application, "Hwnd"),
      CultureInfo.InvariantCulture));
    var threadId = NativeMethods.GetWindowThreadProcessId(windowHandle, out var processId);
    return threadId != 0 && processId == identity.ProcessId;
  }

  private static bool WorkbookMatches(object workbook, WorkbookIdentity identity)
  {
    var name = Convert.ToString(GetRequiredProperty(workbook, "Name"), CultureInfo.CurrentCulture);
    var fullName = Convert.ToString(GetRequiredProperty(workbook, "FullName"), CultureInfo.CurrentCulture);
    return string.Equals(name, identity.Name, StringComparison.OrdinalIgnoreCase) &&
      string.Equals(fullName, identity.FullPath, StringComparison.OrdinalIgnoreCase);
  }

  private static bool WorkbookWindowMatchesIdentity(object workbook, WorkbookIdentity identity)
  {
    if (identity.WindowSessionToken == IntPtr.Zero)
    {
      return false;
    }

    object? windows = null;
    object? window = null;
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

      var windowHandle = new IntPtr(Convert.ToInt64(
        GetRequiredProperty(window, "Hwnd"),
        CultureInfo.InvariantCulture));
      var desktopWindow = NativeMethods.FindWindowEx(windowHandle, IntPtr.Zero, "XLDESK", null);
      var documentWindowHandle = desktopWindow == IntPtr.Zero
        ? IntPtr.Zero
        : NativeMethods.FindWindowEx(desktopWindow, IntPtr.Zero, "EXCEL7", null);
      NativeMethods.GetWindowThreadProcessId(windowHandle, out var processId);
      return windowHandle == identity.ExcelWindowHandle &&
        documentWindowHandle == identity.ExcelDocumentWindowHandle &&
        processId == identity.ProcessId &&
        NativeMethods.GetProp(documentWindowHandle, WorkbookSessionTokenRegistry.WindowPropertyName) ==
          identity.WindowSessionToken;
    }
    finally
    {
      ComRelease.Release(window);
      ComRelease.Release(windows);
    }
  }

  private static void TryDeleteShape(object? shape)
  {
    if (shape is null)
    {
      return;
    }

    try
    {
      _ = InvokeMethod(shape, "Delete");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      // Best-effort compensation must not hide the placement failure.
    }
  }

  private static object GetRequiredProperty(object target, string propertyName, params object[] arguments) =>
    InvokeProperty(target, propertyName, arguments) ??
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

  private static object? InvokeMethod(object target, string methodName, params object[] arguments) =>
    target.GetType().InvokeMember(
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

  private static double ReadDoubleProperty(object target, string propertyName)
  {
    var value = Convert.ToDouble(GetRequiredProperty(target, propertyName), CultureInfo.InvariantCulture);
    return double.IsFinite(value) && value > 0 ? value : 64.0;
  }

  private static bool IsAutomationFailure(Exception exception) =>
    exception is COMException or TargetInvocationException or MissingMemberException or InvalidOperationException;

  private static int GetAutomationHResult(Exception exception) =>
    exception is TargetInvocationException { InnerException: not null } invocationException
      ? invocationException.InnerException!.HResult
      : exception.HResult;

  private static class NativeMethods
  {
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

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern nint GetProp(nint windowHandle, string propertyName);
  }
}

public sealed record ImagePlacementResult(
  bool Succeeded,
  bool FocusSucceeded,
  string ShapeName,
  string WorksheetName,
  CellReference FocusCell,
  ManagedShapeTarget? Target,
  string Message)
{
  public static ImagePlacementResult Failed(string message) =>
    new(false, false, string.Empty, string.Empty, default, null, message);
}

public sealed record ImageDeletionResult(bool Succeeded, string WorksheetName, string Message)
{
  public static ImageDeletionResult Failed(string message) => new(false, string.Empty, message);
}
