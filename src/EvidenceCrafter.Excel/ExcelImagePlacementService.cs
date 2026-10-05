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
  [ThreadStatic] internal static Action<string, object>? PlacementStageObserved;
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
    double? scaleOverride = null,
    double verticalOffsetPoints = PlacementPlanner.VerticalInsetPoints)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ArgumentException.ThrowIfNullOrWhiteSpace(imagePath);
    if (!double.IsFinite(verticalOffsetPoints) || verticalOffsetPoints < 0)
      throw new ArgumentOutOfRangeException(nameof(verticalOffsetPoints));
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
            scaleOverride, verticalOffsetPoints);
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
    double? scaleOverride,
    double verticalOffsetPoints)
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
                scaleOverride, verticalOffsetPoints);
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
          scaleOverride, verticalOffsetPoints)
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
    double? scaleOverride,
    double verticalOffsetPoints)
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
    var insertionAttempted = false;
    ImagePlacementGeometry? expected = null;
    ImagePlacementGeometry? actual = null;
    double? appliedScale = null;
    var stage = "BeforeInsert";
    ImagePlacementResult Fail(string message)
    {
      var diagnostic = CapturePlacementDiagnostic(application, targetCell, focusCell, side,
        imageDimensions, appliedScale, stage, expected, actual);
      var restored = !insertionAttempted || TryDeleteShape(shape, shapes);
      return ImagePlacementResult.Failed(restored ? message :
        message + " 挿入画像の取り消しを確認できません。追加の操作を止め、保存せずExcelの状態を確認してください。") with
      {
        MutationMayHaveOccurred = !restored,
        Diagnostic = diagnostic,
      };
    }
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
      if (cellWidth <= 0)
        return Fail("配置セルの列が非表示、または幅が0のため配置できません。配置先の列を確認してください。");
      var availableWidth = availableWidthPoints ?? Math.Max(cellWidth * 5.0 - (horizontalMarginPoints * 2), 24.0);
      var fittedImage = scaleOverride is { } scale
        ? imageSizingService.AtScale(imageDimensions, availableWidth, scale)
        : imageSizingService.FitToWidth(imageDimensions, availableWidth);
      expected = new(cellLeft + horizontalMarginPoints, cellTop + verticalOffsetPoints,
        fittedImage.WidthPoints, fittedImage.HeightPoints);
      appliedScale = fittedImage.Scale;
      if (!expected.IsValid)
        return Fail("画像の予定位置またはサイズが不正なため配置できません。");

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
      stage = "AddPicture";
      PlacementStageObserved?.Invoke("BeforeInsert", shapes);
      insertionAttempted = true;
      shape = InvokeMethod(
        shapes,
        "AddPicture",
        imagePath,
        MsoFalse,
        MsoTrue,
        cellLeft + horizontalMarginPoints,
        cellTop + verticalOffsetPoints,
        fittedImage.WidthPoints,
        fittedImage.HeightPoints);
      if (shape is null)
      {
        return Fail("Excelから挿入画像を取得できませんでした。");
      }

      PlacementStageObserved?.Invoke("AfterInsert", shape);
      stage = "Attributes";

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
      PlacementStageObserved?.Invoke("AfterAttributes", shape);

      stage = "ReadShape";
      var insertedName = Convert.ToString(GetRequiredProperty(shape, "Name"), CultureInfo.InvariantCulture);
      actual = new(ReadDoubleProperty(shape, "Left"), ReadDoubleProperty(shape, "Top"),
        ReadDoubleProperty(shape, "Width"), ReadDoubleProperty(shape, "Height"));
      stage = "VerifyName";
      if (!string.Equals(insertedName, shapeName, StringComparison.Ordinal))
        return Fail("挿入画像の管理名を確認できませんでした。");
      stage = "VerifyGeometry";
      if (VerifyGeometry(expected, actual) is { } failure) return Fail(failure);
      if (PlacementGeometryComparison.Exceeds(actual.Width, availableWidth))
        return Fail("挿入画像の幅が配置可能幅を超えているため配置できません。");

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
          actual.Left,
          actual.Top,
          actual.Width,
          actual.Height),
        focus.Succeeded
          ? $"画像を{resolvedWorksheetName}!R{focusCell.Value.Row}C{focusCell.Value.Column}へ配置しました。"
          : $"画像を配置しましたが、対象セルの選択に失敗しました: {focus.Message}");
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      return Fail($"Excelへの画像配置に失敗しました（工程: {stage}、0x{GetAutomationHResult(exception):X8}）。");
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

  private static bool TryDeleteShape(object? shape, object? shapes)
  {
    if (shape is null || shapes is null) return false;

    try
    {
      var count = Convert.ToInt32(GetRequiredProperty(shapes, "Count"), CultureInfo.InvariantCulture);
      PlacementStageObserved?.Invoke("BeforeRollback", shape);
      _ = InvokeMethod(shape, "Delete");
      PlacementStageObserved?.Invoke("AfterRollback", shape);
      return Convert.ToInt32(GetRequiredProperty(shapes, "Count"), CultureInfo.InvariantCulture) == count - 1;
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      return false;
    }
  }

  internal static bool PositionMatches(double expected, double actual) =>
    float.IsFinite((float)expected) && float.IsFinite((float)actual) && expected >= 0 && actual >= 0 &&
    PlacementGeometryComparison.Matches(expected, actual);

  internal static string? VerifyGeometry(ImagePlacementGeometry expected, ImagePlacementGeometry actual)
  {
    foreach (var (property, planned, measured, position) in new[]
    {
      ("Left", expected.Left, actual.Left, true), ("Top", expected.Top, actual.Top, true),
      ("Width", expected.Width, actual.Width, false), ("Height", expected.Height, actual.Height, false),
    })
    {
      var valid = position ? PositionMatches(planned, measured) :
        double.IsFinite(planned) && double.IsFinite(measured) && planned > 0 && measured > 0 && PlacementGeometryComparison.Matches(planned, measured);
      if (!valid)
        return FormattableString.Invariant($"{(position ? "挿入画像の位置を確認できませんでした" : "挿入画像のサイズが予定値と一致しませんでした")}（{property}: 予定 {planned:R}pt、実際 {measured:R}pt、差 {measured - planned:R}pt）。");
    }
    return null;
  }

  private static ImagePlacementDiagnostic CapturePlacementDiagnostic(object application, object? cell,
    CellReference? reference, EvidenceSide side, ImageDimensions source, double? scale, string stage,
    ImagePlacementGeometry? expected, ImagePlacementGeometry? actual)
  {
    string? version = null, build = null;
    bool? rowHidden = null, columnHidden = null;
    object? row = null, column = null;
    try
    {
      if (TryGetProperty(application, "Version", out var value)) version = Convert.ToString(value, CultureInfo.InvariantCulture);
      if (TryGetProperty(application, "Build", out value)) build = Convert.ToString(value, CultureInfo.InvariantCulture);
      if (cell is not null)
      {
        row = GetRequiredProperty(cell, "EntireRow");
        column = GetRequiredProperty(cell, "EntireColumn");
        if (TryGetProperty(row, "Hidden", out value) && value is bool hiddenRow) rowHidden = hiddenRow;
        if (TryGetProperty(column, "Hidden", out value) && value is bool hiddenColumn) columnHidden = hiddenColumn;
      }
    }
    catch (Exception exception) when (exception is not OutOfMemoryException) { }
    finally { ComRelease.Release(row); ComRelease.Release(column); }
    return new(stage, version, build, reference, side, source, scale, expected, actual, rowHidden, columnHidden);
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
    return value;
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
  public bool MutationMayHaveOccurred { get; init; }
  public ImagePlacementDiagnostic? Diagnostic { get; init; }
  public static ImagePlacementResult Failed(string message) =>
    new(false, false, string.Empty, string.Empty, default, null, message);
}

public sealed record ImagePlacementGeometry(double Left, double Top, double Width, double Height)
{
  internal bool IsValid => float.IsFinite((float)Left) && Left >= 0 && float.IsFinite((float)Top) && Top >= 0 &&
    float.IsFinite((float)Width) && Width > 0 && float.IsFinite((float)Height) && Height > 0;
}

public sealed record ImagePlacementDiagnostic(string Stage, string? ExcelVersion, string? ExcelBuild,
  CellReference? Cell, EvidenceSide Side, ImageDimensions SourceDimensions, double? Scale,
  ImagePlacementGeometry? Expected, ImagePlacementGeometry? Actual, bool? RowHidden, bool? ColumnHidden)
{
  public double ComparisonTolerancePoints => PlacementGeometryComparison.TolerancePoints;
  public double ComparisonPrecisionPoints => PlacementGeometryComparison.PrecisionPoints;
  public double? ExpectedLeftRounded => Expected is null ? null : PlacementGeometryComparison.Round(Expected.Left);
  public double? ExpectedTopRounded => Expected is null ? null : PlacementGeometryComparison.Round(Expected.Top);
  public double? ActualLeftRounded => Actual is null ? null : PlacementGeometryComparison.Round(Actual.Left);
  public double? ActualTopRounded => Actual is null ? null : PlacementGeometryComparison.Round(Actual.Top);
  public bool? TopMatchesAllowedPosition => Expected is null || Actual is null ? null :
    ExcelImagePlacementService.PositionMatches(Expected.Top, Actual.Top);
}

public sealed record ImageDeletionResult(bool Succeeded, string WorksheetName, string Message)
{
  public static ImageDeletionResult Failed(string message) => new(false, string.Empty, message);
}
