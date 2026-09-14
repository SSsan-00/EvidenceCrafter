using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Excel;

/// <summary>Captures the current worksheet state into immutable, COM-free records.</summary>
public sealed class ExcelSheetSnapshotService
{
  private const int MaximumSnapshotCells = 250_000;
  private const int NewFirstColumn = 3;
  private const int XlEdgeTop = 8;
  private const int XlEdgeBottom = 9;
  private const int XlEdgeRight = 10;
  private const int XlLineStyleNone = -4142;

  public SheetSnapshotResult Capture(
    WorkbookIdentity workbook,
    string worksheetName,
    int? scopeRow = null,
    bool includeWorksheetNames = false) =>
    CaptureCore(workbook, worksheetName, scopeRow, includeWorksheetNames, navigationOnly: false, includeShapes: true);

  public SheetSnapshotResult CaptureForNavigation(
    WorkbookIdentity workbook,
    string worksheetName,
    bool includeWorksheetNames = false,
    bool includeShapes = true) =>
    CaptureCore(workbook, worksheetName, scopeRow: null, includeWorksheetNames, navigationOnly: true, includeShapes);

  private static SheetSnapshotResult CaptureCore(
    WorkbookIdentity workbook,
    string worksheetName,
    int? scopeRow,
    bool includeWorksheetNames,
    bool navigationOnly,
    bool includeShapes)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);

    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return SheetSnapshotResult.Failed(worksheetName, "Excel snapshot capture must run on an STA thread.");
    }

    if (string.IsNullOrWhiteSpace(workbook.RotMonikerDisplayName) ||
      workbook.WindowSessionToken == IntPtr.Zero)
    {
      return SheetSnapshotResult.Failed(
        worksheetName,
        "Workbook connection is unavailable; refresh before analyzing the worksheet.");
    }

    IRunningObjectTable? runningObjectTable = null;
    IEnumMoniker? monikerEnumerator = null;
    IBindCtx? bindContext = null;
    var captureStage = "opening the Running Object Table";
    try
    {
      Marshal.ThrowExceptionForHR(NativeMethods.GetRunningObjectTable(0, out runningObjectTable));
      captureStage = "enumerating running Excel objects";
      runningObjectTable.EnumRunning(out monikerEnumerator);
      captureStage = "creating the COM binding context";
      Marshal.ThrowExceptionForHR(NativeMethods.CreateBindCtx(0, out bindContext));
      var monikers = new IMoniker[1];
      while (true)
      {
        captureStage = "advancing the Running Object Table enumerator";
        if (monikerEnumerator.Next(1, monikers, IntPtr.Zero) != 0)
        {
          break;
        }

        object? runningObject = null;
        try
        {
          string displayName;
          try
          {
            captureStage = "reading a running object name";
            monikers[0].GetDisplayName(bindContext, null, out displayName);
          }
          catch (Exception exception) when (IsAutomationFailure(exception))
          {
            // The ROT contains objects unrelated to the selected Excel instance.
            // A broken unrelated moniker must not abort capture of the target Workbook.
            continue;
          }

          if (!string.Equals(displayName, workbook.RotMonikerDisplayName, StringComparison.Ordinal))
          {
            continue;
          }

          try
          {
            captureStage = "binding the selected Workbook";
            runningObjectTable.GetObject(monikers[0], out runningObject);
            return runningObject is null
              ? SheetSnapshotResult.Failed(worksheetName, "The selected Workbook is no longer available.")
              : TryCaptureRunningObject(
                  runningObject,
                  workbook,
                  worksheetName,
                  scopeRow,
                  includeWorksheetNames,
                  navigationOnly,
                  includeShapes) ??
                SheetSnapshotResult.Failed(worksheetName, "The selected Workbook could not be matched.");
          }
          catch (Exception exception) when (IsAutomationFailure(exception))
          {
            return SheetSnapshotResult.Failed(
              worksheetName,
              $"Excel snapshot capture failed for the selected Workbook (0x{GetAutomationHResult(exception):X8}).");
          }
        }
        finally
        {
          ComRelease.Release(runningObject);
          ComRelease.Release(monikers[0]);
          monikers[0] = null!;
        }
      }

      return SheetSnapshotResult.Failed(
        worksheetName,
        "The selected Workbook is no longer available in the Running Object Table.");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return SheetSnapshotResult.Failed(
        worksheetName,
        $"Excel snapshot capture failed while {captureStage} (0x{GetAutomationHResult(exception):X8}).");
    }
    finally
    {
      ComRelease.Release(bindContext);
      ComRelease.Release(monikerEnumerator);
      ComRelease.Release(runningObjectTable);
    }
  }

  private static SheetSnapshotResult? TryCaptureRunningObject(
    object runningObject,
    WorkbookIdentity identity,
    string worksheetName,
    int? scopeRow,
    bool includeWorksheetNames,
    bool navigationOnly,
    bool includeShapes)
  {
    if (TryGetProperty(runningObject, "Workbooks", out var workbooks))
    {
      try
      {
        if (!ApplicationMatches(runningObject, identity))
        {
          return null;
        }

        var count = ReadInt(workbooks!, "Count");
        for (var index = 1; index <= count; index++)
        {
          object? candidate = null;
          try
          {
            candidate = InvokeProperty(workbooks!, "Item", index);
            if (candidate is not null && WorkbookMatches(candidate, identity))
            {
              return CaptureWorkbook(
                runningObject,
                candidate,
                identity,
                worksheetName,
                scopeRow,
                includeWorksheetNames,
                navigationOnly,
                includeShapes);
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
        ? CaptureWorkbook(
            application,
            runningObject,
            identity,
            worksheetName,
            scopeRow,
            includeWorksheetNames,
            navigationOnly,
            includeShapes)
        : null;
    }
    finally
    {
      ComRelease.Release(application);
    }
  }

  private static SheetSnapshotResult CaptureWorkbook(
    object application,
    object workbook,
    WorkbookIdentity identity,
    string worksheetName,
    int? scopeRow,
    bool includeWorksheetNames,
    bool navigationOnly,
    bool includeShapes)
  {
    if (!WorkbookWindowMatchesIdentity(workbook, identity))
    {
      return SheetSnapshotResult.Failed(
        worksheetName,
        "The selected Workbook was closed or reopened; refresh and select it again.");
    }

    object? worksheet = null;
    object? activeCell = null;
    object? workbookActiveSheet = null;
    object? windows = null;
    object? window = null;
    object? usedRange = null;
    object? usedRows = null;
    object? usedColumns = null;
    var captureStage = "validating the Workbook window";
    try
    {
      captureStage = "resolving the target worksheet";
      worksheet = ResolveWorksheet(application, workbook, worksheetName, out var resolvedName);
      captureStage = "activating the target worksheet";
      ActivateTargetWorksheet(application, workbook, worksheet);
      workbookActiveSheet = GetRequiredProperty(workbook, "ActiveSheet");
      var activeSheetName = Convert.ToString(
        GetRequiredProperty(workbookActiveSheet, "Name"),
        CultureInfo.CurrentCulture);
      if (!string.Equals(activeSheetName, resolvedName, StringComparison.Ordinal))
      {
        return SheetSnapshotResult.Failed(
          resolvedName,
          "対象SheetをExcelでアクティブにできませんでした。Sheetの表示状態を確認してください。");
      }

      captureStage = "reading the active cell and used range";
      windows = GetRequiredProperty(workbook, "Windows");
      window = GetRequiredProperty(windows, "Item", 1);
      activeCell = GetRequiredProperty(window, "RangeSelection");
      usedRange = GetRequiredProperty(worksheet, "UsedRange");
      usedRows = GetRequiredProperty(usedRange, "Rows");
      usedColumns = GetRequiredProperty(usedRange, "Columns");

      var firstRow = ReadInt(usedRange, "Row");
      var firstColumn = ReadInt(usedRange, "Column");
      var rowCount = ReadInt(usedRows, "Count");
      var columnCount = ReadInt(usedColumns, "Count");
      var lastRow = checked(firstRow + rowCount - 1);
      var lastColumn = checked(firstColumn + columnCount - 1);
      if (firstRow < 1 || firstColumn < 1 ||
        lastRow > ExcelWorksheetLimits.MaximumRow ||
        lastColumn > ExcelWorksheetLimits.MaximumColumn ||
        (long)rowCount * columnCount > MaximumSnapshotCells)
      {
        return SheetSnapshotResult.Failed(
          resolvedName,
          $"Worksheet UsedRange is too large or invalid for a safe snapshot ({rowCount:N0} x {columnCount:N0}).");
      }

      var activeReference = new CellReference(ReadInt(activeCell, "Row"), ReadInt(activeCell, "Column"));
      captureStage = "reading worksheet values and formulas";
      var values = InvokeProperty(usedRange, "Value2");
      var formulas = InvokeProperty(usedRange, "Formula");
      var usedRangeHasNoMerges = TryGetProperty(usedRange, "MergeCells", out var mergeCells) &&
        mergeCells is not null &&
        mergeCells is not DBNull &&
        !Convert.ToBoolean(mergeCells, CultureInfo.InvariantCulture);
      if (!usedRangeHasNoMerges)
      {
        return SheetSnapshotResult.Failed(
          resolvedName,
          $"シート {resolvedName} に結合セルがあるか、結合状態を確認できません。Evidenceシートでは結合セルを使用できません。");
      }
      var comments = navigationOnly ? [] : ReadLinkedCells(worksheet, "Comments");
      if (!navigationOnly)
      {
        comments.UnionWith(ReadLinkedCells(worksheet, "CommentsThreaded", optional: true));
      }
      var hyperlinks = navigationOnly ? [] : ReadLinkedCells(worksheet, "Hyperlinks");
      captureStage = "reading Case anchors";
      IReadOnlyList<CaseAnchorSignal> anchors = CaseAnchorNormalizer.Normalize(ReadAnchors(
        worksheet,
        values,
        formulas,
        firstRow,
        firstColumn,
        rowCount,
        columnCount,
        lastColumn));
      var firstAnchorRow = anchors.Count == 0 ? Math.Max(firstRow, 1) : anchors[0].Row;
      captureStage = "reading ordered Side headers";
      var sideHeaderColumns = ReadHeaderColumns(
        values, firstAnchorRow - 1, firstRow, firstColumn, rowCount, columnCount, lastColumn);
      var newHeaderColumns = sideHeaderColumns.Take(1).ToArray();
      var oldHeaderColumns = sideHeaderColumns.Skip(1).ToArray();
      captureStage = "reading vertical Case boundaries";
      var verticalBoundaries = ReadVerticalBoundaries(
        worksheet,
        firstAnchorRow,
        lastRow,
        lastColumn,
        oldHeaderColumns,
        anchors.Count == 0 ? activeReference.Row : anchors[^1].Row);
      captureStage = "reading the Case bottom boundary";
      var expectedLastEvidenceColumn = oldHeaderColumns.Length == 1
        ? checked((oldHeaderColumns[0] * 2) - NewFirstColumn - 1)
        : (int?)null;
      var horizontalBoundaries = ReadBottomBoundary(
        worksheet,
        lastRow,
        lastColumn,
        expectedLastEvidenceColumn);
      var observedLastColumn = lastColumn;
      var logicalLastRow = lastRow;

      var occupancyRow = scopeRow ?? activeReference.Row;
      var currentAnchor = anchors.LastOrDefault(anchor => anchor.Row <= occupancyRow);
      var nextAnchor = currentAnchor is null
        ? null
        : anchors.FirstOrDefault(anchor => anchor.Row > currentAnchor.Row);
      var currentCaseFirstRow = currentAnchor?.Row ?? firstAnchorRow;
      var currentCaseLastRow = Math.Min(nextAnchor?.Row - 1 ?? logicalLastRow, logicalLastRow);

      captureStage = "reading occupied cells";
      var cells = navigationOnly
        ? []
        : ReadOccupiedCells(
          values,
          formulas,
          comments,
          hyperlinks,
          firstRow,
          firstColumn,
          rowCount,
          columnCount,
          currentCaseFirstRow,
          currentCaseLastRow,
          observedLastColumn);
      captureStage = "reading worksheet Shapes";
      // Adjacent CASE navigation depends on structure, not image occupancy.
      // Placement callers always retain the complete, fresh shape snapshot.
      var shapes = includeShapes ? ReadShapes(worksheet) : [];
      captureStage = "reading row heights and column widths";
      var rowHeights = navigationOnly
        ? new Dictionary<int, double>()
        : ReadRowHeights(worksheet, currentCaseFirstRow, currentCaseLastRow);
      var columnWidths = navigationOnly
        ? new Dictionary<int, double>()
        : ReadColumnWidths(worksheet, NewFirstColumn, observedLastColumn);
      var signals = new SheetLayoutSignals(
        activeReference.Row,
        lastRow,
        logicalLastRow,
        NewFirstColumn,
        observedLastColumn,
        anchors,
        verticalBoundaries,
        horizontalBoundaries,
        oldHeaderColumns) { NewHeaderColumns = newHeaderColumns };

      if (!WorkbookWindowMatchesIdentity(workbook, identity))
      {
        return SheetSnapshotResult.Failed(
          resolvedName,
          "Workbook identity changed while the worksheet snapshot was being captured.");
      }

      var snapshot = new SheetSnapshot(
        resolvedName,
        activeReference,
        firstRow,
        lastRow,
        firstColumn,
        lastColumn,
        Convert.ToBoolean(GetRequiredProperty(workbook, "ReadOnly"), CultureInfo.InvariantCulture),
        IsWorksheetProtected(worksheet),
        signals,
        cells,
        shapes,
        rowHeights,
        columnWidths)
      {
        WorksheetNames = includeWorksheetNames ? ReadWorksheetNames(workbook) : [],
      };
      return new SheetSnapshotResult(true, snapshot, $"{resolvedName} のライブSnapshotを取得しました。");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return SheetSnapshotResult.Failed(
        worksheetName,
        $"Excel snapshot capture failed while {captureStage} (0x{GetAutomationHResult(exception):X8}).");
    }
    finally
    {
      ComRelease.Release(usedColumns);
      ComRelease.Release(usedRows);
      ComRelease.Release(usedRange);
      ComRelease.Release(activeCell);
      ComRelease.Release(window);
      ComRelease.Release(windows);
      ComRelease.Release(workbookActiveSheet);
      ComRelease.Release(worksheet);
    }
  }

  private static IReadOnlyList<CaseAnchorSignal> ReadAnchors(
    object worksheet,
    object? values,
    object? formulas,
    int firstRow,
    int firstColumn,
    int rowCount,
    int columnCount,
    int lastColumn)
  {
    var result = new List<CaseAnchorSignal>();
    for (var row = firstRow; row < firstRow + rowCount; row++)
    {
      var hasA = IsNonEmpty(MatrixValue(values, row, 1, firstRow, firstColumn, rowCount, columnCount)) ||
        IsNonEmpty(MatrixValue(formulas, row, 1, firstRow, firstColumn, rowCount, columnCount));
      var hasB = IsNonEmpty(MatrixValue(values, row, 2, firstRow, firstColumn, rowCount, columnCount)) ||
        IsNonEmpty(MatrixValue(formulas, row, 2, firstRow, firstColumn, rowCount, columnCount));
      if (!hasA && !hasB)
      {
        continue;
      }

      result.Add(new CaseAnchorSignal(
        row,
        hasA,
        hasB,
        false,
        DisplayValue(MatrixValue(values, row, 1, firstRow, firstColumn, rowCount, columnCount)),
        DisplayValue(MatrixValue(values, row, 2, firstRow, firstColumn, rowCount, columnCount))));
    }

    return result;
  }

  private static IReadOnlyList<int> ReadHeaderColumns(
    object? values,
    int headerRow,
    int firstRow,
    int firstColumn,
    int rowCount,
    int columnCount,
    int lastColumn)
  {
    if (headerRow < firstRow || headerRow >= firstRow + rowCount)
    {
      return [];
    }

    var result = new List<int>();
    for (var column = NewFirstColumn; column <= lastColumn; column++)
    {
      var label = DisplayValue(MatrixValue(values, headerRow, column, firstRow, firstColumn, rowCount, columnCount))?.Trim();
      if (!string.IsNullOrWhiteSpace(label))
      {
        result.Add(column);
      }
    }

    return result;
  }

  private static IReadOnlyList<VerticalBoundarySignal> ReadVerticalBoundaries(
    object worksheet,
    int firstAnchorRow,
    int lastRow,
    int lastColumn,
    IReadOnlyList<int> oldHeaderColumns,
    int minimumEndRow)
  {
    if (lastColumn <= NewFirstColumn || firstAnchorRow > lastRow)
    {
      return [];
    }

    var columns = oldHeaderColumns.Select(column => column - 1);
    var result = new List<VerticalBoundarySignal>();
    foreach (var column in columns.Distinct().Where(column => column >= NewFirstColumn && column < lastColumn))
    {
      if (HasRangeBorder(worksheet, firstAnchorRow, column, lastRow, column, XlEdgeRight))
      {
        if (lastRow >= minimumEndRow)
        {
          result.Add(new VerticalBoundarySignal(column, firstAnchorRow, lastRow));
        }
        continue;
      }

      // Partial borders are decorative; do not scan individual rows for placement.
    }

    return result;
  }

  private static IReadOnlyList<HorizontalBoundarySignal> ReadBottomBoundary(
    object worksheet,
    int row,
    int lastColumn,
    int? expectedLastColumn)
  {
    if (row < 1 || row > ExcelWorksheetLimits.MaximumRow ||
      !HasCellBorder(worksheet, row, 1, XlEdgeBottom))
    {
      return [];
    }

    if (expectedLastColumn is >= NewFirstColumn + 1 && expectedLastColumn <= lastColumn &&
      HasRangeBorder(worksheet, row, 1, row, expectedLastColumn.Value, XlEdgeBottom) &&
      (expectedLastColumn == lastColumn ||
        !HasCellBorder(worksheet, row, expectedLastColumn.Value + 1, XlEdgeBottom)))
    {
      return [new HorizontalBoundarySignal(row, 1, expectedLastColumn.Value)];
    }

    var endColumn = 1;
    while (endColumn < lastColumn && HasCellBorder(worksheet, row, endColumn + 1, XlEdgeBottom))
    {
      endColumn++;
    }

    return endColumn >= NewFirstColumn + 1
      ? [new HorizontalBoundarySignal(row, 1, endColumn)]
      : [];
  }

  private static IReadOnlyList<SnapshotCell> ReadOccupiedCells(
    object? values,
    object? formulas,
    HashSet<(int Row, int Column)> comments,
    HashSet<(int Row, int Column)> hyperlinks,
    int firstRow,
    int firstColumn,
    int rowCount,
    int columnCount,
    int caseFirstRow,
    int caseLastRow,
    int evidenceLastColumn)
  {
    var result = new List<SnapshotCell>();
    var startRow = Math.Max(firstRow, caseFirstRow);
    var endRow = Math.Min(firstRow + rowCount - 1, caseLastRow);
    var startColumn = Math.Max(firstColumn, NewFirstColumn);
    var endColumn = Math.Min(firstColumn + columnCount - 1, evidenceLastColumn);
    for (var row = startRow; row <= endRow; row++)
    {
      for (var column = startColumn; column <= endColumn; column++)
      {
        var hasContent = IsNonEmpty(
            MatrixValue(values, row, column, firstRow, firstColumn, rowCount, columnCount)) ||
          IsNonEmpty(MatrixValue(formulas, row, column, firstRow, firstColumn, rowCount, columnCount));
        var hasComment = comments.Contains((row, column));
        var hasHyperlink = hyperlinks.Contains((row, column));
        if (hasContent || hasComment || hasHyperlink)
        {
          result.Add(new SnapshotCell(row, column, hasContent, hasComment, hasHyperlink, false));
        }
      }
    }

    return result;
  }

  private static IReadOnlyList<SnapshotShape> ReadShapes(object worksheet)
  {
    object? shapes = null;
    var result = new List<SnapshotShape>();
    try
    {
      shapes = GetRequiredProperty(worksheet, "Shapes");
      var count = ReadInt(shapes, "Count");
      for (var index = 1; index <= count; index++)
      {
        object? shape = null;
        object? topLeft = null;
        object? bottomRight = null;
        try
        {
          shape = InvokeMethod(shapes, "Item", index);
          if (shape is null)
          {
            continue;
          }

          topLeft = GetRequiredProperty(shape, "TopLeftCell");
          bottomRight = GetRequiredProperty(shape, "BottomRightCell");
          var name = Convert.ToString(GetRequiredProperty(shape, "Name"), CultureInfo.CurrentCulture) ?? string.Empty;
          ManagedShapeMetadata? metadata = null;
          var isManaged = ManagedShapeMetadata.IsManagedName(name) &&
            TryGetProperty(shape, "AlternativeText", out var text) &&
            ManagedShapeMetadata.TryParse(Convert.ToString(text, CultureInfo.InvariantCulture), out metadata);
          var start = ReadShapeCellReference(topLeft);
          var end = ReadShapeCellReference(bottomRight);
          result.Add(new SnapshotShape(
            name,
            start.Row,
            end.Row,
            start.Column,
            end.Column,
            isManaged)
          {
            TopPoints = Convert.ToDouble(GetRequiredProperty(shape, "Top"), CultureInfo.InvariantCulture),
            WidthPoints = Convert.ToDouble(GetRequiredProperty(shape, "Width"), CultureInfo.InvariantCulture),
            HeightPoints = Convert.ToDouble(GetRequiredProperty(shape, "Height"), CultureInfo.InvariantCulture),
            HorizontalOffsetPoints = isManaged
              ? Convert.ToDouble(GetRequiredProperty(shape, "Left"), CultureInfo.InvariantCulture) -
                Convert.ToDouble(GetRequiredProperty(topLeft, "Left"), CultureInfo.InvariantCulture) : 0,
            SourceDimensions = metadata?.SourceDimensions,
          });
        }
        finally
        {
          ComRelease.Release(bottomRight);
          ComRelease.Release(topLeft);
          ComRelease.Release(shape);
        }
      }

      return result;
    }
    finally
    {
      ComRelease.Release(shapes);
    }
  }

  private static IReadOnlyDictionary<int, double> ReadRowHeights(
    object worksheet,
    int firstRow,
    int lastRow)
  {
    var result = new Dictionary<int, double>();
    ReadRowHeightBlock(worksheet, firstRow, lastRow, result);
    return result;
  }

  private static void ActivateTargetWorksheet(object application, object workbook, object worksheet)
  {
    var eventsWereEnabled = Convert.ToBoolean(
      GetRequiredProperty(application, "EnableEvents"),
      CultureInfo.InvariantCulture);
    SetProperty(application, "EnableEvents", false);
    try
    {
      InvokeMethod(workbook, "Activate");
      InvokeMethod(worksheet, "Activate");
    }
    finally
    {
      SetProperty(application, "EnableEvents", eventsWereEnabled);
    }
  }

  private static void ReadRowHeightBlock(
    object worksheet,
    int firstRow,
    int lastRow,
    Dictionary<int, double> result)
  {
    object? range = null;
    try
    {
      range = GetRequiredProperty(worksheet, "Range", $"A{firstRow}:A{lastRow}");
      var totalHeight = Convert.ToDouble(GetRequiredProperty(range, "Height"), CultureInfo.InvariantCulture);
      if (totalHeight == 0)
      {
        for (var row = firstRow; row <= lastRow; row++) result[row] = 0;
        return;
      }
      if (firstRow == lastRow)
      {
        if (double.IsFinite(totalHeight) && totalHeight > 0)
        {
          result[firstRow] = totalHeight;
        }
        return;
      }

      if (TryGetProperty(range, "RowHeight", out var uniformValue) &&
        uniformValue is not null && uniformValue is not DBNull)
      {
        var uniformHeight = Convert.ToDouble(uniformValue, CultureInfo.InvariantCulture);
        var rowCount = checked(lastRow - firstRow + 1);
        if (double.IsFinite(uniformHeight) && uniformHeight > 0 &&
          double.IsFinite(totalHeight) &&
          Math.Abs(totalHeight - (uniformHeight * rowCount)) <= Math.Max(0.01, rowCount * 0.001))
        {
          for (var row = firstRow; row <= lastRow; row++)
          {
            result[row] = uniformHeight;
          }
          return;
        }
      }
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      if (firstRow == lastRow)
      {
        object? cell = null;
        try
        {
          cell = GetRequiredProperty(worksheet, "Cells", firstRow, 1);
          var height = Convert.ToDouble(GetRequiredProperty(cell, "Height"), CultureInfo.InvariantCulture);
          if (double.IsFinite(height) && height >= 0) result[firstRow] = height;
        }
        finally { ComRelease.Release(cell); }
        return;
      }
    }
    finally
    {
      ComRelease.Release(range);
    }

    var middleRow = firstRow + ((lastRow - firstRow) / 2);
    ReadRowHeightBlock(worksheet, firstRow, middleRow, result);
    ReadRowHeightBlock(worksheet, middleRow + 1, lastRow, result);
  }

  private static IReadOnlyDictionary<int, double> ReadColumnWidths(
    object worksheet,
    int firstColumn,
    int lastColumn)
  {
    var result = new Dictionary<int, double>();
    object? range = null;
    object? firstCell = null;
    object? columns = null;
    try
    {
      range = GetRequiredProperty(worksheet, "Range", $"{ColumnName(firstColumn)}1:{ColumnName(lastColumn)}1");
      columns = GetRequiredProperty(range, "EntireColumn");
      // A one-row range can report its first column's width even when widths differ.
      // EntireColumn reports the mixed state needed to keep exact point widths.
      if (TryGetProperty(columns, "ColumnWidth", out var uniformWidth) &&
        uniformWidth is not null && uniformWidth is not DBNull &&
        Convert.ToDouble(uniformWidth, CultureInfo.InvariantCulture) > 0 &&
        TryGetProperty(columns, "Hidden", out var hidden) && hidden is false)
      {
        firstCell = GetRequiredProperty(worksheet, "Cells", 1, firstColumn);
        var width = Convert.ToDouble(GetRequiredProperty(firstCell, "Width"), CultureInfo.InvariantCulture);
        if (double.IsFinite(width) && width > 0)
        {
          for (var column = firstColumn; column <= lastColumn; column++)
          {
            result[column] = width;
          }
          return result;
        }
      }
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      // Bulk-property failures retain the existing per-column read below.
    }
    finally
    {
      ComRelease.Release(firstCell);
      ComRelease.Release(columns);
      ComRelease.Release(range);
    }

    for (var column = firstColumn; column <= lastColumn; column++)
    {
      object? cell = null;
      try
      {
        cell = GetRequiredProperty(worksheet, "Cells", 1, column);
        var width = Convert.ToDouble(GetRequiredProperty(cell, "Width"), CultureInfo.InvariantCulture);
        if (double.IsFinite(width) && width > 0)
        {
          result[column] = width;
        }
      }
      finally
      {
        ComRelease.Release(cell);
      }
    }

    return result;
  }

  private static HashSet<(int Row, int Column)> ReadLinkedCells(
    object worksheet,
    string collectionName,
    bool optional = false)
  {
    object? collection = null;
    var result = new HashSet<(int Row, int Column)>();
    try
    {
      if (optional && (!TryGetProperty(worksheet, collectionName, out collection) || collection is null))
      {
        return result;
      }

      collection ??= GetRequiredProperty(worksheet, collectionName);
      var count = ReadInt(collection, "Count");
      for (var index = 1; index <= count; index++)
      {
        object? item = null;
        object? parent = null;
        try
        {
          item = InvokeProperty(collection, "Item", index);
          if (item is null)
          {
            continue;
          }

          parent = GetRequiredProperty(item, "Parent");
          result.Add((ReadInt(parent, "Row"), ReadInt(parent, "Column")));
        }
        finally
        {
          ComRelease.Release(parent);
          ComRelease.Release(item);
        }
      }

      return result;
    }
    finally
    {
      ComRelease.Release(collection);
    }
  }

  private static bool HasRangeBorder(
    object worksheet,
    int firstRow,
    int firstColumn,
    int lastRow,
    int lastColumn,
    int borderIndex)
  {
    object? range = null;
    try
    {
      range = GetRequiredProperty(worksheet, "Range",
        $"{ColumnName(firstColumn)}{firstRow}:{ColumnName(lastColumn)}{lastRow}");
      return HasBorder(range, borderIndex);
    }
    finally
    {
      ComRelease.Release(range);
    }
  }

  internal static CellReference ReadShapeCellReference(object cell)
  {
    try
    {
      // Absolute R1C1 returns both coordinates in one COM call, independent of Excel's display style.
      if (InvokeProperty(cell, "Address", true, true, -4150, false) is string address &&
        address.StartsWith('R') && address.IndexOf('C') is var columnMarker && columnMarker > 1 &&
        int.TryParse(address.AsSpan(1, columnMarker - 1), NumberStyles.None, CultureInfo.InvariantCulture, out var row) &&
        int.TryParse(address.AsSpan(columnMarker + 1), NumberStyles.None, CultureInfo.InvariantCulture, out var column) &&
        row is >= 1 and <= ExcelWorksheetLimits.MaximumRow &&
        column is >= 1 and <= ExcelWorksheetLimits.MaximumColumn)
      {
        return new CellReference(row, column);
      }
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      // Older automation providers can reject the indexed Address property; retain the original reads.
    }
    return new CellReference(ReadInt(cell, "Row"), ReadInt(cell, "Column"));
  }

  private static string ColumnName(int column)
  {
    var name = string.Empty;
    while (column > 0)
    {
      column--;
      name = (char)('A' + column % 26) + name;
      column /= 26;
    }
    return name;
  }

  private static bool HasCellBorder(object worksheet, int row, int column, int borderIndex)
  {
    object? cell = null;
    try
    {
      cell = GetRequiredProperty(worksheet, "Cells", row, column);
      return HasBorder(cell, borderIndex);
    }
    finally
    {
      ComRelease.Release(cell);
    }
  }

  private static bool HasBorder(object range, int borderIndex)
  {
    object? border = null;
    try
    {
      border = InvokeProperty(range, "Borders", borderIndex);
      if (border is null || !TryGetProperty(border, "LineStyle", out var lineStyle) || lineStyle is null)
      {
        return false;
      }

      var value = Convert.ToInt32(lineStyle, CultureInfo.InvariantCulture);
      return value != 0 && value != XlLineStyleNone;
    }
    finally
    {
      ComRelease.Release(border);
    }
  }

  private static object? MatrixValue(
    object? matrix,
    int row,
    int column,
    int firstRow,
    int firstColumn,
    int rowCount,
    int columnCount)
  {
    if (row < firstRow || row >= firstRow + rowCount ||
      column < firstColumn || column >= firstColumn + columnCount ||
      matrix is null)
    {
      return null;
    }

    if (matrix is not Array array)
    {
      return rowCount == 1 && columnCount == 1 ? matrix : null;
    }

    return array.GetValue(
      array.GetLowerBound(0) + row - firstRow,
      array.GetLowerBound(1) + column - firstColumn);
  }

  private static bool IsNonEmpty(object? value) =>
    value is not null &&
    value is not DBNull &&
    (value is not string text || !string.IsNullOrWhiteSpace(text));

  private static string? DisplayValue(object? value)
  {
    var text = Convert.ToString(value, CultureInfo.CurrentCulture)?.Trim();
    if (string.IsNullOrEmpty(text))
    {
      return null;
    }

    text = string.Join(" ", text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    return text.Length <= 60 ? text : $"{text[..57]}...";
  }

  private static IReadOnlyList<string> ReadWorksheetNames(object workbook)
  {
    object? worksheets = null;
    var names = new List<string>();
    try
    {
      worksheets = GetRequiredProperty(workbook, "Worksheets");
      var count = ReadInt(worksheets, "Count");
      for (var index = 1; index <= count; index++)
      {
        object? worksheet = null;
        try
        {
          worksheet = InvokeProperty(worksheets, "Item", index);
          var name = Convert.ToString(GetRequiredProperty(worksheet!, "Name"), CultureInfo.CurrentCulture);
          if (!string.IsNullOrWhiteSpace(name))
          {
            names.Add(name);
          }
        }
        finally
        {
          ComRelease.Release(worksheet);
        }
      }
      return names;
    }
    finally
    {
      ComRelease.Release(worksheets);
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
      var activeSheet = GetRequiredProperty(workbook, "ActiveSheet");
      try
      {
        resolvedName = Convert.ToString(
            GetRequiredProperty(activeSheet, "Name"),
            CultureInfo.CurrentCulture) ??
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
        resolvedName = Convert.ToString(
            GetRequiredProperty(worksheet, "Name"),
            CultureInfo.CurrentCulture) ??
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

  private static bool IsWorksheetProtected(object worksheet) =>
    Convert.ToBoolean(GetRequiredProperty(worksheet, "ProtectContents"), CultureInfo.InvariantCulture) ||
    Convert.ToBoolean(GetRequiredProperty(worksheet, "ProtectDrawingObjects"), CultureInfo.InvariantCulture) ||
    Convert.ToBoolean(GetRequiredProperty(worksheet, "ProtectScenarios"), CultureInfo.InvariantCulture);

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
      if (ReadInt(windows, "Count") < 1)
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

  private static int ReadInt(object target, string propertyName) =>
    Convert.ToInt32(GetRequiredProperty(target, propertyName), CultureInfo.InvariantCulture);

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

  private static void SetProperty(object target, string propertyName, object value) =>
    target.GetType().InvokeMember(
      propertyName,
      BindingFlags.SetProperty,
      binder: null,
      target,
      [value],
      CultureInfo.CurrentCulture);

  private static object? InvokeMethod(object target, string methodName, params object[] arguments) =>
    target.GetType().InvokeMember(
      methodName,
      BindingFlags.InvokeMethod,
      binder: null,
      target,
      arguments,
      CultureInfo.CurrentCulture);

  private static bool IsAutomationFailure(Exception exception) =>
    exception is COMException or TargetInvocationException or MissingMemberException or InvalidOperationException;

  private static int GetAutomationHResult(Exception exception)
  {
    while (exception is TargetInvocationException { InnerException: not null })
    {
      exception = exception.InnerException;
    }
    return exception.HResult;
  }

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

public sealed record SheetSnapshotResult(
  bool Succeeded,
  SheetSnapshot? Snapshot,
  string Message)
{
  public static SheetSnapshotResult Failed(string worksheetName, string message) =>
    new(false, null, $"{worksheetName}: {message}");
}
