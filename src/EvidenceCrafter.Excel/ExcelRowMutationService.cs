using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using System.Text;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Excel;

/// <summary>
/// Applies the small, verified row-mutation slice to an open Excel Workbook.
/// The service never saves the Workbook and never mutates rows unless the caller's
/// immutable safety snapshot produces a fully contiguous trailing deletion plan.
/// </summary>
public sealed class ExcelRowMutationService
{
  private const int MaximumDependencyCells = 250_000;
  private readonly TrailingRowDeletionPlanner deletionPlanner;

  public ExcelRowMutationService(TrailingRowDeletionPlanner? deletionPlanner = null)
  {
    this.deletionPlanner = deletionPlanner ?? new TrailingRowDeletionPlanner();
  }

  /// <summary>
  /// Reads a conservative live safety snapshot for a row interval.
  /// Callers can pass the returned rows to <see cref="DeleteTrailingRows"/>.
  /// The snapshot is intentionally separate from deletion so the UI can show the
  /// planned range before applying it; the delete call repeats all identity checks.
  /// </summary>
  public RowSafetySnapshotResult CaptureRowSafetyStates(
    WorkbookIdentity workbook,
    string worksheetName,
    int firstRow,
    int lastRow)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ValidateRowRange(firstRow, lastRow);

    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return RowSafetySnapshotResult.Failed(
        worksheetName,
        "Excel row safety inspection must run on an STA thread.");
    }

    if (string.IsNullOrWhiteSpace(workbook.RotMonikerDisplayName) ||
      workbook.WindowSessionToken == IntPtr.Zero)
    {
      return RowSafetySnapshotResult.Failed(
        worksheetName,
        "Workbook connection is unavailable; refresh before inspecting rows.");
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
            return RowSafetySnapshotResult.Failed(
              worksheetName,
              "The selected Workbook is no longer available in its Excel instance.");
          }

          return TryCaptureRunningObject(
              runningObject,
              workbook,
              worksheetName,
              firstRow,
              lastRow) ??
            RowSafetySnapshotResult.Failed(
              worksheetName,
              "The selected Workbook could not be matched in its Excel instance.");
        }
        catch (Exception exception) when (IsAutomationFailure(exception))
        {
          return RowSafetySnapshotResult.Failed(
            worksheetName,
            $"Excel row safety inspection failed (0x{GetAutomationHResult(exception):X8}).");
        }
        finally
        {
          ComRelease.Release(runningObject);
          ComRelease.Release(monikers[0]);
          monikers[0] = null!;
        }
      }

      return RowSafetySnapshotResult.Failed(
        worksheetName,
        "The selected Workbook is no longer available in the Running Object Table.");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return RowSafetySnapshotResult.Failed(
        worksheetName,
        $"Excel row safety inspection failed (0x{GetAutomationHResult(exception):X8}).");
    }
    finally
    {
      ComRelease.Release(bindContext);
      ComRelease.Release(monikerEnumerator);
      ComRelease.Release(runningObjectTable);
    }
  }

  /// <summary>Inserts a contiguous row range at the requested row.</summary>
  public RowMutationResult InsertRows(
    WorkbookIdentity workbook,
    string worksheetName,
    RowInsertion insertion)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ValidateInsertion(insertion);

    return Execute(
      workbook,
      worksheetName,
      new RowMutationPlan(
        RowMutationOperation.Insert,
        insertion.AtRow,
        insertion.Count,
        insertion.Reason,
        UseActiveCell: false));
  }

  /// <summary>Normalizes only the rows just inserted by this application.</summary>
  public RowMutationResult NormalizeInsertedRows(
    WorkbookIdentity workbook,
    string worksheetName,
    int startRow,
    int count,
    double rowHeightPoints)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ValidateInsertion(startRow, count);
    if (!double.IsFinite(rowHeightPoints) || rowHeightPoints <= 0 || rowHeightPoints > 409.5)
      throw new ArgumentOutOfRangeException(nameof(rowHeightPoints));

    return Execute(
      workbook,
      worksheetName,
      new RowMutationPlan(
        RowMutationOperation.Insert,
        startRow,
        count,
        "Normalize the newly inserted rows for automatic image placement.",
        UseActiveCell: false,
        InsertedRowHeightPoints: rowHeightPoints));
  }

  /// <summary>
  /// Inserts rows immediately above the active cell on the selected worksheet.
  /// The active row is read only after the Workbook and Worksheet identity checks
  /// have passed, so a stale selection cannot redirect the mutation.
  /// </summary>
  public RowMutationResult InsertRowsAtActiveCell(
    WorkbookIdentity workbook,
    string worksheetName,
    int count,
    string reason = "Insert rows above the selected active cell.")
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ValidateInsertionCount(count);

    return Execute(
      workbook,
      worksheetName,
      new RowMutationPlan(
        RowMutationOperation.Insert,
        StartRow: 0,
        count,
        reason,
        UseActiveCell: true));
  }

  /// <summary>
  /// Plans and deletes only a fully safe, contiguous trailing row range.
  /// Missing row safety observations are treated as unsafe by the planner.
  /// </summary>
  public RowMutationResult DeleteTrailingRows(
    WorkbookIdentity workbook,
    string worksheetName,
    int caseStartRow,
    int caseEndRow,
    int lastContentRow,
    int tailRows,
    IReadOnlyList<RowSafetyState> rows)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ArgumentNullException.ThrowIfNull(rows);

    var plannedRows = deletionPlanner.Plan(
      caseStartRow,
      caseEndRow,
      lastContentRow,
      tailRows,
      rows);
    if (plannedRows.Count == 0)
    {
      return RowMutationResult.NoChange(
        RowMutationOperation.Delete,
        worksheetName,
        "削除可能な安全な末尾行はありません。シートは変更していません。");
    }

    ValidateTrailingPlan(plannedRows, caseStartRow, caseEndRow);
    return Execute(
      workbook,
      worksheetName,
      new RowMutationPlan(
        RowMutationOperation.Delete,
        plannedRows[0],
        plannedRows.Count,
        "Delete the fully safe contiguous trailing rows selected by TrailingRowDeletionPlanner.",
        UseActiveCell: false));
  }

  /// <summary>
  /// Deletes the live Case tail and returns an application-owned native Excel snapshot.
  /// Disposing the snapshot removes its temporary workbook.
  /// </summary>
  public RowMutationResult DeleteTrailingRowsWithSnapshot(
    WorkbookIdentity workbook,
    string worksheetName,
    int caseStartRow,
    int caseEndRow,
    int tailRows = 4)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ValidateCaseRange(caseStartRow, caseEndRow);
    if (tailRows is < 2 or > ExcelWorksheetLimits.MaximumRow)
    {
      throw new ArgumentOutOfRangeException(nameof(tailRows));
    }

    return Execute(
      workbook,
      worksheetName,
      new RowMutationPlan(
        RowMutationOperation.Delete,
        StartRow: 0,
        Count: 0,
        "Delete a live Case tail with an exact native Excel undo snapshot.",
        UseActiveCell: false,
        DeletionRequest: new TrailingDeletionRequest(caseStartRow, caseEndRow, tailRows),
        CaptureUndoSnapshot: true));
  }

  /// <summary>Restores rows and their native Excel formatting from a deletion snapshot.</summary>
  public RowMutationResult RestoreDeletedRows(WorkbookIdentity workbook, RowDeletionSnapshot snapshot)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentNullException.ThrowIfNull(snapshot);
    if (snapshot.IsDisposed || !File.Exists(snapshot.BackupPath))
    {
      return RowMutationResult.Failed(RowMutationOperation.Insert, snapshot.WorksheetName, "Undo用の行Snapshotは利用できません。");
    }

    if (!snapshot.Matches(workbook))
    {
      return RowMutationResult.Failed(RowMutationOperation.Insert, snapshot.WorksheetName, "Workbookの接続世代が変わったため行削除をUndoできません。");
    }

    return Execute(
      workbook,
      snapshot.WorksheetName,
      new RowMutationPlan(
        RowMutationOperation.Insert,
        snapshot.StartRow,
        snapshot.Count,
        "Restore deleted rows from a native Excel snapshot.",
        UseActiveCell: false,
        RestoreSnapshot: snapshot));
  }

  /// <summary>Deletes a previously restored snapshot range after the normal live safety check.</summary>
  public RowMutationResult DeleteRestoredRows(WorkbookIdentity workbook, RowDeletionSnapshot snapshot)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentNullException.ThrowIfNull(snapshot);
    if (snapshot.IsDisposed || !snapshot.Matches(workbook))
    {
      return RowMutationResult.Failed(RowMutationOperation.Delete, snapshot.WorksheetName, "行SnapshotまたはWorkbook接続世代が無効なためRedoできません。");
    }

    return Execute(
      workbook,
      snapshot.WorksheetName,
      new RowMutationPlan(
        RowMutationOperation.Delete,
        snapshot.StartRow,
        snapshot.Count,
        "Redo a snapshotted row deletion after a live safety check.",
        UseActiveCell: false,
        ExactSafeDeletion: true,
        RestoreSnapshot: snapshot));
  }

  /// <summary>
  /// Reads the selected Case from Excel immediately before mutation and deletes only its
  /// fully safe contiguous trailing rows. A failed or incomplete safety read prevents deletion.
  /// </summary>
  public RowMutationResult DeleteTrailingRows(
    WorkbookIdentity workbook,
    string worksheetName,
    int caseStartRow,
    int caseEndRow,
    int tailRows = 4)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ValidateCaseRange(caseStartRow, caseEndRow);
    if (tailRows is < 2 or > ExcelWorksheetLimits.MaximumRow)
    {
      throw new ArgumentOutOfRangeException(nameof(tailRows));
    }

    return Execute(
      workbook,
      worksheetName,
      new RowMutationPlan(
        RowMutationOperation.Delete,
        StartRow: 0,
        Count: 0,
        "Delete only the fully safe contiguous trailing rows from a live Case snapshot.",
        UseActiveCell: false,
        DeletionRequest: new TrailingDeletionRequest(caseStartRow, caseEndRow, tailRows)));
  }

  /// <summary>
  /// Deletes an exact row range only when a fresh live snapshot says every row is empty and structurally safe.
  /// This is used to reverse a row insertion without deleting content added after the insertion.
  /// </summary>
  public RowMutationResult DeleteRowsIfSafe(
    WorkbookIdentity workbook,
    string worksheetName,
    int startRow,
    int count)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ValidateInsertion(startRow, count);

    return Execute(
      workbook,
      worksheetName,
      new RowMutationPlan(
        RowMutationOperation.Delete,
        startRow,
        count,
        "Delete an exact range only after a live safety check.",
        UseActiveCell: false,
        ExactSafeDeletion: true));
  }

  private RowMutationResult Execute(
    WorkbookIdentity workbook,
    string worksheetName,
    RowMutationPlan plan)
  {
    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return RowMutationResult.Failed(
        plan.Operation,
        worksheetName,
        "Excel row mutation must run on an STA thread.");
    }

    if (workbook.IsReadOnly)
    {
      return RowMutationResult.Failed(
        plan.Operation,
        worksheetName,
        "The selected Workbook is read-only.");
    }

    if (string.IsNullOrWhiteSpace(workbook.RotMonikerDisplayName))
    {
      return RowMutationResult.Failed(
        plan.Operation,
        worksheetName,
        "The Workbook does not have a verifiable open session; refresh and select it again.");
    }

    if (workbook.WindowSessionToken == IntPtr.Zero)
    {
      return RowMutationResult.Failed(
        plan.Operation,
        worksheetName,
        "Workbook connection is unavailable; refresh before changing rows.");
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
            return RowMutationResult.Failed(
              plan.Operation,
              worksheetName,
              "The selected Workbook is no longer available in its Excel instance.");
          }

          return TryMutateRunningObject(runningObject, workbook, worksheetName, plan) ??
            RowMutationResult.Failed(
              plan.Operation,
              worksheetName,
              "The selected Workbook could not be matched in its Excel instance.");
        }
        catch (Exception exception) when (IsAutomationFailure(exception))
        {
          return RowMutationResult.Failed(
            plan.Operation,
            worksheetName,
            $"Excel row mutation failed (0x{GetAutomationHResult(exception):X8}).");
        }
        finally
        {
          ComRelease.Release(runningObject);
          ComRelease.Release(monikers[0]);
          monikers[0] = null!;
        }
      }

      return RowMutationResult.Failed(
        plan.Operation,
        worksheetName,
        "The selected Workbook is no longer available in the Running Object Table.");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return RowMutationResult.Failed(
        plan.Operation,
        worksheetName,
        $"Excel row mutation failed (0x{GetAutomationHResult(exception):X8}).");
    }
    finally
    {
      ComRelease.Release(bindContext);
      ComRelease.Release(monikerEnumerator);
      ComRelease.Release(runningObjectTable);
    }
  }

  private RowMutationResult? TryMutateRunningObject(
    object runningObject,
    WorkbookIdentity identity,
    string worksheetName,
    RowMutationPlan plan)
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
              return MutateWorkbook(runningObject, candidate, identity, worksheetName, plan);
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
        ? MutateWorkbook(application, runningObject, identity, worksheetName, plan)
        : null;
    }
    finally
    {
      ComRelease.Release(application);
    }
  }

  private static RowSafetySnapshotResult? TryCaptureRunningObject(
    object runningObject,
    WorkbookIdentity identity,
    string worksheetName,
    int firstRow,
    int lastRow)
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
              return CaptureWorkbookRows(
                runningObject,
                candidate,
                identity,
                worksheetName,
                firstRow,
                lastRow);
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
        ? CaptureWorkbookRows(
          application,
          runningObject,
          identity,
          worksheetName,
          firstRow,
          lastRow)
        : null;
    }
    finally
    {
      ComRelease.Release(application);
    }
  }

  private static RowSafetySnapshotResult CaptureWorkbookRows(
    object application,
    object workbook,
    WorkbookIdentity identity,
    string worksheetName,
    int firstRow,
    int lastRow)
  {
    if (!WorkbookWindowMatchesIdentity(workbook, identity))
    {
      return RowSafetySnapshotResult.Failed(
        worksheetName,
        "The selected Workbook was closed or reopened; refresh and select it again.");
    }

    object? worksheet = null;
    object? rows = null;
    try
    {
      InvokeMethod(workbook, "Activate");
      worksheet = ResolveWorksheet(application, workbook, worksheetName, out var resolvedWorksheetName);
      if (IsWorksheetProtected(worksheet))
      {
        return RowSafetySnapshotResult.Failed(
          resolvedWorksheetName,
          $"シート {resolvedWorksheetName} は保護されています。保護を解除してから行を確認してください。");
      }

      if (!WorkbookWindowMatchesIdentity(workbook, identity) ||
        Convert.ToBoolean(GetRequiredProperty(workbook, "ReadOnly"), CultureInfo.InvariantCulture))
      {
        return RowSafetySnapshotResult.Failed(
          resolvedWorksheetName,
          "The Workbook session changed or became read-only; refresh before inspecting rows.");
      }

      var states = CaptureRowSafetySnapshot(worksheet, firstRow, lastRow, out _);

      return new RowSafetySnapshotResult(
        true,
        resolvedWorksheetName,
        states,
        "Live row safety snapshot captured. Reconfirm the Workbook before deleting.");
    }
    finally
    {
      ComRelease.Release(rows);
      ComRelease.Release(worksheet);
    }
  }

  private static IReadOnlyList<RowSafetyState> CaptureRowSafetySnapshot(
    object worksheet,
    int firstRow,
    int lastRow,
    out int lastContentRow)
  {
    ValidateRowRange(firstRow, lastRow);
    var commentRows = ReadRowsWithComments(worksheet, firstRow, lastRow);
    var hyperlinkRows = ReadRowsWithHyperlinks(worksheet, firstRow, lastRow);
    var shapeRows = ReadRowsWithShapes(worksheet, firstRow, lastRow);
    var rangeHasNoMerges = RangeHasNoMerges(worksheet, firstRow, lastRow);
    object? rows = null;
    try
    {
      rows = GetRequiredProperty(worksheet, "Rows");
      var states = new List<RowSafetyState>(lastRow - firstRow + 1);
      lastContentRow = 0;
      for (var row = firstRow; row <= lastRow; row++)
      {
        object? rowRange = null;
        try
        {
          rowRange = InvokeProperty(rows, "Item", row) ??
            throw new InvalidOperationException("The requested Excel row could not be resolved.");
          var state = new RowSafetyState(
            row,
            HasAnyCellContent(rowRange),
            commentRows.Contains(row),
            hyperlinkRows.Contains(row),
            shapeRows.Contains(row),
            !rangeHasNoMerges && HasMergeOrUnknown(rowRange));
          states.Add(state);
          if (!state.CanDelete)
          {
            lastContentRow = row;
          }
        }
        finally
        {
          ComRelease.Release(rowRange);
        }
      }

      return states;
    }
    finally
    {
      ComRelease.Release(rows);
    }
  }

  private static HashSet<int> ReadRowsWithComments(object worksheet, int firstRow, int lastRow)
  {
    var result = new HashSet<int>();
    object? comments = null;
    try
    {
      comments = GetRequiredProperty(worksheet, "Comments");
      var count = Convert.ToInt32(GetRequiredProperty(comments, "Count"), CultureInfo.InvariantCulture);
      for (var index = 1; index <= count; index++)
      {
        object? comment = null;
        object? parent = null;
        try
        {
          comment = InvokeProperty(comments, "Item", index) ??
            throw new InvalidOperationException("The worksheet comment could not be resolved.");
          parent = GetRequiredProperty(comment, "Parent");
          var row = Convert.ToInt32(GetRequiredProperty(parent, "Row"), CultureInfo.InvariantCulture);
          if (row >= firstRow && row <= lastRow)
          {
            result.Add(row);
          }
        }
        finally
        {
          ComRelease.Release(parent);
          ComRelease.Release(comment);
        }
      }

      return result;
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return AllRows(firstRow, lastRow);
    }
    finally
    {
      ComRelease.Release(comments);
    }
  }

  private static HashSet<int> ReadRowsWithHyperlinks(object worksheet, int firstRow, int lastRow)
  {
    var result = new HashSet<int>();
    object? hyperlinks = null;
    try
    {
      hyperlinks = GetRequiredProperty(worksheet, "Hyperlinks");
      var count = Convert.ToInt32(GetRequiredProperty(hyperlinks, "Count"), CultureInfo.InvariantCulture);
      for (var index = 1; index <= count; index++)
      {
        object? hyperlink = null;
        object? parent = null;
        try
        {
          hyperlink = InvokeProperty(hyperlinks, "Item", index) ??
            throw new InvalidOperationException("The worksheet hyperlink could not be resolved.");
          parent = GetRequiredProperty(hyperlink, "Parent");
          var row = Convert.ToInt32(GetRequiredProperty(parent, "Row"), CultureInfo.InvariantCulture);
          if (row >= firstRow && row <= lastRow)
          {
            result.Add(row);
          }
        }
        finally
        {
          ComRelease.Release(parent);
          ComRelease.Release(hyperlink);
        }
      }

      return result;
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return AllRows(firstRow, lastRow);
    }
    finally
    {
      ComRelease.Release(hyperlinks);
    }
  }

  private static HashSet<int> ReadRowsWithShapes(object worksheet, int firstRow, int lastRow)
  {
    var result = new HashSet<int>();
    object? shapes = null;
    try
    {
      shapes = GetRequiredProperty(worksheet, "Shapes");
      var count = Convert.ToInt32(GetRequiredProperty(shapes, "Count"), CultureInfo.InvariantCulture);
      for (var index = 1; index <= count; index++)
      {
        object? shape = null;
        object? topLeftCell = null;
        object? bottomRightCell = null;
        try
        {
          shape = InvokeMethod(shapes, "Item", index) ??
            throw new InvalidOperationException("The worksheet Shape could not be resolved.");
          topLeftCell = GetRequiredProperty(shape, "TopLeftCell");
          bottomRightCell = GetRequiredProperty(shape, "BottomRightCell");
          var shapeStartRow = Convert.ToInt32(GetRequiredProperty(topLeftCell, "Row"), CultureInfo.InvariantCulture);
          var shapeEndRow = Convert.ToInt32(GetRequiredProperty(bottomRightCell, "Row"), CultureInfo.InvariantCulture);
          for (var row = Math.Max(firstRow, shapeStartRow); row <= Math.Min(lastRow, shapeEndRow); row++)
          {
            result.Add(row);
          }
        }
        finally
        {
          ComRelease.Release(bottomRightCell);
          ComRelease.Release(topLeftCell);
          ComRelease.Release(shape);
        }
      }

      return result;
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return AllRows(firstRow, lastRow);
    }
    finally
    {
      ComRelease.Release(shapes);
    }
  }

  private static bool HasAnyCellContent(object rowRange) =>
    HasAnyValue(rowRange, "Value2") || HasAnyValue(rowRange, "Formula");

  private static bool HasAnyValue(object target, string propertyName)
  {
    if (!TryGetProperty(target, propertyName, out var value) || value is null)
    {
      return false;
    }

    if (value is Array array)
    {
      foreach (var item in array)
      {
        if (IsNonEmptyValue(item))
        {
          return true;
        }
      }

      return false;
    }

    return IsNonEmptyValue(value);
  }

  private static bool IsNonEmptyValue(object? value) =>
    value is not null &&
    value is not DBNull &&
    (value is not string text || text.Length > 0);

  private static bool HasMergeOrUnknown(object rowRange) =>
    !TryGetProperty(rowRange, "MergeCells", out var value) ||
    value is null ||
    Convert.ToBoolean(value, CultureInfo.InvariantCulture);

  private static HashSet<int> AllRows(int firstRow, int lastRow) =>
    Enumerable.Range(firstRow, checked(lastRow - firstRow + 1)).ToHashSet();

  private RowMutationResult MutateWorkbook(
    object application,
    object workbook,
    WorkbookIdentity identity,
    string worksheetName,
    RowMutationPlan plan)
  {
    if (!WorkbookWindowMatchesIdentity(workbook, identity))
    {
      return RowMutationResult.Failed(
        plan.Operation,
        worksheetName,
        "The selected Workbook was closed or reopened; refresh and select it again.");
    }

    var eventsWereEnabled = Convert.ToBoolean(
      GetRequiredProperty(application, "EnableEvents"),
      CultureInfo.InvariantCulture);

    object? worksheet = null;
    object? rows = null;
    object? firstRow = null;
    object? targetRows = null;
    object? activeCell = null;
    RowDeletionSnapshot? deletionSnapshot = null;
    var eventsSuppressed = false;
    var screenUpdatingWasEnabled = true;
    var displayAlertsWereEnabled = true;
    bool? displayPageBreaksWereVisible = null;
    var applicationStateRead = false;
    var resolvedWorksheetName = worksheetName;
    var resolvedStartRow = plan.StartRow;
    var resolvedCount = plan.Count;
    try
    {
      if (IsWorkbookReadOnly(workbook) || !WorkbookWindowMatchesIdentity(workbook, identity))
      {
        return RowMutationResult.Failed(
          plan.Operation,
          worksheetName,
          "The selected Workbook became read-only or changed session before the row mutation; no rows were changed.");
      }

      screenUpdatingWasEnabled = Convert.ToBoolean(
        GetRequiredProperty(application, "ScreenUpdating"),
        CultureInfo.InvariantCulture);
      displayAlertsWereEnabled = Convert.ToBoolean(
        GetRequiredProperty(application, "DisplayAlerts"),
        CultureInfo.InvariantCulture);
      applicationStateRead = true;
      SetProperty(application, "EnableEvents", false);
      eventsSuppressed = true;
      SetProperty(application, "ScreenUpdating", false);
      SetProperty(application, "DisplayAlerts", false);
      InvokeMethod(workbook, "Activate");
      worksheet = ResolveWorksheet(application, workbook, worksheetName, out resolvedWorksheetName);
      if (TryGetProperty(worksheet, "DisplayPageBreaks", out var displayPageBreaks) &&
        displayPageBreaks is not null && displayPageBreaks is not DBNull)
      {
        displayPageBreaksWereVisible = Convert.ToBoolean(displayPageBreaks, CultureInfo.InvariantCulture);
      }
      if (IsWorksheetProtected(worksheet))
      {
        return RowMutationResult.Failed(
          plan.Operation,
          resolvedWorksheetName,
          $"シート {resolvedWorksheetName} は保護されています。保護を解除してから行を変更してください。");
      }

      if (plan.DeletionRequest is not null)
      {
        var snapshot = CaptureRowSafetySnapshot(
          worksheet,
          plan.DeletionRequest.CaseStartRow,
          plan.DeletionRequest.CaseEndRow,
          out var lastContentRow);
        var plannedRows = deletionPlanner.Plan(
          plan.DeletionRequest.CaseStartRow,
          plan.DeletionRequest.CaseEndRow,
          lastContentRow,
          plan.DeletionRequest.TailRows,
          snapshot);
        if (plannedRows.Count == 0)
        {
          return RowMutationResult.NoChange(
            plan.Operation,
            resolvedWorksheetName,
            "削除可能な安全な末尾行はありません。シートは変更していません。");
        }

        ValidateTrailingPlan(
          plannedRows,
          plan.DeletionRequest.CaseStartRow,
          plan.DeletionRequest.CaseEndRow);
        resolvedStartRow = plannedRows[0];
        resolvedCount = plannedRows.Count;
      }

      if (plan.ExactSafeDeletion)
      {
        if (plan.RestoreSnapshot is not null &&
          (plan.RestoreSnapshot.RestoredFingerprint is null ||
          !string.Equals(
            CaptureRowFingerprint(worksheet, resolvedStartRow, resolvedCount, includeWorksheetExtent: false),
            plan.RestoreSnapshot.RestoredFingerprint,
            StringComparison.Ordinal)))
        {
          return RowMutationResult.NoChange(
            plan.Operation,
            resolvedWorksheetName,
            "Undo後の行内容または書式が変更されたためRedoできません。シートは変更していません。");
        }

        var endRow = checked(resolvedStartRow + resolvedCount - 1);
        var snapshot = CaptureRowSafetySnapshot(worksheet, resolvedStartRow, endRow, out _);
        if (snapshot.Any(state => !state.CanDelete))
        {
          return RowMutationResult.NoChange(
            plan.Operation,
            resolvedWorksheetName,
            "対象行は挿入後に変更されたためUndo/Redoできません。シートは変更していません。");
        }
      }

      if (plan.UseActiveCell)
      {
        InvokeMethod(worksheet, "Activate");
        activeCell = GetRequiredProperty(application, "ActiveCell");
        resolvedStartRow = Convert.ToInt32(
          GetRequiredProperty(activeCell, "Row"),
          CultureInfo.InvariantCulture);
        ValidateInsertion(resolvedStartRow, resolvedCount);
      }

      if (resolvedCount < 1 ||
        !WorkbookWindowMatchesIdentity(workbook, identity) ||
        IsWorkbookReadOnly(workbook))
      {
        return RowMutationResult.Failed(
          plan.Operation,
          resolvedWorksheetName,
          "The target Workbook window changed before the row mutation; no rows were changed.");
      }

      if (Convert.ToBoolean(
        GetRequiredProperty(workbook, "ReadOnly"),
        CultureInfo.InvariantCulture))
      {
        return RowMutationResult.Failed(
          plan.Operation,
          resolvedWorksheetName,
          "The selected Workbook became read-only before the row mutation; no rows were changed.");
      }

      rows = GetRequiredProperty(worksheet, "Rows");
      if (plan.RestoreSnapshot is not null && plan.Operation is RowMutationOperation.Insert)
      {
        var currentFingerprint = CaptureRowFingerprint(
          worksheet,
          Math.Max(1, resolvedStartRow - 1),
          Math.Min(resolvedCount + 2, ExcelWorksheetLimits.MaximumRow - Math.Max(1, resolvedStartRow - 1) + 1),
          includeWorksheetExtent: true);
        if (plan.RestoreSnapshot.PostDeleteFingerprint is null ||
          !string.Equals(currentFingerprint, plan.RestoreSnapshot.PostDeleteFingerprint, StringComparison.Ordinal))
        {
          return RowMutationResult.NoChange(
            plan.Operation,
            resolvedWorksheetName,
            "削除後に対象位置の行構造または内容が変更されたためUndoできません。シートは変更していません。");
        }
        if (plan.RestoreSnapshot.PostDeletionFormulas is null ||
          !FormulaMapsEqual(CaptureFormulaMap(worksheet), plan.RestoreSnapshot.PostDeletionFormulas))
        {
          return RowMutationResult.NoChange(
            plan.Operation,
            resolvedWorksheetName,
            "削除後に数式が変更されたためUndoできません。シートは変更していません。");
        }
        if (plan.RestoreSnapshot.PostDeletionNames is null ||
          !StringMapsEqual(CaptureNameMap(workbook), plan.RestoreSnapshot.PostDeletionNames) ||
          !string.Equals(CapturePrintArea(worksheet), plan.RestoreSnapshot.PostDeletionPrintArea, StringComparison.Ordinal))
        {
          return RowMutationResult.NoChange(
            plan.Operation,
            resolvedWorksheetName,
            "削除後に名前定義または印刷範囲が変更されたためUndoできません。シートは変更していません。");
        }
      }

      firstRow = InvokeProperty(rows, "Item", resolvedStartRow) ??
        throw new InvalidOperationException("The requested Excel row could not be resolved.");
      targetRows = InvokeProperty(firstRow, "Resize", resolvedCount) ??
        throw new InvalidOperationException("The requested Excel row range could not be resolved.");
      if (plan.InsertedRowHeightPoints is double insertedRowHeight)
      {
        SetProperty(targetRows, "Hidden", false);
        SetProperty(targetRows, "RowHeight", insertedRowHeight);
        return RowMutationResult.SucceededResult(
          plan.Operation,
          resolvedWorksheetName,
          resolvedStartRow,
          resolvedCount,
          $"自動配置用の追加行 {resolvedWorksheetName}!R{resolvedStartRow}:R{resolvedStartRow + resolvedCount - 1} を表示し、行高を {insertedRowHeight:0.##}pt に揃えました（未保存）。",
          null);
      }
      if (plan.CaptureUndoSnapshot)
      {
        EnsureNoExternalWorksheetFormulas(workbook, resolvedWorksheetName);
        deletionSnapshot = CaptureNativeRowSnapshot(
          application,
          workbook,
          targetRows,
          identity,
          resolvedWorksheetName,
          resolvedStartRow,
          resolvedCount);
        deletionSnapshot.PreDeletionFormulas = CaptureFormulaMap(worksheet);
        deletionSnapshot.PreDeletionNames = CaptureNameMap(workbook);
        deletionSnapshot.PreDeletionPrintArea = CapturePrintArea(worksheet);
      }

      InvokeMethod(targetRows, plan.Operation is RowMutationOperation.Insert ? "Insert" : "Delete");

      if (deletionSnapshot is not null)
      {
        deletionSnapshot.PostDeleteFingerprint = CaptureRowFingerprint(
          worksheet,
          Math.Max(1, resolvedStartRow - 1),
          Math.Min(resolvedCount + 2, ExcelWorksheetLimits.MaximumRow - Math.Max(1, resolvedStartRow - 1) + 1),
          includeWorksheetExtent: true);
        deletionSnapshot.PostDeletionFormulas = CaptureFormulaMap(worksheet);
        deletionSnapshot.PostDeletionNames = CaptureNameMap(workbook);
        deletionSnapshot.PostDeletionPrintArea = CapturePrintArea(worksheet);
      }

      if (plan.RestoreSnapshot is not null)
      {
        // Excel may retarget the Range object passed to Insert to the rows that were shifted.
        // Resolve the newly inserted band again before pasting the native snapshot.
        ComRelease.Release(targetRows);
        targetRows = null;
        ComRelease.Release(firstRow);
        firstRow = InvokeProperty(rows, "Item", resolvedStartRow) ??
          throw new InvalidOperationException("The inserted Excel row could not be resolved.");
        targetRows = InvokeProperty(firstRow, "Resize", resolvedCount) ??
          throw new InvalidOperationException("The inserted Excel row range could not be resolved.");
        try
        {
          RestoreNativeRowSnapshot(application, workbook, targetRows, plan.RestoreSnapshot);
          RestoreExactFormulaMap(worksheet, plan.RestoreSnapshot.PreDeletionFormulas ?? []);
          RestoreNameMap(workbook, plan.RestoreSnapshot.PreDeletionNames ?? []);
          RestorePrintArea(worksheet, plan.RestoreSnapshot.PreDeletionPrintArea);
          plan.RestoreSnapshot.RestoredFingerprint = CaptureRowFingerprint(
            worksheet,
            resolvedStartRow,
            resolvedCount,
            includeWorksheetExtent: false);
        }
        catch
        {
          // The inserted band is owned by this operation. Removing it is the only safe
          // compensation if the native paste fails part-way through.
          try
          {
            _ = InvokeMethod(targetRows, "Delete");
            RestoreExactFormulaMap(worksheet, plan.RestoreSnapshot.PostDeletionFormulas ?? []);
            RestoreNameMap(workbook, plan.RestoreSnapshot.PostDeletionNames ?? []);
            RestorePrintArea(worksheet, plan.RestoreSnapshot.PostDeletionPrintArea);
          }
          catch (Exception exception) when (IsAutomationFailure(exception))
          {
            throw new InvalidOperationException("行Undoの補償に失敗しました。Workbookを保存せず状態を確認してください。", exception);
          }
          throw;
        }
      }

      return RowMutationResult.SucceededResult(
        plan.Operation,
        resolvedWorksheetName,
        resolvedStartRow,
        resolvedCount,
        plan.Operation is RowMutationOperation.Insert
          ? $"{resolvedCount} 行を {resolvedWorksheetName}!R{resolvedStartRow} から挿入しました（未保存）。"
          : $"安全な末尾 {resolvedCount} 行を {resolvedWorksheetName}!R{resolvedStartRow} から削除しました（未保存）。",
        deletionSnapshot);
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      deletionSnapshot?.Dispose();
      return RowMutationResult.Failed(
        plan.Operation,
        resolvedWorksheetName,
        $"Excelへの行変更に失敗しました (0x{GetAutomationHResult(exception):X8})。{exception.Message} 変更結果を確認してください。");
    }
    finally
    {
      if (worksheet is not null && displayPageBreaksWereVisible.HasValue)
      {
        try
        {
          SetProperty(worksheet, "DisplayPageBreaks", displayPageBreaksWereVisible.Value);
        }
        catch (Exception exception) when (IsAutomationFailure(exception))
        {
          // DisplayPageBreaks is unavailable when Excel has no usable printer.
        }
      }

      if (eventsSuppressed)
      {
        try
        {
          SetProperty(application, "EnableEvents", eventsWereEnabled);
        }
        catch (Exception exception) when (IsAutomationFailure(exception))
        {
          // The Workbook may have disconnected after the operation.
        }
      }

      if (applicationStateRead)
      {
        try
        {
          SetProperty(application, "ScreenUpdating", screenUpdatingWasEnabled);
          SetProperty(application, "DisplayAlerts", displayAlertsWereEnabled);
        }
        catch (Exception exception) when (IsAutomationFailure(exception))
        {
          // The Workbook may have disconnected after the operation.
        }
      }

      ComRelease.Release(targetRows);
      ComRelease.Release(firstRow);
      ComRelease.Release(rows);
      ComRelease.Release(activeCell);
      ComRelease.Release(worksheet);
    }
  }

  private static void EnsureNoExternalWorksheetFormulas(object workbook, string targetWorksheetName)
  {
    object? worksheets = null;
    try
    {
      worksheets = GetRequiredProperty(workbook, "Worksheets");
      var count = Convert.ToInt32(GetRequiredProperty(worksheets, "Count"), CultureInfo.InvariantCulture);
      for (var index = 1; index <= count; index++)
      {
        object? sheet = null;
        try
        {
          sheet = InvokeProperty(worksheets, "Item", index);
          var name = Convert.ToString(GetRequiredProperty(sheet!, "Name"), CultureInfo.InvariantCulture);
          if (!string.Equals(name, targetWorksheetName, StringComparison.OrdinalIgnoreCase) &&
            CaptureFormulaMap(sheet!).Count != 0)
          {
            throw new InvalidOperationException(
              $"別Sheet ({name}) に数式があるため、参照を完全に保証できるよう行削除を中止しました。");
          }
        }
        finally
        {
          ComRelease.Release(sheet);
        }
      }
    }
    finally
    {
      ComRelease.Release(worksheets);
    }
  }

  private static bool RangeHasNoMerges(object worksheet, int firstRow, int lastRow)
  {
    object? range = null;
    try
    {
      range = InvokeProperty(worksheet, "Range", $"{firstRow}:{lastRow}");
      return range is not null &&
        TryGetProperty(range, "MergeCells", out var value) &&
        value is not null && value is not DBNull &&
        !Convert.ToBoolean(value, CultureInfo.InvariantCulture);
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return false;
    }
    finally
    {
      ComRelease.Release(range);
    }
  }

  private static Dictionary<CellReference, string> CaptureFormulaMap(object worksheet)
  {
    object? usedRange = null;
    object? columns = null;
    object? rows = null;
    try
    {
      usedRange = GetRequiredProperty(worksheet, "UsedRange");
      columns = GetRequiredProperty(usedRange, "Columns");
      rows = GetRequiredProperty(usedRange, "Rows");
      var firstRow = Convert.ToInt32(GetRequiredProperty(usedRange, "Row"), CultureInfo.InvariantCulture);
      var firstColumn = Convert.ToInt32(GetRequiredProperty(usedRange, "Column"), CultureInfo.InvariantCulture);
      var rowCount = Convert.ToInt32(GetRequiredProperty(rows, "Count"), CultureInfo.InvariantCulture);
      var columnCount = Convert.ToInt32(GetRequiredProperty(columns, "Count"), CultureInfo.InvariantCulture);
      if ((long)rowCount * columnCount > MaximumDependencyCells)
      {
        throw new InvalidOperationException("数式依存の安全確認範囲が250,000セルを超えるため、行削除を中止しました。");
      }
      var formulas = GetRequiredProperty(usedRange, "Formula");
      var result = new Dictionary<CellReference, string>();
      for (var row = firstRow; row < checked(firstRow + rowCount); row++)
      {
        for (var column = firstColumn; column < checked(firstColumn + columnCount); column++)
        {
          var formula = Convert.ToString(
            MatrixValue(formulas, row - firstRow, column - firstColumn, rowCount, columnCount),
            CultureInfo.InvariantCulture);
          if (!string.IsNullOrEmpty(formula) && formula.StartsWith("=", StringComparison.Ordinal))
          {
            result[new CellReference(row, column)] = formula;
          }
        }
      }
      return result;
    }
    finally
    {
      ComRelease.Release(rows);
      ComRelease.Release(columns);
      ComRelease.Release(usedRange);
    }
  }

  private static object? MatrixValue(object? value, int rowOffset, int columnOffset, int rowCount, int columnCount)
  {
    if (value is not Array array)
    {
      return rowCount == 1 && columnCount == 1 ? value : null;
    }

    return array.GetValue(
      checked(rowOffset + array.GetLowerBound(0)),
      checked(columnOffset + array.GetLowerBound(1)));
  }

  private static bool FormulaMapsEqual(
    IReadOnlyDictionary<CellReference, string> left,
    IReadOnlyDictionary<CellReference, string> right) =>
    left.Count == right.Count && left.All(pair =>
      right.TryGetValue(pair.Key, out var value) && string.Equals(pair.Value, value, StringComparison.Ordinal));

  private static Dictionary<string, string> CaptureNameMap(object workbook)
  {
    object? names = null;
    try
    {
      names = GetRequiredProperty(workbook, "Names");
      var count = Convert.ToInt32(GetRequiredProperty(names, "Count"), CultureInfo.InvariantCulture);
      var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
      for (var index = 1; index <= count; index++)
      {
        object? name = null;
        try
        {
          name = InvokeProperty(names, "Item", index);
          var key = Convert.ToString(GetRequiredProperty(name!, "Name"), CultureInfo.InvariantCulture);
          var refersTo = Convert.ToString(GetRequiredProperty(name!, "RefersTo"), CultureInfo.InvariantCulture);
          if (!string.IsNullOrEmpty(key) && refersTo is not null)
          {
            result[key] = refersTo;
          }
        }
        finally
        {
          ComRelease.Release(name);
        }
      }
      return result;
    }
    finally
    {
      ComRelease.Release(names);
    }
  }

  private static bool StringMapsEqual(
    IReadOnlyDictionary<string, string> left,
    IReadOnlyDictionary<string, string> right) =>
    left.Count == right.Count && left.All(pair =>
      right.TryGetValue(pair.Key, out var value) && string.Equals(pair.Value, value, StringComparison.Ordinal));

  private static void RestoreNameMap(object workbook, IReadOnlyDictionary<string, string> before)
  {
    object? names = null;
    try
    {
      names = GetRequiredProperty(workbook, "Names");
      foreach (var pair in before)
      {
        object? name = null;
        try
        {
          name = InvokeProperty(names, "Item", pair.Key);
          if (name is null)
          {
            _ = InvokeMethod(names, "Add", pair.Key, pair.Value);
          }
          else
          {
            SetProperty(name, "RefersTo", pair.Value);
          }
        }
        finally
        {
          ComRelease.Release(name);
        }
      }
    }
    finally
    {
      ComRelease.Release(names);
    }
  }

  private static string CapturePrintArea(object worksheet)
  {
    object? pageSetup = null;
    try
    {
      pageSetup = GetRequiredProperty(worksheet, "PageSetup");
      return TryGetProperty(pageSetup, "PrintArea", out var value)
        ? Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty
        : string.Empty;
    }
    finally
    {
      ComRelease.Release(pageSetup);
    }
  }

  private static void RestorePrintArea(object worksheet, string? printArea)
  {
    object? pageSetup = null;
    try
    {
      pageSetup = GetRequiredProperty(worksheet, "PageSetup");
      SetProperty(pageSetup, "PrintArea", printArea ?? string.Empty);
    }
    finally
    {
      ComRelease.Release(pageSetup);
    }
  }

  private static void RestoreExactFormulaMap(
    object worksheet,
    IReadOnlyDictionary<CellReference, string> desired)
  {
    var current = CaptureFormulaMap(worksheet);
    foreach (var location in current.Keys.Except(desired.Keys))
    {
      object? cell = null;
      try
      {
        cell = InvokeProperty(worksheet, "Cells", location.Row, location.Column);
        _ = InvokeMethod(cell!, "ClearContents");
      }
      finally
      {
        ComRelease.Release(cell);
      }
    }
    foreach (var pair in desired)
    {
      object? cell = null;
      try
      {
        cell = InvokeProperty(worksheet, "Cells", pair.Key.Row, pair.Key.Column);
        SetProperty(cell!, "Formula", pair.Value);
      }
      finally
      {
        ComRelease.Release(cell);
      }
    }
  }

  private static string CaptureRowFingerprint(
    object worksheet,
    int startRow,
    int count,
    bool includeWorksheetExtent)
  {
    object? usedRange = null;
    object? usedColumns = null;
    object? usedRows = null;
    object? rows = null;
    try
    {
      usedRange = GetRequiredProperty(worksheet, "UsedRange");
      usedColumns = GetRequiredProperty(usedRange, "Columns");
      usedRows = GetRequiredProperty(usedRange, "Rows");
      var firstColumn = Convert.ToInt32(GetRequiredProperty(usedRange, "Column"), CultureInfo.InvariantCulture);
      var columnCount = Math.Max(1, Convert.ToInt32(GetRequiredProperty(usedColumns, "Count"), CultureInfo.InvariantCulture));
      if ((long)count * columnCount > MaximumDependencyCells)
      {
        throw new InvalidOperationException("行状態の安全確認範囲が250,000セルを超えるため、履歴操作を中止しました。");
      }
      var builder = new StringBuilder();
      if (includeWorksheetExtent)
      {
        builder.Append(GetRequiredProperty(usedRange, "Row")).Append('|')
          .Append(GetRequiredProperty(usedRows, "Count")).Append('|')
          .Append(firstColumn).Append('|').Append(columnCount).Append(';');
      }

      rows = GetRequiredProperty(worksheet, "Rows");
      for (var row = startRow; row < checked(startRow + count); row++)
      {
        object? wholeRow = null;
        try
        {
          wholeRow = InvokeProperty(rows, "Item", row);
          AppendFingerprintProperty(builder, wholeRow, "RowHeight");
          AppendFingerprintProperty(builder, wholeRow, "Hidden");
          AppendFingerprintProperty(builder, wholeRow, "OutlineLevel");
          AppendFingerprintProperty(builder, wholeRow, "PageBreak");
        }
        finally
        {
          ComRelease.Release(wholeRow);
        }

        for (var column = firstColumn; column < checked(firstColumn + columnCount); column++)
        {
          object? cell = null;
          object? font = null;
          object? interior = null;
          object? validation = null;
          object? conditions = null;
          try
          {
            cell = InvokeProperty(worksheet, "Cells", row, column);
            AppendFingerprintProperty(builder, cell, "Value2");
            AppendFingerprintProperty(builder, cell, "Formula");
            AppendFingerprintProperty(builder, cell, "NumberFormat");
            AppendFingerprintProperty(builder, cell, "Style");
            AppendFingerprintProperty(builder, cell, "WrapText");
            AppendFingerprintProperty(builder, cell, "HorizontalAlignment");
            AppendFingerprintProperty(builder, cell, "VerticalAlignment");
            if (TryGetProperty(cell!, "Font", out font) && font is not null)
            {
              AppendFingerprintProperty(builder, font, "Name");
              AppendFingerprintProperty(builder, font, "Size");
              AppendFingerprintProperty(builder, font, "Bold");
              AppendFingerprintProperty(builder, font, "Italic");
              AppendFingerprintProperty(builder, font, "Color");
            }
            if (TryGetProperty(cell!, "Interior", out interior) && interior is not null)
            {
              AppendFingerprintProperty(builder, interior, "Color");
              AppendFingerprintProperty(builder, interior, "Pattern");
            }
            if (TryGetProperty(cell!, "Validation", out validation) && validation is not null)
            {
              AppendFingerprintProperty(builder, validation, "Type");
              AppendFingerprintProperty(builder, validation, "Formula1");
              AppendFingerprintProperty(builder, validation, "Formula2");
            }
            if (TryGetProperty(cell!, "FormatConditions", out conditions) && conditions is not null)
            {
              AppendFingerprintProperty(builder, conditions, "Count");
            }
          }
          finally
          {
            ComRelease.Release(conditions);
            ComRelease.Release(validation);
            ComRelease.Release(interior);
            ComRelease.Release(font);
            ComRelease.Release(cell);
          }
        }
      }

      return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(builder.ToString())));
    }
    finally
    {
      ComRelease.Release(rows);
      ComRelease.Release(usedRows);
      ComRelease.Release(usedColumns);
      ComRelease.Release(usedRange);
    }
  }

  private static void AppendFingerprintProperty(StringBuilder builder, object? target, string propertyName)
  {
    if (target is null)
    {
      builder.Append("<null>;");
      return;
    }

    try
    {
      builder.Append(Convert.ToString(GetRequiredProperty(target, propertyName), CultureInfo.InvariantCulture)).Append(';');
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      builder.Append("<unsupported>;");
    }
  }

  private static RowDeletionSnapshot CaptureNativeRowSnapshot(
    object application,
    object workbook,
    object sourceRows,
    WorkbookIdentity identity,
    string worksheetName,
    int startRow,
    int count)
  {
    var directory = Path.Combine(Path.GetTempPath(), "EvidenceCrafter", "RowUndo");
    Directory.CreateDirectory(directory);
    var path = Path.Combine(directory, $"{Guid.NewGuid():N}.xlsx");
    object? workbooks = null;
    object? backupWorkbook = null;
    object? worksheets = null;
    object? backupSheet = null;
    object? backupRows = null;
    object? firstBackupRow = null;
    try
    {
      workbooks = GetRequiredProperty(application, "Workbooks");
      backupWorkbook = InvokeMethod(workbooks, "Add") ?? throw new InvalidOperationException("Excel backup Workbook could not be created.");
      worksheets = GetRequiredProperty(backupWorkbook, "Worksheets");
      backupSheet = InvokeProperty(worksheets, "Item", 1) ?? throw new InvalidOperationException("Excel backup Worksheet could not be created.");
      backupRows = GetRequiredProperty(backupSheet, "Rows");
      firstBackupRow = InvokeProperty(backupRows, "Item", 1) ?? throw new InvalidOperationException("Excel backup row could not be resolved.");
      var destinationRows = InvokeProperty(firstBackupRow, "Resize", count) ?? throw new InvalidOperationException("Excel backup row range could not be resolved.");
      try
      {
        _ = InvokeMethod(sourceRows, "Copy", destinationRows);
      }
      finally
      {
        ClearCutCopyMode(application);
        ComRelease.Release(destinationRows);
      }

      _ = InvokeMethod(backupWorkbook, "SaveAs", path, 51);
      _ = InvokeMethod(backupWorkbook, "Close", false);
      ComRelease.Release(backupWorkbook);
      backupWorkbook = null;
      InvokeMethod(workbook, "Activate");
      return new RowDeletionSnapshot(identity, worksheetName, startRow, count, path);
    }
    catch
    {
      try { if (backupWorkbook is not null) _ = InvokeMethod(backupWorkbook, "Close", false); } catch (Exception exception) when (IsAutomationFailure(exception)) { }
      try { File.Delete(path); } catch (IOException) { }
      throw;
    }
    finally
    {
      ComRelease.Release(firstBackupRow);
      ComRelease.Release(backupRows);
      ComRelease.Release(backupSheet);
      ComRelease.Release(worksheets);
      ComRelease.Release(backupWorkbook);
      ComRelease.Release(workbooks);
    }
  }

  private static void RestoreNativeRowSnapshot(
    object application,
    object workbook,
    object destinationRows,
    RowDeletionSnapshot snapshot)
  {
    object? workbooks = null;
    object? backupWorkbook = null;
    object? worksheets = null;
    object? backupSheet = null;
    object? rows = null;
    object? firstRow = null;
    object? sourceRows = null;
    try
    {
      workbooks = GetRequiredProperty(application, "Workbooks");
      backupWorkbook = InvokeMethod(workbooks, "Open", snapshot.BackupPath, 0, true) ?? throw new InvalidOperationException("Undo用Workbookを開けませんでした。");
      worksheets = GetRequiredProperty(backupWorkbook, "Worksheets");
      backupSheet = InvokeProperty(worksheets, "Item", 1) ?? throw new InvalidOperationException("Undo用Worksheetを開けませんでした。");
      rows = GetRequiredProperty(backupSheet, "Rows");
      firstRow = InvokeProperty(rows, "Item", 1) ?? throw new InvalidOperationException("Undo用行を開けませんでした。");
      sourceRows = InvokeProperty(firstRow, "Resize", snapshot.Count) ?? throw new InvalidOperationException("Undo用行範囲を開けませんでした。");
      _ = InvokeMethod(sourceRows, "Copy", destinationRows);
      _ = InvokeMethod(backupWorkbook, "Close", false);
      ComRelease.Release(backupWorkbook);
      backupWorkbook = null;
      _ = InvokeMethod(workbook, "Activate");
    }
    finally
    {
      ClearCutCopyMode(application);
      try { if (backupWorkbook is not null) _ = InvokeMethod(backupWorkbook, "Close", false); } catch (Exception exception) when (IsAutomationFailure(exception)) { }
      ComRelease.Release(sourceRows);
      ComRelease.Release(firstRow);
      ComRelease.Release(rows);
      ComRelease.Release(backupSheet);
      ComRelease.Release(worksheets);
      ComRelease.Release(backupWorkbook);
      ComRelease.Release(workbooks);
    }
  }

  private static void ClearCutCopyMode(object application)
  {
    try
    {
      SetProperty(application, "CutCopyMode", false);
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      // Copy-mode cleanup must not hide the result of the row operation.
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

  private static bool IsWorkbookReadOnly(object workbook) =>
    Convert.ToBoolean(GetRequiredProperty(workbook, "ReadOnly"), CultureInfo.InvariantCulture);

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

  private static void ValidateInsertion(RowInsertion insertion)
  {
    ValidateInsertion(insertion.AtRow, insertion.Count);
  }

  private static void ValidateInsertion(int atRow, int count)
  {
    ValidateInsertionCount(count);
    if (atRow < 1 ||
      atRow > ExcelWorksheetLimits.MaximumRow ||
      (long)atRow + count > ExcelWorksheetLimits.MaximumRow)
    {
      throw new ArgumentOutOfRangeException(
        nameof(atRow),
        "Inserted rows must remain below the worksheet's final row so existing bottom content cannot be discarded.");
    }
  }

  private static void ValidateInsertionCount(int count)
  {
    if (count < 1 || count >= ExcelWorksheetLimits.MaximumRow)
    {
      throw new ArgumentOutOfRangeException(
        nameof(count),
        "The inserted row count must be positive and leave the worksheet's final row intact.");
    }
  }

  private static void ValidateRowRange(int firstRow, int lastRow)
  {
    if (firstRow < 1 ||
      lastRow < firstRow ||
      lastRow > ExcelWorksheetLimits.MaximumRow)
    {
      throw new ArgumentOutOfRangeException(nameof(firstRow));
    }
  }

  private static void ValidateCaseRange(int caseStartRow, int caseEndRow) =>
    ValidateRowRange(caseStartRow, caseEndRow);

  private static void ValidateTrailingPlan(
    IReadOnlyList<int> plannedRows,
    int caseStartRow,
    int caseEndRow)
  {
    if (plannedRows[0] < caseStartRow || plannedRows[^1] != caseEndRow - 1)
    {
      throw new InvalidOperationException("The row deletion plan is not a trailing range of the selected Case.");
    }

    for (var index = 1; index < plannedRows.Count; index++)
    {
      if (plannedRows[index] != plannedRows[index - 1] + 1)
      {
        throw new InvalidOperationException("The row deletion plan is not contiguous.");
      }
    }
  }

  private static bool IsAutomationFailure(Exception exception) =>
    exception is COMException or TargetInvocationException or MissingMemberException or InvalidOperationException;

  private static int GetAutomationHResult(Exception exception) =>
    exception is TargetInvocationException { InnerException: not null } invocationException
      ? invocationException.InnerException!.HResult
      : exception.HResult;

  private sealed record RowMutationPlan(
    RowMutationOperation Operation,
    int StartRow,
    int Count,
    string Reason,
    bool UseActiveCell,
    TrailingDeletionRequest? DeletionRequest = null,
    bool ExactSafeDeletion = false,
    bool CaptureUndoSnapshot = false,
    RowDeletionSnapshot? RestoreSnapshot = null,
    double? InsertedRowHeightPoints = null);

  private sealed record TrailingDeletionRequest(int CaseStartRow, int CaseEndRow, int TailRows);

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

public enum RowMutationOperation
{
  Insert,
  Delete,
}

public sealed record RowMutationResult(
  bool Succeeded,
  bool Changed,
  RowMutationOperation Operation,
  string WorksheetName,
  int StartRow,
  int Count,
  string Message,
  RowDeletionSnapshot? DeletionSnapshot = null)
{
  public static RowMutationResult SucceededResult(
    RowMutationOperation operation,
    string worksheetName,
    int startRow,
    int count,
    string message,
    RowDeletionSnapshot? deletionSnapshot = null) =>
    new(true, true, operation, worksheetName, startRow, count, message, deletionSnapshot);

  public static RowMutationResult NoChange(
    RowMutationOperation operation,
    string worksheetName,
    string message) =>
    new(true, false, operation, worksheetName, 0, 0, message);

  public static RowMutationResult Failed(
    RowMutationOperation operation,
    string worksheetName,
    string message) =>
    new(false, false, operation, worksheetName, 0, 0, message);
}

/// <summary>Owns the native temporary workbook used by exact row-deletion Undo.</summary>
public sealed class RowDeletionSnapshot : IDisposable
{
  private int disposed;

  internal RowDeletionSnapshot(
    WorkbookIdentity workbook,
    string worksheetName,
    int startRow,
    int count,
    string backupPath)
  {
    ProcessId = workbook.ProcessId;
    WindowSessionToken = workbook.WindowSessionToken;
    WorkbookFullPath = workbook.FullPath;
    WorksheetName = worksheetName;
    StartRow = startRow;
    Count = count;
    BackupPath = backupPath;
  }

  public uint ProcessId { get; }
  public IntPtr WindowSessionToken { get; }
  public string WorkbookFullPath { get; }
  public string WorksheetName { get; }
  public int StartRow { get; }
  public int Count { get; }
  internal string BackupPath { get; }
  internal string? PostDeleteFingerprint { get; set; }
  internal string? RestoredFingerprint { get; set; }
  internal Dictionary<CellReference, string>? PreDeletionFormulas { get; set; }
  internal Dictionary<CellReference, string>? PostDeletionFormulas { get; set; }
  internal Dictionary<string, string>? PreDeletionNames { get; set; }
  internal Dictionary<string, string>? PostDeletionNames { get; set; }
  internal string? PreDeletionPrintArea { get; set; }
  internal string? PostDeletionPrintArea { get; set; }
  public bool IsDisposed => Volatile.Read(ref disposed) != 0;

  internal bool Matches(WorkbookIdentity workbook) =>
    workbook.ProcessId == ProcessId &&
    workbook.WindowSessionToken == WindowSessionToken &&
    string.Equals(workbook.FullPath, WorkbookFullPath, StringComparison.OrdinalIgnoreCase);

  public void Dispose()
  {
    if (Interlocked.Exchange(ref disposed, 1) != 0)
    {
      return;
    }

    try
    {
      File.Delete(BackupPath);
    }
    catch (IOException)
    {
      // A later process cleanup may remove a temporarily locked snapshot.
    }
    catch (UnauthorizedAccessException)
    {
      // Undo cleanup must never terminate the application.
    }
  }
}

public sealed record RowSafetySnapshotResult(
  bool Succeeded,
  string WorksheetName,
  IReadOnlyList<RowSafetyState> Rows,
  string Message)
{
  public static RowSafetySnapshotResult Failed(string worksheetName, string message) =>
    new(false, worksheetName, [], message);
}
