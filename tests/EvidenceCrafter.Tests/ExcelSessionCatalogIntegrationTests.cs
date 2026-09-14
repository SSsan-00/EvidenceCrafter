using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

[TestClass]
[TestCategory("ExcelIntegration")]
[DoNotParallelize]
public sealed partial class ExcelSessionCatalogIntegrationTests
{
  [TestMethod]
  public void DiscoverAndFocus_WithRealTemporaryWorkbook_VerifiesIdentityAndActiveCell() =>
    RunSupervisedScenario(Scenario.Operations);

  [TestMethod]
  public void SnapshotReads_WithRealTemporaryWorkbook_PreserveBoundariesAndWidths() =>
    RunSupervisedScenario(Scenario.SnapshotReads);

  [TestMethod]
  public void RowHeightReads_WithRealTemporaryWorkbook_MatchIndividualCells() =>
    RunSupervisedScenario(Scenario.RowHeights);

  [TestMethod]
  public void PlacementAnalysis_WithRealTemporaryWorkbook_ReportsTimings() =>
    RunSupervisedScenario(Scenario.PlacementAnalysis);

  [TestMethod]
  public void AppendImages_InReferenceCopies_DoNotOverlap() =>
    RunSupervisedScenario(Scenario.ReferenceAppend);

  [TestMethod]
  public void CaseNavigation_InReferenceCopies_CrossesSheetsBothWays() =>
    RunSupervisedScenario(Scenario.ReferenceNavigation);

  [TestMethod]
  public void PairedImages_AlignAfterRowGrowthAndRestoreReference() =>
    RunSupervisedScenario(Scenario.PairAlignment);

  private enum Scenario { Operations, SnapshotReads, RowHeights, PlacementAnalysis, ReferenceAppend, ReferenceNavigation, PairAlignment }

  private static void RunSupervisedScenario(Scenario scenario)
  {
    Exception? failure = null;
    string? inconclusiveReason = null;
    var supervisor = new ScenarioSupervisor(GetExcelProcessStartTimes());
    var completed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
    var thread = new Thread(() =>
    {
      try
      {
        RunRealWorkbookScenario(supervisor, scenario);
      }
      catch (OfficeUnavailableException exception)
      {
        inconclusiveReason = exception.Message;
      }
      catch (Exception exception)
      {
        failure = exception;
      }
      finally
      {
        completed.TrySetResult();
      }
    })
    {
      IsBackground = true,
    };
    try
    {
      thread.SetApartmentState(ApartmentState.STA);
      thread.Start();
      if (!completed.Task.Wait(TimeSpan.FromSeconds(scenario == Scenario.ReferenceAppend ? 180 : 55)))
      {
        supervisor.SuppressComCleanup();
        var terminated = supervisor.TryTerminate(out var terminationFailure);
        var joined = thread.Join(TimeSpan.FromSeconds(5));
        var terminationMessage = terminated
          ? "the generated Excel process was terminated"
          : $"generated Excel termination failed: {terminationFailure}";
        Assert.Fail(joined
          ? $"Excel integration scenario timed out; {terminationMessage}."
          : $"Excel integration scenario timed out; its STA did not terminate after {terminationMessage}.");
      }

      thread.Join();
      if (inconclusiveReason is not null)
      {
        Assert.Inconclusive(inconclusiveReason);
      }

      Assert.IsNull(failure, failure?.ToString());
    }
    finally
    {
      if (!thread.IsAlive)
      {
        supervisor.Dispose();
      }
    }
  }

  private static void RunRealWorkbookScenario(ScenarioSupervisor supervisor, Scenario scenario)
  {
    T RunExcelSta<T>(Func<T> action) => RunOnSta(
      action,
      () => supervisor.TryTerminate(out var terminationFailure) ? null : terminationFailure,
      supervisor.SuppressComCleanup);

    var excelType = Type.GetTypeFromProgID("Excel.Application", throwOnError: false) ??
      throw new OfficeUnavailableException("Desktop Excel is not installed for this integration test.");
    var temporaryDirectory = Path.Combine(
      Path.GetTempPath(),
      "EvidenceCrafter.Tests",
      Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
    Directory.CreateDirectory(temporaryDirectory);
    var workbookPath = Path.Combine(temporaryDirectory, "focus-integration.xlsx");
    var otherWorkbookPath = Path.Combine(temporaryDirectory, "focus-other.xlsx");
    var placementImagePath = Path.Combine(temporaryDirectory, "placement.png");

    object? application = null;
    object? workbooks = null;
    object? workbook = null;
    object? otherWorkbook = null;
    object? worksheet = null;
    object? otherWorksheet = null;
    object? activeCell = null;
    object? activeWorkbook = null;
    object? editedCell = null;
    object? otherAnchorCell = null;
    object? workbookForCancelledClose = null;
    IConnectionPoint? cancelConnectionPoint = null;
    ExcelApplicationSessionMonitor.AppEventsSink? cancelCloseSink = null;
    var cancelConnectionCookie = 0;
    ExcelApplicationSessionMonitor? sessionMonitor = null;
    var excelProcessId = 0u;
    Process? excelProcess = null;
    var ownsExcelProcess = false;
    Exception? scenarioFailure = null;
    var cleanupFailures = new List<Exception>();
    try
    {
      application = Activator.CreateInstance(excelType) ??
        throw new OfficeUnavailableException("Desktop Excel could not be started.");
      var applicationWindow = new IntPtr(Convert.ToInt64(
        GetRequiredProperty(application, "Hwnd"),
        CultureInfo.InvariantCulture));
      Assert.AreNotEqual(0u, NativeMethods.GetWindowThreadProcessId(applicationWindow, out excelProcessId));
      Assert.AreNotEqual(0u, excelProcessId);
      if (supervisor.IsPreexistingProcess(excelProcessId))
      {
        throw new OfficeUnavailableException(
          "Excel activation reused a pre-existing Excel process; the integration test requires an isolated process.");
      }

      ownsExcelProcess = true;
      excelProcess = Process.GetProcessById(checked((int)excelProcessId));
      _ = excelProcess.Handle;
      supervisor.Capture(excelProcess);
      SetProperty(application, "Visible", true);
      SetProperty(application, "DisplayAlerts", false);
      workbooks = GetRequiredProperty(application, "Workbooks");
      workbook = InvokeMethod(workbooks, "Add") ??
        throw new InvalidOperationException("Excel did not create a Workbook.");
      worksheet = GetRequiredProperty(application, "ActiveSheet");
      SetProperty(worksheet, "Name", "FocusTarget");
      _ = InvokeMethod(workbook, "SaveAs", workbookPath);

      otherWorkbook = InvokeMethod(workbooks, "Add") ??
        throw new InvalidOperationException("Excel did not create the second Workbook.");
      otherWorksheet = GetRequiredProperty(application, "ActiveSheet");
      SetProperty(otherWorksheet, "Name", "OtherTarget");
      _ = InvokeMethod(otherWorkbook, "SaveAs", otherWorkbookPath);

      if (scenario is Scenario.ReferenceAppend or Scenario.ReferenceNavigation)
      {
        VerifyReferenceAppend(workbooks, temporaryDirectory, placementImagePath, scenario == Scenario.ReferenceNavigation);
      }
      else if (scenario == Scenario.PairAlignment)
      {
        var identity = new ExcelSessionCatalog().Discover().Workbooks.Single(item =>
          string.Equals(item.FullPath, otherWorkbookPath, StringComparison.OrdinalIgnoreCase));
        VerifyPairedImageAlignment(otherWorksheet, identity, placementImagePath);
      }
      else if (scenario == Scenario.SnapshotReads)
      {
        VerifySnapshotReadPerformance(otherWorksheet);
      }
      else if (scenario == Scenario.RowHeights)
      {
        VerifyRowHeightReadPerformance(otherWorksheet);
      }
      else if (scenario == Scenario.PlacementAnalysis)
      {
        var discovery = RunExcelSta(() => new ExcelSessionCatalog().Discover());
        var identity = discovery.Workbooks.Single(item =>
          string.Equals(item.FullPath, otherWorkbookPath, StringComparison.OrdinalIgnoreCase));
        VerifyPlacementAnalysisPerformance(otherWorksheet, identity);
      }
      else
      {
      var discovery = RunExcelSta(() => new ExcelSessionCatalog().Discover());
      var identity = discovery.Workbooks.SingleOrDefault(item =>
        string.Equals(item.FullPath, workbookPath, StringComparison.OrdinalIgnoreCase));
      var otherIdentity = discovery.Workbooks.SingleOrDefault(item =>
        string.Equals(item.FullPath, otherWorkbookPath, StringComparison.OrdinalIgnoreCase));
      Assert.IsNotNull(identity, $"Temporary Workbook was not discovered. Warnings: {string.Join(" | ", discovery.Warnings)}");
      Assert.IsNotNull(otherIdentity, $"Second temporary Workbook was not discovered. Warnings: {string.Join(" | ", discovery.Warnings)}");
      Assert.AreNotEqual(IntPtr.Zero, identity.ExcelWindowHandle);
      Assert.AreNotEqual(0u, identity.ProcessId);
      Assert.AreNotEqual(IntPtr.Zero, identity.WindowSessionToken);
      Assert.IsFalse(identity.IsReadOnly);
      Assert.AreNotEqual(
        identity.ExcelWindowHandle,
        otherIdentity.ExcelWindowHandle,
        "Each Workbook must retain its own top-level Excel window handle.");
      Assert.IsTrue(Environment.Is64BitProcess, "The win-x64 compatibility test must use a 64-bit test host.");
      sessionMonitor = new ExcelApplicationSessionMonitor();
      var monitorWarnings = sessionMonitor.Refresh(
        discovery.Workbooks.Where(item => item.ProcessId == identity.ProcessId).ToArray());
      Assert.HasCount(0, monitorWarnings, string.Join(" | ", monitorWarnings));

      otherAnchorCell = GetRequiredProperty(otherWorksheet, "Cells", 1, 1);
      SetProperty(otherAnchorCell, "Value2", "keep row");
      Release(otherAnchorCell);
      otherAnchorCell = null;
      var rowService = new ExcelRowMutationService();
      var insertedRows = RunExcelSta(() => rowService.InsertRows(
        otherIdentity,
        "OtherTarget",
        new RowInsertion(3, 2, "integration test insertion")));
      Assert.IsTrue(insertedRows.Succeeded, insertedRows.Message);
      Assert.IsTrue(insertedRows.Changed, insertedRows.Message);
      Assert.AreEqual(2, insertedRows.Count);
      otherAnchorCell = GetRequiredProperty(otherWorksheet, "Rows", insertedRows.StartRow);
      var insertedBand = GetRequiredProperty(otherAnchorCell, "Resize", insertedRows.Count);
      SetProperty(insertedBand, "Hidden", true);
      SetProperty(insertedBand, "RowHeight", 5.0);
      Release(insertedBand);
      Release(otherAnchorCell);
      otherAnchorCell = null;
      var normalizedRows = RunExcelSta(() => rowService.NormalizeInsertedRows(
        otherIdentity,
        "OtherTarget",
        insertedRows.StartRow,
        insertedRows.Count,
        15.0));
      Assert.IsTrue(normalizedRows.Succeeded && normalizedRows.Changed, normalizedRows.Message);
      otherAnchorCell = GetRequiredProperty(otherWorksheet, "Rows", insertedRows.StartRow);
      insertedBand = GetRequiredProperty(otherAnchorCell, "Resize", insertedRows.Count);
      Assert.IsFalse(Convert.ToBoolean(GetRequiredProperty(insertedBand, "Hidden"), CultureInfo.InvariantCulture));
      Assert.AreEqual(15.0, Convert.ToDouble(GetRequiredProperty(insertedBand, "RowHeight"), CultureInfo.InvariantCulture), 0.01);
      Release(insertedBand);
      Release(otherAnchorCell);
      otherAnchorCell = null;
      otherAnchorCell = GetRequiredProperty(otherWorksheet, "Cells", insertedRows.StartRow, 1);
      SetProperty(otherAnchorCell, "Value2", "post-insert edit");
      Release(otherAnchorCell);
      otherAnchorCell = null;
      var refusedInsertUndo = RunExcelSta(() => rowService.DeleteRowsIfSafe(
        otherIdentity,
        "OtherTarget",
        insertedRows.StartRow,
        insertedRows.Count));
      Assert.IsTrue(refusedInsertUndo.Succeeded, refusedInsertUndo.Message);
      Assert.IsFalse(refusedInsertUndo.Changed, "Undo must not delete inserted rows that the user edited.");
      otherAnchorCell = GetRequiredProperty(otherWorksheet, "Cells", insertedRows.StartRow, 1);
      _ = InvokeMethod(otherAnchorCell, "ClearContents");
      Release(otherAnchorCell);
      otherAnchorCell = null;
      var insertUndo = RunExcelSta(() => rowService.DeleteRowsIfSafe(
        otherIdentity,
        "OtherTarget",
        insertedRows.StartRow,
        insertedRows.Count));
      Assert.IsTrue(insertUndo.Succeeded && insertUndo.Changed, insertUndo.Message);
      var insertRedo = RunExcelSta(() => rowService.InsertRows(
        otherIdentity,
        "OtherTarget",
        new RowInsertion(insertedRows.StartRow, insertedRows.Count, "integration redo")));
      Assert.IsTrue(insertRedo.Succeeded && insertRedo.Changed, insertRedo.Message);
      var safetySnapshot = RunExcelSta(() => rowService.CaptureRowSafetyStates(
        otherIdentity,
        "OtherTarget",
        firstRow: 1,
        lastRow: 12));
      Assert.IsTrue(safetySnapshot.Succeeded, safetySnapshot.Message);
      Assert.IsTrue(
        safetySnapshot.Rows.Where(state => state.Row >= 6).All(state => state.CanDelete),
        string.Join(" | ", safetySnapshot.Rows.Select(state =>
          $"{state.Row}:v={state.HasValueOrFormula},c={state.HasCommentOrNote},h={state.HasHyperlink},s={state.HasShape},m={state.IntersectsMerge}")));
      object? formattedRow = null;
      object? formattedCell = null;
      object? formattedFont = null;
      object? formattedInterior = null;
      try
      {
        formattedRow = GetRequiredProperty(otherWorksheet, "Rows", 6);
        SetProperty(formattedRow, "RowHeight", 31.5);
        formattedCell = GetRequiredProperty(otherWorksheet, "Cells", 6, 3);
        SetProperty(formattedCell, "NumberFormat", "0.00");
        SetProperty(formattedCell, "WrapText", true);
        formattedFont = GetRequiredProperty(formattedCell, "Font");
        SetProperty(formattedFont, "Bold", true);
        formattedInterior = GetRequiredProperty(formattedCell, "Interior");
        SetProperty(formattedInterior, "Color", 65535);
      }
      finally
      {
        Release(formattedInterior);
        Release(formattedFont);
        Release(formattedCell);
        Release(formattedRow);
      }
      formattedInterior = null;
      formattedFont = null;
      formattedCell = null;
      formattedRow = null;
      otherAnchorCell = GetRequiredProperty(otherWorksheet, "Cells", 1, 2);
      SetProperty(otherAnchorCell, "Formula", "=C6");
      Release(otherAnchorCell);
      otherAnchorCell = null;
      SetProperty(otherWorksheet, "DisplayPageBreaks", false);
      var deletedRows = RunExcelSta(() => rowService.DeleteTrailingRowsWithSnapshot(
        otherIdentity,
        "OtherTarget",
        caseStartRow: 1,
        caseEndRow: 12,
        tailRows: 4));
      Assert.IsTrue(deletedRows.Succeeded, deletedRows.Message);
      Assert.IsTrue(deletedRows.Changed, deletedRows.Message);
      Assert.AreEqual(6, deletedRows.Count);
      Assert.IsNotNull(deletedRows.DeletionSnapshot);
      Assert.IsFalse(Convert.ToBoolean(
        GetRequiredProperty(otherWorksheet, "DisplayPageBreaks"),
        CultureInfo.InvariantCulture));
      Assert.IsFalse(Convert.ToBoolean(
        GetRequiredProperty(application, "CutCopyMode"),
        CultureInfo.InvariantCulture));
      using var deletionSnapshot = deletedRows.DeletionSnapshot;
      object? shiftedRow = null;
      try
      {
        shiftedRow = GetRequiredProperty(otherWorksheet, "Rows", 1);
        _ = InvokeMethod(shiftedRow, "Insert");
      }
      finally
      {
        Release(shiftedRow);
      }
      var refusedShiftedUndo = RunExcelSta(() => rowService.RestoreDeletedRows(otherIdentity, deletionSnapshot));
      Assert.IsTrue(refusedShiftedUndo.Succeeded && !refusedShiftedUndo.Changed, refusedShiftedUndo.Message);
      try
      {
        shiftedRow = GetRequiredProperty(otherWorksheet, "Rows", 1);
        _ = InvokeMethod(shiftedRow, "Delete");
      }
      finally
      {
        Release(shiftedRow);
      }
      var deletionUndo = RunExcelSta(() => rowService.RestoreDeletedRows(otherIdentity, deletionSnapshot));
      Assert.IsTrue(deletionUndo.Succeeded && deletionUndo.Changed, deletionUndo.Message);
      Assert.IsFalse(Convert.ToBoolean(
        GetRequiredProperty(otherWorksheet, "DisplayPageBreaks"),
        CultureInfo.InvariantCulture));
      Assert.IsFalse(Convert.ToBoolean(
        GetRequiredProperty(application, "CutCopyMode"),
        CultureInfo.InvariantCulture));
      otherAnchorCell = GetRequiredProperty(otherWorksheet, "Cells", 1, 2);
      Assert.AreEqual("=C6", Convert.ToString(GetRequiredProperty(otherAnchorCell, "Formula"), CultureInfo.InvariantCulture));
      Release(otherAnchorCell);
      otherAnchorCell = null;
      try
      {
        formattedRow = GetRequiredProperty(otherWorksheet, "Rows", 6);
        Assert.AreEqual(31.5, Convert.ToDouble(GetRequiredProperty(formattedRow, "RowHeight"), CultureInfo.InvariantCulture), 0.05);
        formattedCell = GetRequiredProperty(otherWorksheet, "Cells", 6, 3);
        Assert.AreEqual("0.00", Convert.ToString(GetRequiredProperty(formattedCell, "NumberFormat"), CultureInfo.InvariantCulture));
        Assert.IsTrue(Convert.ToBoolean(GetRequiredProperty(formattedCell, "WrapText"), CultureInfo.InvariantCulture));
        formattedFont = GetRequiredProperty(formattedCell, "Font");
        Assert.IsTrue(Convert.ToBoolean(GetRequiredProperty(formattedFont, "Bold"), CultureInfo.InvariantCulture));
        formattedInterior = GetRequiredProperty(formattedCell, "Interior");
        Assert.AreEqual(65535, Convert.ToInt32(GetRequiredProperty(formattedInterior, "Color"), CultureInfo.InvariantCulture));
      }
      finally
      {
        Release(formattedInterior);
        Release(formattedFont);
        Release(formattedCell);
        Release(formattedRow);
      }
      try
      {
        formattedRow = GetRequiredProperty(otherWorksheet, "Rows", 6);
        SetProperty(formattedRow, "RowHeight", 32.5);
      }
      finally
      {
        Release(formattedRow);
      }
      var refusedFormattedRedo = RunExcelSta(() => rowService.DeleteRestoredRows(otherIdentity, deletionSnapshot));
      Assert.IsTrue(refusedFormattedRedo.Succeeded && !refusedFormattedRedo.Changed, refusedFormattedRedo.Message);
      try
      {
        formattedRow = GetRequiredProperty(otherWorksheet, "Rows", 6);
        SetProperty(formattedRow, "RowHeight", 31.5);
      }
      finally
      {
        Release(formattedRow);
      }
      var deletionRedo = RunExcelSta(() => rowService.DeleteRestoredRows(otherIdentity, deletionSnapshot));
      Assert.IsTrue(deletionRedo.Succeeded && deletionRedo.Changed, deletionRedo.Message);
      Assert.IsFalse(Convert.ToBoolean(
        GetRequiredProperty(otherWorksheet, "DisplayPageBreaks"),
        CultureInfo.InvariantCulture));
      otherAnchorCell = GetRequiredProperty(otherWorksheet, "Cells", 1, 1);
      Assert.AreEqual("keep row", Convert.ToString(
        GetRequiredProperty(otherAnchorCell, "Value2"),
        CultureInfo.CurrentCulture));
      Release(otherAnchorCell);
      otherAnchorCell = null;
      _ = InvokeMethod(otherWorksheet, "Protect");
      var protectedRows = RunExcelSta(() => rowService.InsertRows(
        otherIdentity,
        "OtherTarget",
        new RowInsertion(3, 1, "protected integration test")));
      Assert.IsFalse(protectedRows.Succeeded, "A protected Worksheet must reject row insertion.");
      Assert.IsTrue(protectedRows.Message.Contains("保護", StringComparison.Ordinal), protectedRows.Message);
      _ = InvokeMethod(otherWorksheet, "Unprotect");

      editedCell = GetRequiredProperty(worksheet, "Cells", 1, 1);
      SetProperty(editedCell, "Value2", "ordinary unsaved edit");
      Release(editedCell);
      editedCell = null;
      var editedDiscovery = RunExcelSta(() => new ExcelSessionCatalog().Discover());
      var editedIdentity = editedDiscovery.Workbooks.Single(item =>
        string.Equals(item.FullPath, workbookPath, StringComparison.OrdinalIgnoreCase));
      Assert.AreEqual(
        identity.ConnectionId,
        editedIdentity.ConnectionId,
        "An ordinary edit must not replace the identity of an open Workbook session.");

      _ = InvokeMethod(otherWorkbook, "Activate");
      SetProperty(application, "EnableEvents", true);
      var focus = RunExcelSta(() => new ExcelPlacementFocusService().FocusPlacedImage(
        identity,
        "FocusTarget",
        new CellReference(5, 4)));
      Assert.IsTrue(focus.Succeeded, focus.Message);
      Assert.IsTrue(Convert.ToBoolean(
        GetRequiredProperty(application, "EnableEvents"),
        CultureInfo.InvariantCulture));

      activeWorkbook = GetRequiredProperty(application, "ActiveWorkbook");
      Assert.AreEqual(
        workbookPath,
        Convert.ToString(GetRequiredProperty(activeWorkbook, "FullName"), CultureInfo.CurrentCulture),
        ignoreCase: true);
      activeCell = GetRequiredProperty(application, "ActiveCell");
      Assert.AreEqual(5, Convert.ToInt32(GetRequiredProperty(activeCell, "Row"), CultureInfo.InvariantCulture));
      Assert.AreEqual(4, Convert.ToInt32(GetRequiredProperty(activeCell, "Column"), CultureInfo.InvariantCulture));
      object? focusedWindow = null;
      try
      {
        focusedWindow = GetRequiredProperty(application, "ActiveWindow");
        Assert.AreEqual(2, Convert.ToInt32(GetRequiredProperty(focusedWindow, "ScrollRow"), CultureInfo.InvariantCulture));
        Assert.AreEqual(1, Convert.ToInt32(GetRequiredProperty(focusedWindow, "ScrollColumn"), CultureInfo.InvariantCulture));
      }
      finally
      {
        Release(focusedWindow);
      }

      File.WriteAllBytes(
        placementImagePath,
        Convert.FromBase64String(
           "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));

      VerifyOptionalLayoutScenarios(worksheet, identity, placementImagePath);

      SetCellValue(worksheet, 2, 3, "新");
      SetCellValue(worksheet, 2, 6, "旧");
      SetCellValue(worksheet, 3, 1, 1);
      SetCellValue(worksheet, 3, 2, 1);
      SetRangeBorder(worksheet, 3, 1, 3, 8, 8);
      SetRangeBorder(worksheet, 1, 5, 8, 5, 10);
      SetRangeBorder(worksheet, 8, 1, 8, 8, 9);
      // Arbitrary sheet names must navigate in tab order, in both directions.
      object? navigationSheet = null;
      try
      {
        _ = InvokeMethod(worksheet, "Copy", Type.Missing, worksheet);
        navigationSheet = GetRequiredProperty(workbook, "ActiveSheet");
        SetProperty(navigationSheet, "Name", "次の帳票");
        _ = InvokeMethod(worksheet, "Activate");
        var nextSheet = RunExcelSta(() => new ExcelCaseNavigationService().Navigate(
          identity, "FocusTarget", CaseNavigationDirection.Next, "1-1", EvidenceSide.Old));
        Assert.IsTrue(nextSheet.Succeeded, nextSheet.Message);
        Assert.AreEqual("次の帳票", nextSheet.WorksheetName);
        Assert.AreEqual(EvidenceSide.New, nextSheet.Side);
        var previousSheet = RunExcelSta(() => new ExcelCaseNavigationService().Navigate(
          identity, "次の帳票", CaseNavigationDirection.Previous, "1-1", EvidenceSide.New));
        Assert.IsTrue(previousSheet.Succeeded, previousSheet.Message);
        Assert.AreEqual("FocusTarget", previousSheet.WorksheetName);
        Assert.AreEqual(EvidenceSide.Old, previousSheet.Side);
        Assert.AreEqual(4, previousSheet.Target.Row);
        SetCellValue(navigationSheet, 3, 1, "");
        SetCellValue(navigationSheet, 3, 2, "");
        var failedNavigation = RunExcelSta(() => new ExcelCaseNavigationService().Navigate(
          identity, "FocusTarget", CaseNavigationDirection.Next, "1-1", EvidenceSide.Old));
        Assert.IsFalse(failedNavigation.Succeeded);
        StringAssert.Contains(failedNavigation.Message, "次の帳票");
        StringAssert.Contains(failedNavigation.Message, "解析できません");

        _ = InvokeMethod(navigationSheet, "Activate");
        var selectedSheetSnapshot = RunExcelSta(() =>
          new ExcelSheetSnapshotService().Capture(identity, "FocusTarget"));
        Assert.IsTrue(selectedSheetSnapshot.Succeeded, selectedSheetSnapshot.Message);
        object? selectedSheet = null;
        try
        {
          selectedSheet = GetRequiredProperty(workbook, "ActiveSheet");
          Assert.AreEqual("FocusTarget", Convert.ToString(
            GetRequiredProperty(selectedSheet, "Name"), CultureInfo.CurrentCulture));
        }
        finally
        {
          ReleaseOnce(selectedSheet);
        }
      }
      finally
      {
        if (navigationSheet is not null) _ = InvokeMethod(navigationSheet, "Delete");
        Release(navigationSheet);
      }
      object? automaticTargetCell = null;
      object? originalOtherCell = null;
      object? restoredWorkbook = null;
      object? automaticShapes = null;
      object? automaticShape = null;
      try
      {
        _ = InvokeMethod(otherWorkbook, "Activate");
        _ = InvokeMethod(otherWorksheet, "Activate");
        originalOtherCell = GetRequiredProperty(otherWorksheet, "Cells", 2, 2);
        _ = InvokeMethod(originalOtherCell, "Select");
        SetProperty(application, "EnableEvents", true);

        var directValidationIdentity = identity with { HasWorkbookRegistration = false };
        WorkbookSessionTokenRegistry.MarkProcessUnmonitored(identity.ProcessId);
        var unmonitoredSnapshot = RunExcelSta(() => new ExcelSheetSnapshotService().Capture(
          directValidationIdentity,
          "FocusTarget"));
        Assert.IsTrue(unmonitoredSnapshot.Succeeded, unmonitoredSnapshot.Message);
        var unmonitoredFocus = RunExcelSta(() => new ExcelPlacementFocusService().FocusPlacedImage(
          directValidationIdentity,
          "FocusTarget",
          new CellReference(5, 3)));
        Assert.IsTrue(unmonitoredFocus.Succeeded, unmonitoredFocus.Message);
        Assert.HasCount(0, sessionMonitor.Refresh([identity]));
        _ = InvokeMethod(otherWorkbook, "Activate");
        _ = InvokeMethod(otherWorksheet, "Activate");
        _ = InvokeMethod(originalOtherCell, "Select");

        var snapshotResult = RunExcelSta(() => new ExcelSheetSnapshotService().Capture(identity, "FocusTarget"));
        Assert.IsTrue(snapshotResult.Succeeded, snapshotResult.Message);
        Assert.IsNotNull(snapshotResult.Snapshot);
        var navigationSnapshot = RunExcelSta(() =>
          new ExcelSheetSnapshotService().CaptureForNavigation(identity, "FocusTarget", includeWorksheetNames: true));
        Assert.IsTrue(navigationSnapshot.Succeeded, navigationSnapshot.Message);
        Assert.IsNotNull(navigationSnapshot.Snapshot);
        Assert.IsEmpty(navigationSnapshot.Snapshot.Cells);
        Assert.IsEmpty(navigationSnapshot.Snapshot.RowHeights);
        Assert.IsEmpty(navigationSnapshot.Snapshot.ColumnWidths);
        CollectionAssert.AreEqual(
          snapshotResult.Snapshot.LayoutSignals.Anchors.ToArray(),
          navigationSnapshot.Snapshot.LayoutSignals.Anchors.ToArray());
        Assert.IsTrue(Convert.ToBoolean(GetRequiredProperty(application, "EnableEvents"), CultureInfo.InvariantCulture));
        restoredWorkbook = GetRequiredProperty(application, "ActiveWorkbook");
        Assert.AreEqual(
          workbookPath,
          Convert.ToString(GetRequiredProperty(restoredWorkbook, "FullName"), CultureInfo.CurrentCulture),
          ignoreCase: true,
          "Snapshot capture must activate the Workbook selected in EvidenceCrafter.");
        ReleaseOnce(restoredWorkbook);
        restoredWorkbook = null;

        _ = InvokeMethod(workbook, "Activate");
        _ = InvokeMethod(worksheet, "Activate");
        automaticTargetCell = GetRequiredProperty(worksheet, "Cells", 5, 3);
        _ = InvokeMethod(automaticTargetCell, "Select");
        var automaticService = new ExcelAutomaticPlacementService();
        var request = new[]
        {
          new AutomaticPlacementImage(placementImagePath, new ImageDimensions(120, 240)),
        };
        var automaticAnalysis = RunExcelSta(() => automaticService.Analyze(
          identity,
          "FocusTarget",
          EvidenceSide.New,
          request));
        Assert.IsTrue(automaticAnalysis.Succeeded, automaticAnalysis.Message);
        Assert.HasCount(1, automaticAnalysis.Steps);
        Assert.IsNotEmpty(automaticAnalysis.Steps[0].Plan.Insertions);

        var automaticPlacement = RunExcelSta(() => automaticService.PlaceImages(
          identity,
          "FocusTarget",
          EvidenceSide.New,
          request));
        Assert.IsTrue(automaticPlacement.Succeeded, automaticPlacement.Message);
        Assert.HasCount(1, automaticPlacement.PlacedImages);
        Assert.IsNotEmpty(automaticPlacement.AppliedInsertions);
        var automaticImage = automaticPlacement.PlacedImages[0];
        Assert.AreEqual("FocusTarget", automaticImage.WorksheetName);
        Assert.AreEqual(4, automaticImage.FocusCell.Column);
        var oppositeSide = RunExcelSta(() => new ExcelCaseNavigationService().Navigate(
          identity,
          "FocusTarget",
          CaseNavigationDirection.Next,
          automaticPlacement.Analysis!.CaseLabel,
          EvidenceSide.New,
          sameCaseThenNext: true));
        Assert.IsTrue(oppositeSide.Succeeded, oppositeSide.Message);
        Assert.AreEqual("1-1", oppositeSide.CaseLabel);
        Assert.AreEqual(EvidenceSide.Old, oppositeSide.Side);
        Assert.AreEqual(4, oppositeSide.Target.Row);

        automaticShapes = GetRequiredProperty(worksheet, "Shapes");
        automaticShape = InvokeMethod(automaticShapes, "Item", automaticImage.ShapeName) ??
          throw new InvalidOperationException("The automatic Shape could not be retrieved.");
        _ = InvokeMethod(automaticShape, "Select");
        var managedService = new ExcelManagedShapeService();
        var inspected = RunExcelSta(() => managedService.InspectSelection(identity));
        Assert.IsTrue(inspected.Succeeded, inspected.Message);
        Assert.IsNotNull(inspected.Shape);
        Assert.AreEqual(automaticImage.ShapeName, inspected.Shape.ShapeName);

        var managedBackupPath = Path.Combine(temporaryDirectory, "managed-backup.png");
        var exported = RunExcelSta(() => managedService.Export(
          identity,
          inspected.Shape,
          managedBackupPath));
        Assert.IsTrue(exported.Succeeded, exported.Message);
        Assert.IsGreaterThan(0, new FileInfo(managedBackupPath).Length);

        var replaced = RunExcelSta(() => managedService.Replace(
          identity,
          inspected.Shape,
          placementImagePath,
          new ImageDimensions(100, 50)));
        Assert.IsTrue(replaced.Succeeded && replaced.Changed, replaced.Message);
        Assert.IsNotNull(replaced.After);
        var exactUndo = RunExcelSta(() => managedService.Replace(
          identity,
          replaced.After,
          managedBackupPath,
          new ImageDimensions(inspected.Shape.WidthPoints, inspected.Shape.HeightPoints),
          inspected.Shape));
        Assert.IsTrue(exactUndo.Succeeded && exactUndo.Changed, exactUndo.Message);
        Assert.IsNotNull(exactUndo.After);
        Assert.AreEqual(inspected.Shape.WidthPoints, exactUndo.After.WidthPoints, 0.05);
        Assert.AreEqual(inspected.Shape.HeightPoints, exactUndo.After.HeightPoints, 0.05);
        var deleted = RunExcelSta(() => managedService.Delete(identity, exactUndo.After));
        Assert.IsTrue(deleted.Succeeded && deleted.Changed, deleted.Message);
        var restored = RunExcelSta(() => managedService.Restore(identity, inspected.Shape, managedBackupPath));
        Assert.IsTrue(restored.Succeeded && restored.Changed, restored.Message);
        Assert.IsNotNull(restored.After);
        var deletedRestored = RunExcelSta(() => managedService.Delete(identity, restored.After));
        Assert.IsTrue(deletedRestored.Succeeded && deletedRestored.Changed, deletedRestored.Message);
        File.Delete(managedBackupPath);

        foreach (var insertion in automaticPlacement.AppliedInsertions.Reverse())
        {
          var cleanup = RunExcelSta(() => rowService.DeleteRowsIfSafe(
            identity,
            insertion.WorksheetName,
            insertion.StartRow,
            insertion.Count));
          Assert.IsTrue(cleanup.Succeeded && cleanup.Changed, cleanup.Message);
        }
      }
      finally
      {
        Release(automaticShape);
        Release(automaticShapes);
        Release(restoredWorkbook);
        Release(originalOtherCell);
        Release(automaticTargetCell);
      }

      object? protectedShapes = null;
      try
      {
        protectedShapes = GetRequiredProperty(worksheet, "Shapes");
        var shapeCountBeforeProtection = Convert.ToInt32(
          GetRequiredProperty(protectedShapes, "Count"),
          CultureInfo.InvariantCulture);
        _ = InvokeMethod(worksheet, "Protect");
        var protectedPlacement = RunExcelSta(() => new ExcelImagePlacementService().PlaceImage(
          identity,
          "FocusTarget",
          new CellReference(5, 4),
          EvidenceSide.New,
          placementImagePath,
          new ImageDimensions(120, 60)));
        Assert.IsFalse(protectedPlacement.Succeeded, "A protected Worksheet must reject image placement.");
        Assert.AreEqual(
          shapeCountBeforeProtection,
          Convert.ToInt32(GetRequiredProperty(protectedShapes, "Count"), CultureInfo.InvariantCulture),
          "A protected Worksheet must not receive a Shape.");
        Assert.IsTrue(
          protectedPlacement.Message.Contains("保護", StringComparison.Ordinal),
          protectedPlacement.Message);
      }
      finally
      {
        TryInvoke(worksheet, "Unprotect");
        Release(protectedShapes);
      }

      var placementService = new ExcelImagePlacementService();
      var placement = RunExcelSta(() => placementService.PlaceImage(
        identity,
        "FocusTarget",
        new CellReference(5, 4),
        EvidenceSide.New,
        placementImagePath,
        new ImageDimensions(120, 60)));
      Assert.IsTrue(placement.Succeeded, placement.Message);
      Assert.IsTrue(placement.FocusSucceeded, placement.Message);
      Assert.AreEqual(5, placement.FocusCell.Row);
      Assert.AreEqual(4, placement.FocusCell.Column);
      object? placementShapes = null;
      object? placedShape = null;
      try
      {
        placementShapes = GetRequiredProperty(worksheet, "Shapes");
        var shapeCount = Convert.ToInt32(
          GetRequiredProperty(placementShapes, "Count"),
          CultureInfo.InvariantCulture);
        Assert.IsGreaterThan(0, shapeCount, "The placement did not create an Excel Shape.");
        placedShape = InvokeMethod(placementShapes, "Item", shapeCount) ??
          throw new InvalidOperationException("The inserted Shape could not be retrieved.");
        Assert.AreEqual(placement.ShapeName, Convert.ToString(
          GetRequiredProperty(placedShape!, "Name"),
          CultureInfo.InvariantCulture));
        Assert.IsTrue(
          Convert.ToBoolean(GetRequiredProperty(placedShape!, "LockAspectRatio"), CultureInfo.InvariantCulture),
          "The inserted image Shape must preserve its aspect ratio.");
      }
      finally
      {
        Release(placedShape);
        Release(placementShapes);
      }

      var placementUndo = RunExcelSta(() => placementService.DeletePlacedImage(
        identity,
        placement.WorksheetName,
        placement.ShapeName));
      Assert.IsTrue(placementUndo.Succeeded, placementUndo.Message);
      placement = RunExcelSta(() => placementService.PlaceImage(
        identity,
        "FocusTarget",
        new CellReference(5, 4),
        EvidenceSide.New,
        placementImagePath,
        new ImageDimensions(120, 60)));
      Assert.IsTrue(placement.Succeeded, placement.Message);

      Release(activeCell);
      activeCell = null;
      Release(activeWorkbook);
      activeWorkbook = null;
      _ = NativeMethods.ShowWindow(identity.ExcelWindowHandle, NativeMethods.MinimizeWindowCommand);
      Assert.IsTrue(NativeMethods.IsIconic(identity.ExcelWindowHandle));

      Assert.IsTrue(ExcelPlacementFocusService.BringToForeground(identity));
      Assert.IsTrue(NativeMethods.IsZoomed(identity.ExcelWindowHandle));

      var minimizedFocus = RunExcelSta(() => new ExcelPlacementFocusService().FocusPlacedImage(
        identity,
        "FocusTarget",
        new CellReference(6, 5)));
      Assert.IsTrue(minimizedFocus.Succeeded, minimizedFocus.Message);
      Assert.IsTrue(Convert.ToBoolean(
        GetRequiredProperty(application, "EnableEvents"),
        CultureInfo.InvariantCulture));

      Release(worksheet);
      worksheet = null;
      Release(workbook);
      workbook = null;
      Release(workbooks);
      workbooks = GetRequiredProperty(application, "Workbooks");
      workbookForCancelledClose = GetRequiredProperty(workbooks, "Item", "focus-integration.xlsx");
      var connectionPointContainer = (IConnectionPointContainer)application;
      var appEventsInterface = new Guid("00024413-0000-0000-C000-000000000046");
      connectionPointContainer.FindConnectionPoint(ref appEventsInterface, out cancelConnectionPoint);
      Assert.IsNotNull(cancelConnectionPoint, "Excel AppEvents connection point was not found for cancellation testing.");
      cancelCloseSink = new ExcelApplicationSessionMonitor.AppEventsSink((object candidate, ref bool cancel) =>
      {
        var fullPath = Convert.ToString(
          GetRequiredProperty(candidate, "FullName"),
          CultureInfo.CurrentCulture);
        if (string.Equals(fullPath, workbookPath, StringComparison.OrdinalIgnoreCase))
        {
          cancel = true;
        }
      });
      cancelConnectionPoint.Advise(cancelCloseSink, out cancelConnectionCookie);
      _ = InvokeMethod(workbookForCancelledClose, "Close", false);
      Assert.AreEqual(
        2,
        Convert.ToInt32(GetRequiredProperty(workbooks, "Count"), CultureInfo.InvariantCulture),
        "The cancellation sink did not keep the target Workbook open.");
      Release(workbookForCancelledClose);
      workbookForCancelledClose = null;
      Release(workbooks);
      workbooks = null;
      Thread.Sleep(300);
      Assert.IsTrue(
        WorkbookSessionTokenRegistry.IsValid(identity.WindowSessionToken),
        "A cancelled close must preserve the current Workbook session token.");
      var cancelledCloseFocus = RunExcelSta(() => new ExcelPlacementFocusService().FocusPlacedImage(
        identity,
        "FocusTarget",
        new CellReference(6, 5)));
      Assert.IsTrue(cancelledCloseFocus.Succeeded, cancelledCloseFocus.Message);

      var cancelledSessionToken = identity.WindowSessionToken;
      SetProperty(application, "EnableEvents", false);
      var cancelledCloseDisabledWarnings = sessionMonitor.Refresh(
        discovery.Workbooks.Where(item => item.ProcessId == identity.ProcessId).ToArray());
      Assert.IsTrue(
        cancelledCloseDisabledWarnings.Any(warning =>
          warning.Contains("events are disabled", StringComparison.OrdinalIgnoreCase)),
        string.Join(" | ", cancelledCloseDisabledWarnings));
      Assert.IsTrue(
        WorkbookSessionTokenRegistry.IsValid(cancelledSessionToken),
        "Optional monitoring failure must preserve a directly verifiable open Workbook token.");

      SetProperty(application, "EnableEvents", true);
      var recoveredAfterCancelDiscovery = RunExcelSta(() => new ExcelSessionCatalog().Discover());
      var recoveredAfterCancelIdentity = recoveredAfterCancelDiscovery.Workbooks.Single(item =>
        string.Equals(item.FullPath, workbookPath, StringComparison.OrdinalIgnoreCase));
      Assert.AreEqual(
        cancelledSessionToken,
        recoveredAfterCancelIdentity.WindowSessionToken,
        "Rediscovery must preserve the open Workbook window token when only monitoring failed.");
      var recoveredAfterCancelWarnings = sessionMonitor.Refresh(
        recoveredAfterCancelDiscovery.Workbooks
          .Where(item => item.ProcessId == recoveredAfterCancelIdentity.ProcessId)
          .ToArray());
      Assert.HasCount(0, recoveredAfterCancelWarnings, string.Join(" | ", recoveredAfterCancelWarnings));
      Thread.Sleep(750);
      Assert.IsTrue(
        WorkbookSessionTokenRegistry.IsValid(recoveredAfterCancelIdentity.WindowSessionToken),
        "The stale cancelled-close watcher must not invalidate the newly registered session token.");
      var recoveredAfterCancelFocus = RunExcelSta(() => new ExcelPlacementFocusService().FocusPlacedImage(
        recoveredAfterCancelIdentity,
        "FocusTarget",
        new CellReference(6, 5)));
      Assert.IsTrue(recoveredAfterCancelFocus.Succeeded, recoveredAfterCancelFocus.Message);
      identity = recoveredAfterCancelIdentity;

      cancelConnectionPoint.Unadvise(cancelConnectionCookie);
      cancelConnectionCookie = 0;
      ReleaseOnce(cancelConnectionPoint);
      cancelConnectionPoint = null;
      cancelCloseSink = null;

      workbooks = GetRequiredProperty(application, "Workbooks");
      var workbookToClose = GetRequiredProperty(workbooks, "Item", "focus-integration.xlsx");
      var closeEventCountBeforeConfirmedClose = sessionMonitor.CloseEventCount;
      _ = InvokeMethod(workbookToClose, "Close", false);
      Release(workbookToClose);
      Assert.AreEqual(
        1,
        Convert.ToInt32(GetRequiredProperty(workbooks, "Count"), CultureInfo.InvariantCulture),
        "The target Workbook did not close.");
      Assert.AreEqual(
        closeEventCountBeforeConfirmedClose + 1,
        sessionMonitor.CloseEventCount,
        "WorkbookBeforeClose monitoring did not receive the confirmed close event.");
      Assert.IsTrue(
        SpinWait.SpinUntil(
          () => !WorkbookSessionTokenRegistry.IsValid(identity.WindowSessionToken),
          TimeSpan.FromSeconds(5)),
        "The closed Workbook session token was not invalidated after its document window disappeared.");

      workbook = InvokeMethod(workbooks, "Open", workbookPath) ??
        throw new InvalidOperationException("Excel did not reopen the target Workbook.");
      worksheet = GetRequiredProperty(application, "ActiveSheet");
      var rediscovery = RunExcelSta(() => new ExcelSessionCatalog().Discover());
      var reopenedIdentity = rediscovery.Workbooks.SingleOrDefault(item =>
        string.Equals(item.FullPath, workbookPath, StringComparison.OrdinalIgnoreCase));
      Assert.IsNotNull(reopenedIdentity, "The reopened Workbook was not rediscovered.");
      Assert.AreNotEqual(
        identity.ConnectionId,
        reopenedIdentity.ConnectionId,
        "Closing and reopening the same path must create a new Workbook session identity.");
      Assert.AreNotEqual(
        identity.WindowSessionToken,
        reopenedIdentity.WindowSessionToken,
        "Closing and reopening the same path must create a new Workbook window session token.");

      var staleFocus = RunExcelSta(() => new ExcelPlacementFocusService().FocusPlacedImage(
        identity,
        "FocusTarget",
        new CellReference(7, 6)));
      Assert.IsFalse(staleFocus.Succeeded, "A closed Workbook identity must not attach to a reopened Workbook.");

      var reopenedFocus = RunExcelSta(() => new ExcelPlacementFocusService().FocusPlacedImage(
        reopenedIdentity,
        "FocusTarget",
        new CellReference(7, 6)));
      Assert.IsTrue(reopenedFocus.Succeeded, reopenedFocus.Message);

      SetProperty(application, "EnableEvents", false);
      var disabledWarnings = sessionMonitor.Refresh(
        rediscovery.Workbooks.Where(item => item.ProcessId == reopenedIdentity.ProcessId).ToArray());
      Assert.IsTrue(
        disabledWarnings.Any(warning => warning.Contains("events are disabled", StringComparison.OrdinalIgnoreCase)),
        string.Join(" | ", disabledWarnings));
      var disabledFocus = RunExcelSta(() => new ExcelPlacementFocusService().FocusPlacedImage(
        reopenedIdentity,
        "FocusTarget",
        new CellReference(8, 7)));
      Assert.IsTrue(disabledFocus.Succeeded, disabledFocus.Message);
      Assert.IsFalse(
        Convert.ToBoolean(GetRequiredProperty(application, "EnableEvents"), CultureInfo.InvariantCulture),
        "Focus must preserve the caller's disabled Excel event state.");
      Assert.IsTrue(
        WorkbookSessionTokenRegistry.IsValid(reopenedIdentity.WindowSessionToken),
        "Disabled Excel events must not invalidate an otherwise open Workbook connection.");
      var disabledPlacement = RunExcelSta(() => new ExcelImagePlacementService().PlaceImage(
        reopenedIdentity,
        "FocusTarget",
        new CellReference(9, 4),
        EvidenceSide.New,
        placementImagePath,
        new ImageDimensions(120, 60)));
      Assert.IsTrue(disabledPlacement.Succeeded, disabledPlacement.Message);
      Assert.IsTrue(disabledPlacement.FocusSucceeded, disabledPlacement.Message);
      Assert.IsFalse(
        Convert.ToBoolean(GetRequiredProperty(application, "EnableEvents"), CultureInfo.InvariantCulture),
        "Placement must preserve the caller's disabled Excel event state.");
      var disabledPlacementUndo = RunExcelSta(() => new ExcelImagePlacementService().DeletePlacedImage(
        reopenedIdentity,
        disabledPlacement.WorksheetName,
        disabledPlacement.ShapeName));
      Assert.IsTrue(disabledPlacementUndo.Succeeded, disabledPlacementUndo.Message);

      SetProperty(application, "EnableEvents", true);
      var enabledDiscovery = RunExcelSta(() => new ExcelSessionCatalog().Discover());
      var enabledIdentity = enabledDiscovery.Workbooks.Single(item =>
        string.Equals(item.FullPath, workbookPath, StringComparison.OrdinalIgnoreCase));
      var enabledWarnings = sessionMonitor.Refresh(
        enabledDiscovery.Workbooks.Where(item => item.ProcessId == enabledIdentity.ProcessId).ToArray());
      Assert.HasCount(0, enabledWarnings, string.Join(" | ", enabledWarnings));
      var enabledFocus = RunExcelSta(() => new ExcelPlacementFocusService().FocusPlacedImage(
        enabledIdentity,
        "FocusTarget",
        new CellReference(8, 7)));
      Assert.IsTrue(enabledFocus.Succeeded, enabledFocus.Message);
      }
    }
    catch (Exception exception)
    {
      scenarioFailure = exception;
    }
    finally
    {
      if (!supervisor.ShouldSkipComCleanup && cancelConnectionPoint is not null && cancelConnectionCookie != 0)
      {
        TryCleanup(
          () => cancelConnectionPoint.Unadvise(cancelConnectionCookie),
          cleanupFailures);
      }

      if (!supervisor.ShouldSkipComCleanup)
      {
        TryCleanup(() => ReleaseOnce(cancelConnectionPoint), cleanupFailures);
        cancelConnectionPoint = null;
        TryCleanup(() => sessionMonitor?.Dispose(), cleanupFailures);
        sessionMonitor = null;
        cancelCloseSink = null;
        TryCleanup(() => Release(activeCell), cleanupFailures);
        TryCleanup(() => Release(activeWorkbook), cleanupFailures);
        TryCleanup(() => Release(otherAnchorCell), cleanupFailures);
        TryCleanup(() => Release(editedCell), cleanupFailures);
        TryCleanup(() => Release(workbookForCancelledClose), cleanupFailures);
        if (ownsExcelProcess && otherWorkbook is not null)
        {
          TryCleanup(() => TryInvoke(otherWorkbook, "Close", false), cleanupFailures);
        }

        if (ownsExcelProcess && workbook is not null)
        {
          TryCleanup(() => TryInvoke(workbook, "Close", false), cleanupFailures);
        }

        TryCleanup(() => Release(otherWorksheet), cleanupFailures);
        TryCleanup(() => Release(otherWorkbook), cleanupFailures);
        TryCleanup(() => Release(worksheet), cleanupFailures);
        TryCleanup(() => Release(workbook), cleanupFailures);
        TryCleanup(() => Release(workbooks), cleanupFailures);
        if (ownsExcelProcess && application is not null)
        {
          TryCleanup(() => TryInvoke(application, "Quit"), cleanupFailures);
        }

        TryCleanup(() => Release(application), cleanupFailures);
        application = null;
        if (ownsExcelProcess && excelProcess is not null)
        {
          TryCleanup(() => EnsureExcelProcessExited(excelProcess), cleanupFailures);
        }
      }
      else
      {
        TryCleanup(() => EnsureSupervisedExcelProcessesExited(supervisor), cleanupFailures);
      }

      if (!supervisor.ShouldSkipComCleanup && ownsExcelProcess && excelProcess is null)
      {
        TryCleanup(() => EnsureSupervisedExcelProcessesExited(supervisor), cleanupFailures);
      }

      if (File.Exists(workbookPath))
      {
        TryCleanup(() => RetryIo(() => File.Delete(workbookPath)), cleanupFailures);
      }

      if (File.Exists(otherWorkbookPath))
      {
        TryCleanup(() => RetryIo(() => File.Delete(otherWorkbookPath)), cleanupFailures);
      }

      if (File.Exists(placementImagePath))
      {
        TryCleanup(() => RetryIo(() => File.Delete(placementImagePath)), cleanupFailures);
      }

      if (Directory.Exists(temporaryDirectory))
      {
        TryCleanup(() => RetryIo(() => Directory.Delete(temporaryDirectory)), cleanupFailures);
      }
    }

    if (scenarioFailure is not null)
    {
      if (cleanupFailures.Count > 0)
      {
        throw new AggregateException(
          "The Excel integration scenario and its cleanup both failed.",
          [scenarioFailure, .. cleanupFailures]);
      }

      ExceptionDispatchInfo.Capture(scenarioFailure).Throw();
    }

    if (cleanupFailures.Count > 0)
    {
      throw new AggregateException("Excel integration cleanup failed.", cleanupFailures);
    }
  }

  private static object GetRequiredProperty(object target, string propertyName, params object[]? arguments) =>
    target.GetType().InvokeMember(
      propertyName,
      BindingFlags.GetProperty,
      binder: null,
      target,
      arguments,
      CultureInfo.CurrentCulture) ??
    throw new InvalidOperationException($"COM property returned null: {propertyName}");

  private static T RunOnSta<T>(
    Func<T> action,
    Func<Exception?>? onTimeout = null,
    Action? onThreadStillRunning = null)
  {
    T? result = default;
    Exception? failure = null;
    var completed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
    var thread = new Thread(() =>
    {
      try
      {
        result = action();
      }
      catch (Exception exception)
      {
        failure = exception;
      }
      finally
      {
        completed.TrySetResult();
      }
    })
    {
      IsBackground = true,
    };
    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    if (!completed.Task.Wait(TimeSpan.FromSeconds(20)))
    {
      Exception? terminationFailure = null;
      try
      {
        terminationFailure = onTimeout?.Invoke();
      }
      catch (Exception exception) when (exception is not OutOfMemoryException)
      {
        terminationFailure = exception;
      }

      var joined = thread.Join(TimeSpan.FromSeconds(5));
      if (!joined)
      {
        onThreadStillRunning?.Invoke();
      }

      throw new StaOperationTimeoutException(joined, terminationFailure);
    }

    thread.Join();
    if (failure is not null)
    {
      throw new InvalidOperationException("The Excel STA operation failed.", failure);
    }

    return result!;
  }

  private sealed class StaOperationTimeoutException : TimeoutException
  {
    public StaOperationTimeoutException(bool staStopped, Exception? terminationFailure)
      : base(CreateMessage(staStopped, terminationFailure), terminationFailure)
    {
      StaStopped = staStopped;
    }

    public bool StaStopped { get; }

    private static string CreateMessage(bool staStopped, Exception? terminationFailure) =>
      staStopped
        ? terminationFailure is null
          ? "The Excel STA operation timed out after its supervised Excel process was terminated."
          : $"The Excel STA operation timed out; its supervised Excel process could not be terminated ({terminationFailure.Message})."
        : terminationFailure is null
          ? "The Excel STA operation timed out and its STA did not terminate after process supervision."
          : $"The Excel STA operation timed out; its STA did not terminate and Excel termination also failed ({terminationFailure.Message}).";
  }

  private static void RetryIo(Action action)
  {
    for (var attempt = 1; ; attempt++)
    {
      try
      {
        action();
        return;
      }
      catch (IOException) when (attempt < 50)
      {
        Thread.Sleep(100);
      }
      catch (UnauthorizedAccessException) when (attempt < 50)
      {
        Thread.Sleep(100);
      }
    }
  }

  private static object? InvokeMethod(object target, string methodName, params object[] arguments) =>
    target.GetType().InvokeMember(
      methodName,
      BindingFlags.InvokeMethod,
      binder: null,
      target,
      arguments,
      CultureInfo.CurrentCulture);

  private static void SetCellValue(object worksheet, int row, int column, object value)
  {
    object? cell = null;
    try
    {
      cell = GetRequiredProperty(worksheet, "Cells", row, column);
      SetProperty(cell, "Value2", value);
    }
    finally
    {
      Release(cell);
    }
  }

  private static void SetRangeBorder(
    object worksheet,
    int firstRow,
    int firstColumn,
    int lastRow,
    int lastColumn,
    int borderIndex)
  {
    object? firstCell = null;
    object? lastCell = null;
    object? range = null;
    object? borders = null;
    object? border = null;
    try
    {
      firstCell = GetRequiredProperty(worksheet, "Cells", firstRow, firstColumn);
      lastCell = GetRequiredProperty(worksheet, "Cells", lastRow, lastColumn);
      range = GetRequiredProperty(worksheet, "Range", firstCell, lastCell);
      borders = GetRequiredProperty(range, "Borders");
      border = GetRequiredProperty(borders, "Item", borderIndex);
      SetProperty(border, "LineStyle", 1);
    }
    finally
    {
      Release(border);
      Release(borders);
      Release(range);
      Release(lastCell);
      Release(firstCell);
    }
  }

  private static void SetProperty(object target, string propertyName, object value) =>
    _ = target.GetType().InvokeMember(
      propertyName,
      BindingFlags.SetProperty,
      binder: null,
      target,
      [value]);

  private static void TryInvoke(object target, string methodName, params object[] arguments)
  {
    try
    {
      _ = InvokeMethod(target, methodName, arguments);
    }
    catch (Exception exception) when (
      exception is COMException or TargetInvocationException or InvalidComObjectException)
    {
      // Best-effort cleanup must not hide the original assertion or automation failure.
    }
  }

  private static void Release(object? value)
  {
    if (value is not null && Marshal.IsComObject(value))
    {
      _ = Marshal.FinalReleaseComObject(value);
    }
  }

  private static void ReleaseOnce(object? value)
  {
    if (value is not null && Marshal.IsComObject(value))
    {
      _ = Marshal.ReleaseComObject(value);
    }
  }

  private static void TryCleanup(Action action, ICollection<Exception> failures)
  {
    try
    {
      action();
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      failures.Add(exception);
    }
  }

  private static Dictionary<int, DateTime> GetExcelProcessStartTimes()
  {
    var processes = Process.GetProcessesByName("EXCEL");
    try
    {
      var startTimes = new Dictionary<int, DateTime>();
      foreach (var process in processes)
      {
        try
        {
          startTimes[process.Id] = process.StartTime;
        }
        catch (InvalidOperationException)
        {
          // A process that exited while the baseline was captured is not a protected process.
        }
        catch (System.ComponentModel.Win32Exception)
        {
          // A process that exited while the baseline was captured is not a protected process.
        }
      }

      return startTimes;
    }
    finally
    {
      foreach (var process in processes)
      {
        process.Dispose();
      }
    }
  }

  private static void EnsureSupervisedExcelProcessesExited(ScenarioSupervisor supervisor)
  {
    if (!supervisor.TryTerminate(out var failure))
    {
      throw new InvalidOperationException(
        "A generated Excel process remained after integration cleanup.",
        failure);
    }
  }

  private static void EnsureExcelProcessExited(Process process)
  {
    if (process.WaitForExit(milliseconds: 10_000))
    {
      return;
    }

    process.Kill(entireProcessTree: true);
    if (!process.WaitForExit(milliseconds: 5_000))
    {
      throw new InvalidOperationException(
        $"Generated Excel process {process.Id} could not be terminated after the integration scenario.");
    }

    throw new InvalidOperationException(
      $"Generated Excel process {process.Id} required forced termination after the integration scenario.");
  }

  private sealed class ScenarioSupervisor : IDisposable
  {
    private readonly Lock stateLock = new();
    private readonly Dictionary<int, DateTime> preexistingProcessStartTimes;
    private readonly List<Process> generatedExcelProcesses = [];
    private readonly DateTime startedAt = DateTime.Now;
    private bool suppressComCleanup;

    public ScenarioSupervisor(IReadOnlyDictionary<int, DateTime> preexistingProcessStartTimes)
    {
      this.preexistingProcessStartTimes = new(preexistingProcessStartTimes);
    }

    public bool ShouldSkipComCleanup
    {
      get
      {
        lock (stateLock)
        {
          return suppressComCleanup;
        }
      }
    }

    public bool IsPreexistingProcess(uint processId)
    {
      try
      {
        using var process = Process.GetProcessById(checked((int)processId));
        lock (stateLock)
        {
          return preexistingProcessStartTimes.TryGetValue(process.Id, out var startTime) &&
            process.StartTime == startTime;
        }
      }
      catch (ArgumentException)
      {
        return false;
      }
      catch (InvalidOperationException)
      {
        return false;
      }
      catch (System.ComponentModel.Win32Exception)
      {
        return false;
      }
    }

    public void SuppressComCleanup()
    {
      lock (stateLock)
      {
        suppressComCleanup = true;
      }
    }

    public void Capture(Process process)
    {
      lock (stateLock)
      {
        if (generatedExcelProcesses.Any(candidate => candidate.Id == process.Id))
        {
          process.Dispose();
        }
        else
        {
          generatedExcelProcesses.Add(process);
        }
      }
    }

    public bool TryTerminate(out Exception? failure)
    {
      CaptureNewProcesses();
      Process[] processes;
      lock (stateLock)
      {
        processes = generatedExcelProcesses.ToArray();
      }

      var failures = new List<Exception>();
      foreach (var process in processes)
      {
        try
        {
          if (process.HasExited)
          {
            continue;
          }

          process.Kill(entireProcessTree: true);
          if (!process.WaitForExit(milliseconds: 5_000))
          {
            failures.Add(new InvalidOperationException(
              $"Generated Excel process {process.Id} did not exit after forced termination."));
          }
        }
        catch (Exception exception) when (exception is not OutOfMemoryException)
        {
          failures.Add(new InvalidOperationException(
            $"Generated Excel process {process.Id} could not be terminated.",
            exception));
        }
      }

      failure = failures.Count switch
      {
        0 => null,
        1 => failures[0],
        _ => new AggregateException("One or more generated Excel processes could not be terminated.", failures),
      };
      return failure is null;
    }

    private void CaptureNewProcesses()
    {
      var processes = Process.GetProcessesByName("EXCEL");
      foreach (var process in processes)
      {
        try
        {
          if (process.HasExited ||
            (preexistingProcessStartTimes.TryGetValue(process.Id, out var startTime) &&
              process.StartTime == startTime) ||
            process.StartTime < startedAt)
          {
            process.Dispose();
            continue;
          }

          Capture(process);
        }
        catch (Exception exception) when (
          exception is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception)
        {
          process.Dispose();
        }
      }
    }

    public void Dispose()
    {
      Process[] processes;
      lock (stateLock)
      {
        processes = generatedExcelProcesses.ToArray();
        generatedExcelProcesses.Clear();
      }

      foreach (var process in processes)
      {
        process.Dispose();
      }
    }
  }

  private sealed class OfficeUnavailableException(string message) : Exception(message);

  private static class NativeMethods
  {
    internal const int MinimizeWindowCommand = 6;

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool ShowWindow(nint windowHandle, int command);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool IsIconic(nint windowHandle);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool IsZoomed(nint windowHandle);

    [DllImport("user32.dll")]
    internal static extern uint GetWindowThreadProcessId(nint windowHandle, out uint processId);

  }
}
