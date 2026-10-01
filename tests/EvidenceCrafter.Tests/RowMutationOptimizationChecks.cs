using System.Globalization;
using System.Drawing.Imaging;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  private static void VerifyRowMutationOptimizations(object sheet, WorkbookIdentity identity, string imagePath)
  {
    using (var image = new Bitmap(120, 80)) image.Save(imagePath, ImageFormat.Png);
    var fingerprint = typeof(ExcelRowMutationService).GetMethod("CaptureRowFingerprint", BindingFlags.Static | BindingFlags.NonPublic)!;
    void Compare(object target, int start, int count, bool extent)
    {
      var expected = CaptureLegacyFingerprint(target, start, count, extent);
      Assert.AreEqual(expected, (string)fingerprint.Invoke(null, [target, start, count, extent])!);
    }
    SetRangeProperty(sheet, "A1:AF42", "NumberFormat", "0.00");
    SetRangeProperty(sheet, "A1:AF42", "RowHeight", 15);
    SetCellValue(sheet, 10, 3, "text");
    SetCellValue(sheet, 10, 4, 0);
    SetCellValue(sheet, 10, 5, false);
    SetCellValue(sheet, 11, 3, 12.75);
    var formulaCell = GetRequiredProperty(sheet, "Cells", 11, 4);
    try { SetProperty(formulaCell, "Formula", "=1/0"); }
    finally { Release(formulaCell); }
    Compare(sheet, 10, 2, false);
    Compare(sheet, 10, 2, true);
    SetRangeProperty(sheet, "D10:D11", "NumberFormat", "@");
    SetRangeProperty(sheet, "E10:E11", "WrapText", true);
    SetRangeProperty(sheet, "F10:F11", "Style", "Percent");
    Compare(sheet, 10, 2, true);
    Compare(sheet, 10, 1, false);
    var workbook = GetRequiredProperty(sheet, "Parent");
    var sheets = GetRequiredProperty(workbook, "Worksheets");
    var application = GetRequiredProperty(sheet, "Application");
    var oldSheetCount = GetRequiredProperty(application, "SheetsInNewWorkbook");
    try
    {
      var single = InvokeMethod(sheets, "Add")!;
      try
      {
        SetCellValue(single, 7, 26, false);
        Compare(single, 7, 1, false);
        Compare(single, 7, 1, true);
      }
      finally { _ = InvokeMethod(single, "Delete"); Release(single); }
      var values = GetRequiredProperty(sheet, "Range", "A1:AF42");
      try { _ = InvokeMethod(values, "ClearContents"); }
      finally { Release(values); }

      // Verify the fast filter against the complete, independent boundary oracle.
      SetCellValue(sheet, 2, 3, "NEW"); SetCellValue(sheet, 2, 18, "OLD");
      SetCellValue(sheet, 3, 1, 1); SetCellValue(sheet, 3, 2, 1);
      SetCellValue(sheet, 13, 1, 1); SetCellValue(sheet, 13, 2, 2);
      var shapes = GetRequiredProperty(sheet, "Shapes");
      try
      {
        foreach (var top in new[] { 0f, 28f, 40f, 165f, 180f, 400f })
        {
          var shape = InvokeMethod(shapes, "AddShape", 1, 100f, top, 120f, 35f)!;
          Release(shape);
        }
        var rotated = InvokeMethod(shapes, "AddShape", 1, 400f, 90f, 500f, 40f)!;
        try { SetProperty(rotated, "Rotation", 25f); }
        finally { Release(rotated); }
        foreach (var name in new[] { "GroupA", "GroupB" })
        {
          var member = InvokeMethod(shapes, "AddShape", 1, 700f, name == "GroupA" ? 110f : 140f, 40f, 20f)!;
          try { SetProperty(member, "Name", name); }
          finally { Release(member); }
        }
        var groupRange = GetRequiredProperty(shapes, "Range", (object)new[] { "GroupA", "GroupB" });
        try { Release(InvokeMethod(groupRange, "Group")); }
        finally { Release(groupRange); }
        SetRangeProperty(sheet, "9:9", "RowHeight", 15.75);
        foreach (var (side, cell) in new[] { (EvidenceSide.New, new CellReference(4, 4)), (EvidenceSide.Old, new CellReference(6, 19)) })
        {
          var image = new ExcelImagePlacementService().PlaceImage(identity, "OtherTarget", cell, side, imagePath, new ImageDimensions(120, 80), availableWidthPoints: 120);
          Assert.IsTrue(image.Succeeded, image.Message);
        }
        SetRangeProperty(sheet, "10:10", "Hidden", true);
        var full = new ExcelSheetSnapshotService().Capture(identity, "OtherTarget");
        var scoped = new ExcelSheetSnapshotService().Capture(identity, "OtherTarget", scopeCaseLabel: "1-1", scopeShapes: true);
        Assert.IsTrue(full.Succeeded, full.Message); Assert.IsTrue(scoped.Succeeded, scoped.Message);
        var scope = scoped.Snapshot!.ShapeScope!.Value;
        CollectionAssert.AreEquivalent(full.Snapshot!.Shapes.Where(item => item.EndRow >= scope.FirstRow && item.StartRow <= scope.LastRow).ToArray(), scoped.Snapshot.Shapes.ToArray());
        var outside = full.Snapshot.Shapes.First(item => item.StartRow > scope.LastRow);
        var referenced = new ExcelSheetSnapshotService().Capture(identity, "OtherTarget", scopeCaseLabel: "1-1", scopeShapes: true, referenceShapeName: outside.Name);
        Assert.IsTrue(referenced.Succeeded, referenced.Message);
        CollectionAssert.AreEquivalent(scoped.Snapshot.Shapes.Append(outside).ToArray(), referenced.Snapshot!.Shapes.ToArray());
        var safety = new ExcelRowMutationService().CaptureRowSafetyStates(identity, "OtherTarget", 3, 12);
        Assert.IsTrue(safety.Succeeded, safety.Message);
        foreach (var row in safety.Rows)
          Assert.AreEqual(full.Snapshot.Shapes.Any(item => item.StartRow <= row.Row && item.EndRow >= row.Row), row.HasShape);
        var layout = new CaseLayoutAnalyzer().Analyze(full.Snapshot.LayoutSignals with { ActiveRow = 3 }).Layout!;
        var managedSides = new HashSet<EvidenceSide>();
        var readShapes = typeof(ExcelRowMutationService).GetMethod("ReadRowsWithShapes", BindingFlags.Static | BindingFlags.NonPublic)!;
        var shapeRows = (HashSet<int>)readShapes.Invoke(null, [sheet, 3, 12, layout, managedSides])!;
        CollectionAssert.AreEquivalent(safety.Rows.Where(row => row.HasShape).Select(row => row.Row).ToArray(), shapeRows.ToArray());
        Assert.AreEqual(ExcelCaseMaintenanceService.HasBothSides(full.Snapshot, layout), managedSides.Count == 2);
      }
      finally { Release(shapes); }

      foreach (var fault in new[] { "Backup.Add", "Backup.Copy", "Backup.SaveAs", "Dependencies.Pre", "Rows.BeforeMutation", "Rows.AfterMutation", "PostDeleteFingerprint", "Dependencies.Post", "Rows.Delete.Unknown", "PostDelete.ManagedFailure", "Rows.FinalSafety.Block", "Rows.FinalSafety.Expand", "Paired.Insert.Unknown", "Paired.Delete.Unknown", "None" })
      {
        var target = InvokeMethod(sheets, "Add")!;
        RowDeletionSnapshot? backup = null;
        try
        {
          SetProperty(target, "Name", "FaultTarget");
          SetRangeProperty(target, "A1:F20", "NumberFormat", "0.00");
          SetCellValue(target, 3, 1, 1); SetCellValue(target, 3, 2, 1);
          SetCellValue(target, 13, 1, 1); SetCellValue(target, 13, 2, 2);
          SetCellValue(target, 16, 3, "shifted marker");
          SetRangeProperty(target, "A9:F9", "RowHeight", 31.5);
          SetProperty(application, "SheetsInNewWorkbook", 3);
          var service = new ExcelRowMutationService();
          if (fault.StartsWith("Paired.", StringComparison.Ordinal))
          {
            var placed = new ExcelImagePlacementService().PlaceImage(identity, "FaultTarget", new CellReference(4, 4), EvidenceSide.New, imagePath, new ImageDimensions(120, 80), availableWidthPoints: 120);
            Assert.IsTrue(placed.Succeeded, placed.Message);
            var before = new ExcelManagedShapeService().Inspect(identity, "FaultTarget", placed.ShapeName!).Shape!;
            var reference = new PairedImageResize(before, before, new AppliedRowInsertion("FaultTarget", 20, 2, "Fault check"));
            ExcelRowMutationService.StageReached = stage =>
            {
              if (stage == "Rows.AfterMutation") throw new COMException("Injected ambiguous paired row change", unchecked((int)0x80010001));
            };
            var paired = reference.SetApplied(identity, fault == "Paired.Insert.Unknown");
            Assert.IsFalse(paired.Succeeded);
            Assert.IsTrue(paired.MutationMayHaveOccurred, paired.Message);
            Assert.IsFalse(reference.CompensationSucceeded);
            Assert.IsFalse(reference.CanRetryPreparation);
            continue;
          }
          ExcelRowMutationService.StageReached = stage =>
          {
            if (fault == "Rows.FinalSafety.Block" && stage == "Rows.FinalSafety") SetCellValue(target, 9, 6, "written during backup");
            if (fault == "Rows.FinalSafety.Expand" && stage == "Rows.FinalSafety")
            {
              var inserted = GetRequiredProperty(target, "Rows", 9);
              try { _ = InvokeMethod(inserted, "Insert"); }
              finally { Release(inserted); }
            }
            if (fault == "PostDelete.ManagedFailure" && stage == "PostDeleteFingerprint") throw new InvalidDataException("Injected managed failure");
            if (stage == fault) throw new COMException("Injected " + fault, unchecked((int)0x80010001));
          };
          ExcelRowMutationService.StageMeasured = (stage, _) =>
          {
            if (fault == "Rows.Delete.Unknown" && stage == "Rows.Delete") throw new COMException("Ambiguous Delete result", unchecked((int)0x80010001));
          };
          var result = service.DeleteTrailingRowsWithSnapshot(identity, "FaultTarget", 3, 12, 4);
          backup = result.DeletionSnapshot;
          ExcelRowMutationService.StageReached = null; ExcelRowMutationService.StageMeasured = null;
          if (fault is "Rows.FinalSafety.Block" or "Rows.FinalSafety.Expand")
          {
            Assert.IsTrue(result.Succeeded && !result.Changed, result.Message);
            Assert.IsFalse(result.MutationMayHaveOccurred);
            Assert.IsNull(backup);
            var marker = GetRequiredProperty(target, "Cells", fault == "Rows.FinalSafety.Expand" ? 17 : 16, 3);
            try { Assert.AreEqual("shifted marker", GetRequiredProperty(marker, "Value2")); }
            finally { Release(marker); }
            continue;
          }
          if (fault == "None")
          {
            Assert.IsTrue(result.Succeeded && result.Changed, result.Message);
            Assert.IsNotNull(backup);
            var postDelete = CaptureLegacyFingerprint(target, 3, 18, false);
            var workbooks = GetRequiredProperty(application, "Workbooks");
            var native = InvokeMethod(workbooks, "Open", backup.BackupPath, 0, true)!;
            try
            {
              var nativeSheets = GetRequiredProperty(native, "Worksheets");
              try { Assert.AreEqual(1, Convert.ToInt32(GetRequiredProperty(nativeSheets, "Count"))); }
              finally { Release(nativeSheets); }
            }
            finally { _ = InvokeMethod(native, "Close", false); Release(native); ReleaseOnce(workbooks); }
            Assert.IsTrue(service.RestoreDeletedRows(identity, backup).Succeeded);
            var restoredRow = GetRequiredProperty(target, "Rows", 9);
            try { Assert.AreEqual(31.5, Convert.ToDouble(GetRequiredProperty(restoredRow, "RowHeight")), 0.05); }
            finally { Release(restoredRow); }
            var redo = service.DeleteRestoredRows(identity, backup);
            Assert.IsTrue(redo.Succeeded && redo.Changed, redo.Message);
            Assert.AreEqual(postDelete, CaptureLegacyFingerprint(target, 3, 18, false), "Redo must preserve the shifted next CASE and marker.");
          }
          else
          {
            Assert.IsFalse(result.Succeeded);
            var after = fault is "Rows.AfterMutation" or "PostDeleteFingerprint" or "Dependencies.Post" or "Rows.Delete.Unknown" or "PostDelete.ManagedFailure";
            Assert.AreEqual(after, result.MutationMayHaveOccurred);
            if (after)
            {
              Assert.IsNotNull(backup, result.Message);
              Assert.IsTrue(backup.RecoveryRequired);
              backup.Dispose();
              Assert.IsTrue(File.Exists(backup.RecoveryBackupPath));
              Assert.IsTrue(File.Exists(backup.BackupPath + ".recovery.json"));
              Assert.IsFalse(service.RestoreDeletedRows(identity, backup).Succeeded);
              Assert.IsFalse(service.DeleteRestoredRows(identity, backup).Succeeded);
              Assert.AreEqual(fault != "Rows.Delete.Unknown", result.Changed);
            }
            else Assert.IsNull(backup);
          }
        }
        finally
        {
          ExcelRowMutationService.StageReached = null; ExcelRowMutationService.StageMeasured = null;
          if (backup is not null)
          {
            backup.Dispose();
            if (backup.RecoveryRequired) { File.Delete(backup.BackupPath); File.Delete(backup.BackupPath + ".recovery.json"); }
          }
          _ = InvokeMethod(target, "Delete"); Release(target);
        }
      }
    }
    finally
    {
      SetProperty(application, "SheetsInNewWorkbook", oldSheetCount);
      // Parent/Application are aliases of the supervisor's live wrappers.
      ReleaseOnce(application); Release(sheets); ReleaseOnce(workbook);
    }
  }
  private static bool TryReadFingerprintProperty(object target, string property, out object? value)
  {
    try { value = GetRequiredProperty(target, property); return true; }
    catch (Exception exception) when (exception is COMException or TargetInvocationException or MissingMemberException or InvalidOperationException)
    { value = null; return false; }
  }
  private static string CaptureLegacyFingerprint(
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
      if ((long)count * columnCount > 250_000)
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
          wholeRow = GetRequiredProperty(rows, "Item", row);
          AppendLegacyFingerprintProperty(builder, wholeRow, "RowHeight");
          AppendLegacyFingerprintProperty(builder, wholeRow, "Hidden");
          AppendLegacyFingerprintProperty(builder, wholeRow, "OutlineLevel");
          AppendLegacyFingerprintProperty(builder, wholeRow, "PageBreak");
        }
        finally
        {
          Release(wholeRow);
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
            cell = GetRequiredProperty(worksheet, "Cells", row, column);
            AppendLegacyFingerprintProperty(builder, cell, "Value2");
            AppendLegacyFingerprintProperty(builder, cell, "Formula");
            AppendLegacyFingerprintProperty(builder, cell, "NumberFormat");
            AppendLegacyFingerprintProperty(builder, cell, "Style");
            AppendLegacyFingerprintProperty(builder, cell, "WrapText");
            AppendLegacyFingerprintProperty(builder, cell, "HorizontalAlignment");
            AppendLegacyFingerprintProperty(builder, cell, "VerticalAlignment");
            if (TryReadFingerprintProperty(cell!, "Font", out font) && font is not null)
            {
              AppendLegacyFingerprintProperty(builder, font, "Name");
              AppendLegacyFingerprintProperty(builder, font, "Size");
              AppendLegacyFingerprintProperty(builder, font, "Bold");
              AppendLegacyFingerprintProperty(builder, font, "Italic");
              AppendLegacyFingerprintProperty(builder, font, "Color");
            }
            if (TryReadFingerprintProperty(cell!, "Interior", out interior) && interior is not null)
            {
              AppendLegacyFingerprintProperty(builder, interior, "Color");
              AppendLegacyFingerprintProperty(builder, interior, "Pattern");
            }
            if (TryReadFingerprintProperty(cell!, "Validation", out validation) && validation is not null)
            {
              AppendLegacyFingerprintProperty(builder, validation, "Type");
              AppendLegacyFingerprintProperty(builder, validation, "Formula1");
              AppendLegacyFingerprintProperty(builder, validation, "Formula2");
            }
            if (TryReadFingerprintProperty(cell!, "FormatConditions", out conditions) && conditions is not null)
            {
              AppendLegacyFingerprintProperty(builder, conditions, "Count");
            }
          }
          finally
          {
            Release(conditions);
            Release(validation);
            Release(interior);
            Release(font);
            Release(cell);
          }
        }
      }

      return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(builder.ToString())));
    }
    finally
    {
      Release(rows);
      Release(usedRows);
      Release(usedColumns);
      Release(usedRange);
    }
  }

  private static void AppendLegacyFingerprintProperty(StringBuilder builder, object? target, string propertyName)
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
    catch (Exception exception) when (exception is COMException or TargetInvocationException or MissingMemberException or InvalidOperationException)
    {
      builder.Append("<unsupported>;");
    }
  }

}
