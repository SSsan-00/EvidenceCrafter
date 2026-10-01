using System.Diagnostics;
using System.Drawing.Imaging;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  private static void VerifySameSideBackfill(object sheet, WorkbookIdentity identity, string imagePath)
  {
    SetRangeProperty(sheet, "A1:AF202", "RowHeight", 15.75);
    SetRangeProperty(sheet, "A1:AF202", "NumberFormat", "0.00");
    var pageSetup = GetRequiredProperty(sheet, "PageSetup");
    try { SetProperty(pageSetup, "PrintArea", "$A$1:$AF$202"); }
    finally { Release(pageSetup); }
    SetRangeProperty(sheet, "20:21", "RowHeight", 24);
    var hiddenCell = GetRequiredProperty(sheet, "Cells", 30, 1);
    var hiddenRow = GetRequiredProperty(hiddenCell, "EntireRow");
    try { SetProperty(hiddenRow, "Hidden", true); }
    finally { Release(hiddenRow); Release(hiddenCell); }
    SetCellValue(sheet, 2, 3, "NEW");
    SetCellValue(sheet, 2, 18, "OLD");
    for (var index = 0; index < 4; index++)
    {
      SetCellValue(sheet, 3 + index * 50, 1, 1);
      SetCellValue(sheet, 3 + index * 50, 2, index + 1);
    }
    using (var bitmap = new Bitmap(120, 80)) bitmap.Save(imagePath, ImageFormat.Png);
    var service = new ExcelAutomaticPlacementService();
    var snapshots = new ExcelSheetSnapshotService();
    var navigator = new ExcelCaseNavigationService();
    var images = new[] { new AutomaticPlacementImage(imagePath, new ImageDimensions(120, 80)) };
    var pairs = new List<(string New, string Old)>();
    var names = new List<string>();
    var clock = Stopwatch.StartNew();
    for (var index = 0; index < 4; index++)
    {
      var label = $"1-{index + 1}";
      var placed = service.PlaceImages(identity, "OtherTarget", EvidenceSide.New, images, requestedCaseLabel: label);
      Assert.IsTrue(placed.Succeeded, placed.Message);
      names.Add(placed.PlacedImages[0].ShapeName);
      // A cell-aligned plan must also handle a picture moved within its cell.
      object? collection = null, shape = null;
      try
      {
        collection = GetRequiredProperty(sheet, "Shapes");
        shape = InvokeMethod(collection, "Item", names[^1])!;
        SetProperty(shape, "Top", placed.PlacedImages[0].Target.TopPoints + 3.25);
      }
      finally { Release(shape); Release(collection); }
      if (index < 3)
        Assert.IsTrue(navigator.Navigate(identity, "OtherTarget", CaseNavigationDirection.Next,
          label, EvidenceSide.New, sameCaseThenNext: false).Succeeded);
    }
    for (var index = 0; index < 4; index++)
    {
      var label = $"1-{index + 1}";
      var analysis = service.Analyze(identity, "OtherTarget", EvidenceSide.Old, images, requestedCaseLabel: label);
      Assert.IsTrue(analysis.Succeeded, analysis.Message);
      if (index == 2)
      {
        var managed = new ExcelManagedShapeService();
        var before = managed.Inspect(identity, "OtherTarget", names[index]).Shape!;
        var edit = managed.Resize(identity, before, before with { TopPoints = before.TopPoints + 1.25 });
        Assert.IsTrue(edit.Succeeded, edit.Message);
      }
      var placed = service.PlaceImages(identity, "OtherTarget", EvidenceSide.Old, images,
        preparedAnalysis: analysis, requestedCaseLabel: label);
      Assert.IsTrue(placed.Succeeded, placed.Message);
      pairs.Add((names[index], placed.PlacedImages[0].ShapeName));
      if (placed.Analysis!.CompletesCaseAfterPlacement)
      {
        var cleanup = new ExcelCaseMaintenanceService().TrimCompletedCaseTail(identity,
          "OtherTarget", placed.PlacedImages[0].FocusCell.Row);
        Assert.IsTrue(cleanup.Succeeded, cleanup.Message);
        cleanup.DeletionSnapshot?.Dispose();
      }
      var snapshot = snapshots.Capture(identity, "OtherTarget").Snapshot!;
      var scoped = snapshots.Capture(identity, "OtherTarget", scopeCaseLabel: label, scopeShapes: true).Snapshot!;
      Assert.IsNotNull(scoped.ShapeScope);
      var scope = scoped.ShapeScope.Value;
      CollectionAssert.AreEqual(snapshot.Shapes.Where(shape => shape.EndRow >= scope.FirstRow &&
        shape.StartRow <= scope.LastRow).ToArray(), scoped.Shapes.ToArray());
      foreach (var pair in pairs)
      {
        var reference = snapshot.Shapes.Single(shape => shape.Name == pair.New);
        var added = snapshot.Shapes.Single(shape => shape.Name == pair.Old);
        Assert.AreEqual(reference.TopPoints, added.TopPoints, 0.05, $"Pair {pair.New}, after {label}");
      }
      if (index < 3)
        Assert.IsTrue(navigator.Navigate(identity, "OtherTarget", CaseNavigationDirection.Next,
          label, EvidenceSide.Old, sameCaseThenNext: false).Succeeded);
      if (index == 3)
      {
        var guard = placed.ReferenceResize;
        Assert.IsNotNull(guard, "A no-op reference must still protect Undo/Redo from external edits.");
        var managed = new ExcelManagedShapeService();
        Assert.IsTrue(managed.Delete(identity, placed.PlacedImages[0].Target).Succeeded);
        Assert.IsTrue(guard.SetApplied(identity, false).Succeeded);
        var moved = managed.Resize(identity, guard.Before, guard.Before with { TopPoints = guard.Before.TopPoints + 1 });
        Assert.IsTrue(moved.Succeeded, moved.Message);
        Assert.IsFalse(guard.SetApplied(identity, true).Succeeded, "Redo must reject a moved reference.");
        Assert.IsTrue(managed.Resize(identity, moved.After!, guard.Before).Succeeded);
        Assert.IsTrue(guard.SetApplied(identity, true).Succeeded);
        var original = placed.PlacedImages[0];
        var redone = new ExcelImagePlacementService().PlaceImage(identity, "OtherTarget", original.FocusCell,
          EvidenceSide.Old, imagePath, images[0].Dimensions, original.AvailableWidthPoints,
          scaleOverride: original.Plan.Image.Scale, verticalOffsetPoints: original.Plan.VerticalOffsetPoints);
        Assert.IsTrue(redone.Succeeded, redone.Message);
        Assert.AreEqual(guard.Before.TopPoints, redone.Target!.TopPoints, 0.05);
      }
    }
    Console.WriteLine($"PERF same-side backfill: {clock.Elapsed.TotalMilliseconds:F2} ms");
    var named = snapshots.Capture(identity, "OtherTarget", scopeCaseLabel: "1-4", scopeShapes: true,
      referenceShapeName: names[0]);
    Assert.IsTrue(named.Succeeded, named.Message);
    Assert.IsTrue(named.Snapshot!.Shapes.Any(shape => shape.Name == names[0]), "Named references outside the CASE must be retained.");
    SetCellValue(sheet, 220, 200, "Keep content outside the evidence columns");
    SetCellValue(sheet, 222, 202, 0);
    SetCellValue(sheet, 223, 203, false);
    var formulaCell = GetRequiredProperty(sheet, "Cells", 221, 201);
    try { SetProperty(formulaCell, "Formula", "=\"\""); }
    finally { Release(formulaCell); }
    var safety = new ExcelRowMutationService().CaptureRowSafetyStates(identity, "OtherTarget", 219, 224);
    Assert.IsTrue(safety.Succeeded, safety.Message);
    foreach (var row in safety.Rows)
      Assert.AreEqual(row.Row is >= 220 and <= 223, row.HasValueOrFormula,
        "Safety reads must preserve formulas returning empty strings, zero, false and cells beyond the evidence columns.");
  }
}
