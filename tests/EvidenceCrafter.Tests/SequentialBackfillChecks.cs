using System.Drawing.Imaging;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  [TestMethod]
  [DataRow(120, 60)]
  [DataRow(120, 180)]
  [DataRow(180, 180)]
  public void PairedImages_BackfillSixImagesWithoutDrift(int oldWidth, int oldHeight) =>
    RunSupervisedScenario(Scenario.SequentialBackfill,
      (sheet, identity, path) => VerifySequentialBackfill(sheet, identity, path, oldWidth, oldHeight));

  [TestMethod]
  public void PairedImages_BackfillAfterRowInsertionsWithNegativeOffsets() =>
    RunSupervisedScenario(Scenario.SequentialBackfill, VerifyNegativeOffsetBackfill);

  private static void VerifyNegativeOffsetBackfill(object sheet, WorkbookIdentity identity, string path)
  {
    SetRangeProperty(sheet, "1:1048576", "RowHeight", 409.5);
    SetCellValue(sheet, 199997, 3, "NEW");
    SetCellValue(sheet, 199997, 18, "OLD");
    SetCellValue(sheet, 199998, 1, 1);
    SetCellValue(sheet, 199998, 2, 1);
    SetCellValue(sheet, 200098, 1, 1);
    SetCellValue(sheet, 200098, 2, 2);
    SetRangeBorder(sheet, 199998, 17, 200148, 17, 10);
    SetRangeBorder(sheet, 200148, 1, 200148, 32, 9);
    using (var bitmap = new Bitmap(120, 60)) bitmap.Save(path, ImageFormat.Png);
    var collection = GetRequiredProperty(sheet, "Shapes");
    var names = new List<string>();
    try
    {
      for (var index = 0; index < 6; index++)
      {
        var row = 200000 + index * 4;
        var cell = GetRequiredProperty(sheet, "Cells", row, 4);
        object? shape = null;
        try
        {
          shape = InvokeMethod(collection, "AddPicture", path, 0, -1,
            Convert.ToDouble(GetRequiredProperty(cell, "Left")) + 6,
            Convert.ToDouble(GetRequiredProperty(cell, "Top")) + 2, 120, 60)!;
          var name = $"EST_IMG_{Guid.NewGuid():N}";
          names.Add(name);
          SetProperty(shape, "Name", name);
          SetProperty(shape, "AlternativeText", new ManagedShapeMetadata(1, EvidenceSide.New, new(row, 4))
            { SourceDimensions = new(120, 60), AppliedScale = 1 }.Serialize());
          SetProperty(shape, "Placement", 2);
        }
        finally { Release(shape); Release(cell); }
      }
      for (var pass = 0; pass < 4; pass++)
      {
        var rows = GetRequiredProperty(sheet, "Range", "199980:199984");
        try { InvokeMethod(rows, "Insert"); SetProperty(rows, "RowHeight", 15); }
        finally { Release(rows); }
      }
    }
    finally { Release(collection); }
    var snapshots = new ExcelSheetSnapshotService();
    var captured = snapshots.Capture(identity, "OtherTarget", scopeCaseLabel: "1-1");
    Assert.IsTrue(captured.Succeeded, captured.Message);
    var before = captured.Snapshot!;
    Assert.IsTrue(before.Shapes.Any(shape => shape.VerticalOffsetPoints < -1),
      "Native Excel row insertions must reproduce the reported negative offset.");
    Console.WriteLine("Native offsets: " + string.Join(", ", before.Shapes.Select(shape => shape.VerticalOffsetPoints)));
    var service = new ExcelAutomaticPlacementService();
    var images = new[] { new AutomaticPlacementImage(path, new(120, 60)) };
    foreach (var name in names)
    {
      var placed = service.PlaceImages(identity, "OtherTarget", EvidenceSide.Old, images, requestedCaseLabel: "1-1");
      Assert.IsTrue(placed.Succeeded, placed.Message);
      Assert.AreEqual(name, placed.Analysis!.Steps.Single().Pair!.ShapeName);
      var reference = before.Shapes.Single(shape => shape.Name == name);
      Assert.AreEqual(reference.TopPoints, placed.PlacedImages.Single().Target.TopPoints, 1);
      Assert.IsEmpty(placed.AppliedInsertions, "Matching images already have enough space.");
      if (name == names[0])
      {
        var original = placed.PlacedImages.Single();
        var target = original.Target;
        for (var replay = 0; replay < 3; replay++)
        {
          var deletion = new ExcelManagedShapeService().Delete(identity, target);
          Assert.IsTrue(deletion.Succeeded, deletion.Message);
          Assert.IsTrue(placed.ReferenceResize!.SetApplied(identity, false).Succeeded);
          Assert.IsTrue(placed.ReferenceResize.SetApplied(identity, true).Succeeded);
          var redone = new ExcelImagePlacementService().PlaceImage(identity, "OtherTarget", original.FocusCell,
            EvidenceSide.Old, path, images[0].Dimensions, original.AvailableWidthPoints,
            scaleOverride: original.Plan.Image.Scale, verticalOffsetPoints: original.Plan.VerticalOffsetPoints,
            topPointsOverride: original.Target.TopPoints);
          Assert.IsTrue(redone.Succeeded, redone.Message);
          target = redone.Target!;
          Assert.AreEqual(reference.TopPoints, target.TopPoints, 1, "Repeated replay must not drift.");
        }
      }
    }
    var after = snapshots.Capture(identity, "OtherTarget", scopeCaseLabel: "1-1").Snapshot!;
    foreach (var original in before.Shapes)
      Assert.AreEqual(original.TopPoints, after.Shapes.Single(shape => shape.Name == original.Name).TopPoints, 1,
        "Backfill must not accumulate downward movement of the references.");
  }

  private static void VerifySequentialBackfill(object sheet, WorkbookIdentity identity, string path, int oldWidth, int oldHeight)
  {
    SetCellValue(sheet, 2, 3, "NEW");
    SetCellValue(sheet, 2, 18, "OLD");
    SetCellValue(sheet, 3, 1, 1);
    SetCellValue(sheet, 3, 2, 1);
    SetCellValue(sheet, 103, 1, 1);
    SetCellValue(sheet, 103, 2, 2);
    SetRangeProperty(sheet, "A1:AF150", "RowHeight", 13.5);
    SetRangeBorder(sheet, 3, 17, 150, 17, 10);
    SetRangeBorder(sheet, 150, 1, 150, 32, 9);
    var service = new ExcelAutomaticPlacementService();
    var snapshots = new ExcelSheetSnapshotService();
    var references = new List<string>();
    var counterparts = new List<string>();
    for (var pass = 0; pass < 12; pass++)
    {
      var side = pass < 6 ? EvidenceSide.New : EvidenceSide.Old;
      var height = pass < 6 ? 60 : oldHeight;
      var width = pass < 6 ? 120 : oldWidth;
      using (var bitmap = new Bitmap(width, height)) bitmap.Save(path, ImageFormat.Png);
      var before = snapshots.Capture(identity, "OtherTarget", 3).Snapshot!;
      var placed = service.PlaceImages(identity, "OtherTarget", side,
        [new AutomaticPlacementImage(path, new ImageDimensions(width, height))], requestedCaseLabel: "1-1");
      Assert.IsTrue(placed.Succeeded, $"Pass {pass}: {placed.Message}");
      if (side == EvidenceSide.New)
      {
        references.Add(placed.PlacedImages.Single().ShapeName);
        continue;
      }
      counterparts.Add(placed.PlacedImages.Single().ShapeName);
      if (placed.Analysis!.CompletesCaseAfterPlacement)
      {
        var cleanup = new ExcelCaseMaintenanceService().TrimCompletedCaseTail(identity,
          "OtherTarget", placed.PlacedImages.Single().FocusCell.Row);
        Assert.IsTrue(cleanup.Succeeded, cleanup.Message);
        cleanup.DeletionSnapshot?.Dispose();
      }
      var after = snapshots.Capture(identity, "OtherTarget", 3).Snapshot!;
      for (var pair = 0; pair < counterparts.Count; pair++)
      {
        var reference = after.Shapes.Single(shape => shape.Name == references[pair]);
        var counterpart = after.Shapes.Single(shape => shape.Name == counterparts[pair]);
        Assert.AreEqual(reference.TopPoints, counterpart.TopPoints, 1, $"Pass {pass}, pair {pair}");
      }
      if (oldWidth == 120 && oldHeight == 60)
      {
        Assert.IsTrue(placed.AppliedInsertions.All(insertion => insertion.StartRow > before.Shapes.Max(shape => shape.EndRow)),
          "Only tail padding removed by cleanup may need to be restored.");
        foreach (var original in before.Shapes)
          Assert.AreEqual(original.TopPoints, after.Shapes.Single(shape => shape.Name == original.Name).TopPoints, 1,
            "Equal-sized backfill must not move existing images down.");
      }
    }
  }
}
