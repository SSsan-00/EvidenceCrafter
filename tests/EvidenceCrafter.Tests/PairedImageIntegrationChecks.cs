using System.Drawing.Imaging;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  private static void VerifyPairedCollision(object sheet, WorkbookIdentity identity, string imagePath)
  {
    SetCellValue(sheet, 2, 3, "NEW");
    SetCellValue(sheet, 2, 18, "OLD");
    SetCellValue(sheet, 3, 1, 1);
    SetCellValue(sheet, 3, 2, 1);
    SetCellValue(sheet, 53, 1, 1);
    SetCellValue(sheet, 53, 2, 2);
    SetRangeBorder(sheet, 3, 17, 102, 17, 10);
    SetRangeBorder(sheet, 102, 1, 102, 32, 9);
    using (var bitmap = new Bitmap(120, 60)) bitmap.Save(imagePath, ImageFormat.Png);
    var service = new ExcelAutomaticPlacementService();
    var images = new[] { new AutomaticPlacementImage(imagePath, new ImageDimensions(120, 60)) };
    var first = service.PlaceImages(identity, "OtherTarget", EvidenceSide.New, images, requestedCaseLabel: "1-1");
    Assert.IsTrue(first.Succeeded, first.Message);
    SetCellValue(sheet, 40, 19, "Preserve this cell");
    var placed = service.PlaceImages(identity, "OtherTarget", EvidenceSide.Old,
      [new AutomaticPlacementImage(imagePath, new ImageDimensions(120, 600))], requestedCaseLabel: "1-1");
    Assert.IsTrue(placed.Succeeded, placed.Message);
    Assert.AreEqual(5, placed.PlacedImages[0].FocusCell.Row);
    var shiftedCellRow = 40 + (placed.ReferenceResize?.Insertion?.Count ?? 0);
    Assert.IsTrue(placed.AppliedInsertions.Any(row => row.StartRow == shiftedCellRow));
    var snapshot = new ExcelSheetSnapshotService().Capture(identity, "OtherTarget", 3).Snapshot!;
    var picture = snapshot.Shapes.Single(shape => shape.Name == placed.PlacedImages[0].ShapeName);
    Assert.IsTrue(snapshot.Cells.Any(cell => cell.Column == 19 && cell.Row > picture.EndRow && cell.HasValueOrFormula));
    Assert.IsTrue(new ExcelManagedShapeService().Delete(identity, placed.PlacedImages[0].Target).Succeeded);
    foreach (var row in placed.AppliedInsertions.Reverse())
    {
      var undo = new ExcelRowMutationService().DeleteRowsIfSafe(identity, row.WorksheetName, row.StartRow, row.Count);
      Assert.IsTrue(undo.Succeeded && undo.Changed, undo.Message);
    }
    if (placed.ReferenceResize is { } reference) Assert.IsTrue(reference.SetApplied(identity, false).Succeeded);
    var restored = new ExcelSheetSnapshotService().Capture(identity, "OtherTarget", 3).Snapshot!;
    Assert.IsTrue(restored.Cells.Any(cell => cell.Row == 40 && cell.Column == 19 && cell.HasValueOrFormula));
  }

  private static void VerifyPairedImageAlignment(object sheet, WorkbookIdentity identity, string imagePath)
  {
    SetCellValue(sheet, 2, 3, "NEW");
    SetCellValue(sheet, 2, 18, "OLD");
    SetCellValue(sheet, 3, 1, 1);
    SetCellValue(sheet, 3, 2, 1);
    SetCellValue(sheet, 53, 1, 1);
    SetCellValue(sheet, 53, 2, 2);
    SetRangeBorder(sheet, 3, 17, 102, 17, 10);
    SetRangeBorder(sheet, 102, 1, 102, 32, 9);
    var service = new ExcelAutomaticPlacementService();
    var shapes = new ExcelManagedShapeService();
    var newNames = new List<string>();
    for (var pass = 0; pass < 4; pass++)
    {
      var side = pass < 2 ? EvidenceSide.New : EvidenceSide.Old;
      var height = side == EvidenceSide.New ? 80 : 240;
      using (var bitmap = new Bitmap(120, height))
      {
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.CornflowerBlue);
        bitmap.Save(imagePath, ImageFormat.Png);
      }
      var beforePlacement = pass == 2
        ? new ExcelSheetSnapshotService().Capture(identity, "OtherTarget", 3).Snapshot! : null;
      var placed = service.PlaceImages(identity, "OtherTarget", side,
        [new AutomaticPlacementImage(imagePath, new ImageDimensions(120, height))], requestedCaseLabel: "1-1");
      Assert.IsTrue(placed.Succeeded, placed.Message);
      if (side == EvidenceSide.New)
      {
        newNames.Add(placed.PlacedImages[0].ShapeName);
        if (pass == 0)
        {
          var original = shapes.Inspect(identity, "OtherTarget", placed.PlacedImages[0].ShapeName).Shape!;
          var staleResize = new PairedImageResize(original, original, null);
          var edited = shapes.Resize(identity, original, original with { WidthPoints = original.WidthPoints * 0.9 });
          Assert.IsTrue(edited.Succeeded, edited.Message);
          Assert.IsFalse(staleResize.SetApplied(identity, true).Succeeded);
          Assert.IsTrue(staleResize.CanRetryPreparation, "Initial placement may re-analyze a changed reference.");
          Assert.IsTrue(staleResize.CompensationSucceeded, "The failed precheck must not mutate Excel.");
          Assert.IsFalse(staleResize.SetApplied(identity, false).Succeeded);
          Assert.IsFalse(staleResize.CanRetryPreparation, "Undo must keep its strict expected-state check.");
          Assert.IsTrue(shapes.Resize(identity, edited.After!, original).Succeeded);
          object? collection = null, obstacle = null;
          try
          {
            collection = GetRequiredProperty(sheet, "Shapes");
            obstacle = InvokeMethod(collection, "AddShape", 1,
              original.LeftPoints + original.WidthPoints + 10, original.TopPoints, 8, 8)!;
            var recovered = service.PlaceImages(identity, "OtherTarget", EvidenceSide.Old,
              [new AutomaticPlacementImage(imagePath, new ImageDimensions(120, height))], requestedCaseLabel: "1-1");
            Assert.IsTrue(recovered.Succeeded, recovered.Message);
            Assert.IsNull(recovered.ReferenceResize, "The obstacle must prevent reference enlargement and trigger recovery.");
            var kept = shapes.Inspect(identity, "OtherTarget", original.ShapeName).Shape!;
            Assert.AreEqual(original.WidthPoints, kept.WidthPoints, 0.05);
            Assert.AreEqual(original.HeightPoints, kept.HeightPoints, 0.05);
            Assert.AreEqual(original.TopPoints, recovered.PlacedImages[0].Target.TopPoints, 0.05);
            Assert.IsTrue(shapes.Delete(identity, recovered.PlacedImages[0].Target).Succeeded);
            foreach (var inserted in recovered.AppliedInsertions.Reverse())
            {
              var deleted = new ExcelRowMutationService().DeleteRowsIfSafe(identity,
                inserted.WorksheetName, inserted.StartRow, inserted.Count);
              Assert.IsTrue(deleted.Succeeded && deleted.Changed, deleted.Message);
            }
          }
          finally
          {
            if (obstacle is not null) InvokeMethod(obstacle, "Delete");
            Release(obstacle);
            Release(collection);
          }
        }
        continue;
      }
      var snapshot = new ExcelSheetSnapshotService().Capture(identity, "OtherTarget", 3).Snapshot!;
      var reference = snapshot.Shapes.Single(shape => shape.Name == newNames[pass - 2]);
      var added = snapshot.Shapes.Single(shape => shape.Name == placed.PlacedImages[0].ShapeName);
      Assert.AreEqual(reference.StartRow, added.StartRow, "Corresponding pictures must start on the same actual Excel row.");
      Assert.AreEqual(reference.TopPoints, added.TopPoints, 0.05);
      if (pass == 2)
      {
        var nextImage = snapshot.Shapes.Single(shape => shape.Name == newNames[1]);
        Assert.IsGreaterThan(Math.Max(reference.EndRow, added.EndRow) + 2, nextImage.StartRow,
          "The second band must move below both images in the first band.");
      }
      Assert.IsNotNull(placed.ReferenceResize);
      Assert.IsNotNull(placed.ReferenceResize.Insertion, "The small initial image must grow when its pair is added.");
      if (pass == 2)
      {
        Assert.IsTrue(shapes.Delete(identity, placed.PlacedImages[0].Target).Succeeded);
        foreach (var insertion in placed.AppliedInsertions.Reverse())
        {
          var removed = new ExcelRowMutationService().DeleteRowsIfSafe(identity,
            insertion.WorksheetName, insertion.StartRow, insertion.Count);
          Assert.IsTrue(removed.Succeeded && removed.Changed, removed.Message);
        }
        var undone = placed.ReferenceResize.SetApplied(identity, false);
        Assert.IsTrue(undone.Succeeded, undone.Message);
        var restored = new ExcelSheetSnapshotService().Capture(identity, "OtherTarget", 3).Snapshot!;
        foreach (var original in beforePlacement!.Shapes)
        {
          var actual = restored.Shapes.Single(item => item.Name == original.Name);
          Assert.AreEqual(original.StartRow, actual.StartRow);
          Assert.AreEqual(original.TopPoints, actual.TopPoints, 0.05);
          Assert.AreEqual(original.HeightPoints, actual.HeightPoints, 0.05);
        }
        var repeated = service.PlaceImages(identity, "OtherTarget", side,
          [new AutomaticPlacementImage(imagePath, new ImageDimensions(120, height))], requestedCaseLabel: "1-1");
        Assert.IsTrue(repeated.Succeeded, repeated.Message);
      }
      if (pass == 3)
      {
        var deleted = shapes.Delete(identity, placed.PlacedImages[0].Target);
        Assert.IsTrue(deleted.Succeeded, deleted.Message);
        foreach (var insertion in placed.AppliedInsertions.Reverse())
        {
          var removed = new ExcelRowMutationService().DeleteRowsIfSafe(identity,
            insertion.WorksheetName, insertion.StartRow, insertion.Count);
          Assert.IsTrue(removed.Succeeded, removed.Message);
        }
        var reservedRow = placed.ReferenceResize.Insertion!.StartRow;
        SetCellValue(sheet, reservedRow, 1, "Keep this user entry");
        var blockedUndo = placed.ReferenceResize.SetApplied(identity, false);
        Assert.IsFalse(blockedUndo.Succeeded, "Skipped row deletion must not be reported as successful Undo.");
        Assert.IsTrue(placed.ReferenceResize.Matches(identity, true), "Blocked Undo must restore the applied image size.");
        SetCellValue(sheet, reservedRow, 1, null!);
        var undo = placed.ReferenceResize.SetApplied(identity, false);
        Assert.IsTrue(undo.Succeeded, undo.Message);
        Assert.IsTrue(placed.ReferenceResize.Matches(identity, false));
        var redo = placed.ReferenceResize.SetApplied(identity, true);
        Assert.IsTrue(redo.Succeeded, redo.Message);
        Assert.IsTrue(placed.ReferenceResize.Matches(identity, true));
      }
    }
  }
}
