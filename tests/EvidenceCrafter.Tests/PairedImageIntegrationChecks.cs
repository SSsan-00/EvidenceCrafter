using System.Drawing.Imaging;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
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
