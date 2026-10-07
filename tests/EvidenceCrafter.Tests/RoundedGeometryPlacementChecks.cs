using System.Globalization;
using System.Drawing.Imaging;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  private static void VerifyRoundedPlacementGeometry(object sheet, WorkbookIdentity identity, string path)
  {
    var service = new ExcelImagePlacementService();
    var managed = new ExcelManagedShapeService();
    var shapes = GetRequiredProperty(sheet, "Shapes");
    double Read(object shape, string property) => Convert.ToDouble(GetRequiredProperty(shape, property), CultureInfo.InvariantCulture);
    int Count() => Convert.ToInt32(GetRequiredProperty(shapes, "Count"), CultureInfo.InvariantCulture);
    void Change(object shape, string property, double value)
    {
      SetProperty(shape, "LockAspectRatio", 0);
      SetProperty(shape, property, value);
      SetProperty(shape, "LockAspectRatio", -1);
    }
    try
    {
      SetCellValue(sheet, 1, 1, "keep existing content");
      var sentinel = InvokeMethod(shapes, "AddShape", 1, 1800f, 0f, 30f, 30f)!;
      try { SetProperty(sentinel, "Name", "ExistingShape"); }
      finally { Release(sentinel); }
      using (var bitmap = new Bitmap(994, 994)) bitmap.Save(path, ImageFormat.Png);
      foreach (var row in new[] { 6, 100000 })
      foreach (var property in new[] { "Left", "Top", "Width", "Height" })
      foreach (var delta in new[] { -1d, 1d, -0.3, 1.04, -1.25, 1.25 })
      {
        var cell = GetRequiredProperty(sheet, "Cells", row, 4);
        double planned;
        try { planned = property == "Left" ? Read(cell, "Left") + 6 : property == "Top" ? Read(cell, "Top") + 2 : 745.5; }
        finally { Release(cell); }
        var observed = double.NaN;
        ExcelImagePlacementService.PlacementStageObserved = (stage, shape) =>
        {
          if (stage != "AfterAttributes") return;
          Change(shape, property, planned + delta);
          observed = Read(shape, property);
        };
        var result = service.PlaceImage(identity, "OtherTarget", new(row, 4), EvidenceSide.New, path, new(745.5, 745.5), 745.5);
        // Excel may quantize the requested position again; compare its independent raw read.
        var accepted = Math.Abs((decimal)Math.Round(observed, 1, MidpointRounding.AwayFromZero) -
          (decimal)Math.Round(planned, 1, MidpointRounding.AwayFromZero)) <= 1m;
        Assert.AreEqual(accepted, result.Succeeded, $"{row}/{property}/{delta}, planned={planned:R}, observed={observed:R}: {result.Message}");
        if (accepted)
        {
          var raw = InvokeMethod(shapes, "Item", result.ShapeName)!;
          try
          {
            Assert.AreEqual(Read(raw, "Left"), result.Target!.LeftPoints);
            Assert.AreEqual(Read(raw, "Top"), result.Target.TopPoints);
            Assert.AreEqual(Read(raw, "Width"), result.Target.WidthPoints);
            Assert.AreEqual(Read(raw, "Height"), result.Target.HeightPoints);
            Assert.IsTrue(PlacementGeometryComparison.Matches(planned, Read(raw, property)));
            if (row == 6 && property == "Width" && delta == -0.3)
            {
              Change(raw, "Width", result.Target.WidthPoints + 0.3);
              Assert.IsFalse(managed.Delete(identity, result.Target).Succeeded, "External-edit detection remains independent from placement tolerance.");
              Change(raw, "Width", result.Target.WidthPoints);
            }
          }
          finally { Release(raw); }
          Assert.IsTrue(managed.Delete(identity, result.Target!).Succeeded);
        }
        else Assert.IsFalse(result.MutationMayHaveOccurred);
        Assert.AreEqual(1, Count());
      }

      ExcelImagePlacementService.PlacementStageObserved = null;
      var offsetCell = GetRequiredProperty(sheet, "Cells", 6, 4);
      double cellTop;
      try { cellTop = Read(offsetCell, "Top"); }
      finally { Release(offsetCell); }
      foreach (var offset in new[] { -1d, -0.3 })
      {
        var normalized = service.PlaceImage(identity, "OtherTarget", new(6, 4), EvidenceSide.New,
          path, new(745.5, 745.5), 745.5, verticalOffsetPoints: offset);
        Assert.IsTrue(normalized.Succeeded, normalized.Message);
        Assert.AreEqual(cellTop, normalized.Target!.TopPoints, 0.05);
        Assert.IsTrue(managed.Delete(identity, normalized.Target).Succeeded);
        Assert.AreEqual(1, Count());
      }
      Assert.ThrowsExactly<ArgumentOutOfRangeException>(() => service.PlaceImage(identity, "OtherTarget",
        new(6, 4), EvidenceSide.New, path, new(745.5, 745.5), 745.5, verticalOffsetPoints: -1.001));

      // Width and scale may each be valid, but their combined excess still exceeds the side limit.
      ExcelImagePlacementService.PlacementStageObserved = (stage, shape) =>
      { if (stage == "AfterAttributes") Change(shape, "Width", 746.5); };
      var overflow = service.PlaceImage(identity, "OtherTarget", new(6, 4), EvidenceSide.New,
        path, new(745.5, 745.5), 744.5, scaleOverride: 1);
      Assert.IsFalse(overflow.Succeeded);
      StringAssert.Contains(overflow.Message, "配置可能幅");
      Assert.AreEqual(1, Count());

      foreach (var fault in new[] { "Size", "Name", "Position", "Rollback" })
      {
        ExcelImagePlacementService.PlacementStageObserved = (stage, shape) =>
        {
          if (stage == "AfterAttributes")
          {
            if (fault == "Name") SetProperty(shape, "Name", "WrongName");
            else if (fault == "Position") Change(shape, "Left", Read(shape, "Left") + 10);
            else Change(shape, "Width", 744.25);
          }
          if (stage == "BeforeRollback" && fault == "Rollback") throw new InvalidOperationException("Injected deletion failure");
        };
        var failed = service.PlaceImage(identity, "OtherTarget", new(6, 4), EvidenceSide.New, path, new(745.5, 745.5), 745.5);
        Assert.IsFalse(failed.Succeeded, fault);
        Assert.AreEqual(1d, failed.Diagnostic!.ComparisonTolerancePoints);
        Assert.AreEqual(fault == "Rollback", failed.MutationMayHaveOccurred);
        Assert.AreEqual(fault == "Rollback" ? 2 : 1, Count());
        if (fault == "Rollback")
        {
          var retained = InvokeMethod(shapes, "Item", 2)!;
          try { _ = InvokeMethod(retained, "Delete"); }
          finally { Release(retained); }
        }
      }

      SetRangeProperty(sheet, "A1:AF202", "NumberFormat", "0.00");
      SetCellValue(sheet, 2, 3, "NEW"); SetCellValue(sheet, 2, 18, "OLD");
      SetCellValue(sheet, 3, 1, 1); SetCellValue(sheet, 3, 2, 1);
      SetCellValue(sheet, 13, 1, 1); SetCellValue(sheet, 13, 2, 2);
      using (var bitmap = new Bitmap(120, 80)) bitmap.Save(path, ImageFormat.Png);
      var automatic = new ExcelAutomaticPlacementService();
      var images = new[] { new AutomaticPlacementImage(path, new(120, 80)) };
      var insertions = 0;
      ExcelImagePlacementService.PlacementStageObserved = (stage, shape) =>
      {
        if (stage != "AfterAttributes") return;
        insertions++;
        Change(shape, "Width", Read(shape, "Width") - 0.3);
        Change(shape, "Height", Read(shape, "Height") - 0.3);
        // The OLD/Redo image ends up 1pt below the reference, which must also pass the pair check.
        Change(shape, "Top", Read(shape, "Top") + (insertions == 1 ? 0 : 1));
      };
      var first = automatic.PlaceImages(identity, "OtherTarget", EvidenceSide.New, images, requestedCaseLabel: "1-1");
      Assert.IsTrue(first.Succeeded, first.Message);
      var pair = automatic.PlaceImages(identity, "OtherTarget", EvidenceSide.Old, images, requestedCaseLabel: "1-1");
      Assert.IsTrue(pair.Succeeded, pair.Message);
      var reference = managed.Inspect(identity, "OtherTarget", first.PlacedImages[0].ShapeName).Shape!;
      var old = pair.PlacedImages[0];
      Assert.AreEqual(1, old.Target.TopPoints - reference.TopPoints, 0.001);
      Assert.IsTrue(managed.Delete(identity, old.Target).Succeeded);
      Assert.IsTrue(pair.ReferenceResize!.SetApplied(identity, false).Succeeded);
      Assert.IsTrue(pair.ReferenceResize.SetApplied(identity, true).Succeeded);
      var redone = service.PlaceImage(identity, "OtherTarget", old.FocusCell, EvidenceSide.Old,
        path, images[0].Dimensions, old.AvailableWidthPoints,
        scaleOverride: old.Plan.Image.Scale, verticalOffsetPoints: old.Plan.VerticalOffsetPoints);
      Assert.IsTrue(redone.Succeeded, redone.Message);
      Assert.AreEqual(old.Target.TopPoints, redone.Target!.TopPoints);
      Assert.AreEqual(old.Target.WidthPoints, redone.Target.WidthPoints);
      Assert.AreEqual(old.Target.HeightPoints, redone.Target.HeightPoints);
      Assert.IsTrue(managed.Delete(identity, redone.Target).Succeeded);
      Assert.IsTrue(managed.Delete(identity, managed.Inspect(identity, "OtherTarget", reference.ShapeName).Shape!).Succeeded);

      var snapshots = new ExcelSheetSnapshotService();
      var previous = snapshots.Capture(identity, "OtherTarget").Snapshot!;
      insertions = 0;
      ExcelImagePlacementService.PlacementStageObserved = (stage, shape) =>
      {
        if (stage != "AfterAttributes") return;
        Change(shape, "Height", Read(shape, "Height") - (++insertions == 2 ? 1.25 : 0.3));
      };
      var compensated = automatic.PlaceImages(identity, "OtherTarget", EvidenceSide.New, [images[0], images[0]], requestedCaseLabel: "1-1");
      Assert.IsFalse(compensated.Succeeded);
      Assert.IsTrue(compensated.CompensationSucceeded, compensated.Message);
      var after = snapshots.Capture(identity, "OtherTarget").Snapshot!;
      CollectionAssert.AreEqual(previous.LayoutSignals.Anchors.ToArray(), after.LayoutSignals.Anchors.ToArray());
      Assert.AreEqual(1, Count());
      var original = GetRequiredProperty(sheet, "Cells", 1, 1);
      try { Assert.AreEqual("keep existing content", GetRequiredProperty(original, "Value2")); }
      finally { Release(original); }
    }
    finally { ExcelImagePlacementService.PlacementStageObserved = null; Release(shapes); }
  }
}
