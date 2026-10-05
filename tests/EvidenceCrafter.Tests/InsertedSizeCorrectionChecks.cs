using System.Globalization;
using System.Drawing.Imaging;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  private static void VerifyInsertedSizeCorrection(object sheet, WorkbookIdentity identity, string path)
  {
    const double reportedActual = 745.2000122070312;
    var service = new ExcelImagePlacementService();
    var shapes = GetRequiredProperty(sheet, "Shapes");
    double Read(object shape, string property) => Convert.ToDouble(GetRequiredProperty(shape, property), CultureInfo.InvariantCulture);
    int Count() => Convert.ToInt32(GetRequiredProperty(shapes, "Count"), CultureInfo.InvariantCulture);
    void SetDrift(object shape, string property)
    {
      SetProperty(shape, "LockAspectRatio", 0);
      SetProperty(shape, property, reportedActual);
      SetProperty(shape, "LockAspectRatio", -1);
      Assert.AreEqual(reportedActual, Read(shape, property));
    }
    try
    {
      SetCellValue(sheet, 1, 1, "keep existing content");
      var sentinel = InvokeMethod(shapes, "AddShape", 1, 1800f, 0f, 30f, 30f)!;
      try { SetProperty(sentinel, "Name", "ExistingShape"); }
      finally { Release(sentinel); }
      foreach (var property in new[] { "Width", "Height" })
      foreach (var factor in new[] { 1, 2 })
      {
        var pixels = property == "Width" ? (994 * factor, 500 * factor) : (500 * factor, 994 * factor);
        using (var bitmap = new Bitmap(pixels.Item1, pixels.Item2)) bitmap.Save(path, ImageFormat.Png);
        var dimensions = new ImageDimensions(pixels.Item1 * 0.75, pixels.Item2 * 0.75);
        var expected = new ImagePlacementGeometry(0, 0, dimensions.WidthPoints / factor, dimensions.HeightPoints / factor);
        foreach (var row in new[] { 1, 100000 })
        foreach (var injectionStage in new[] { "AfterInsert", "AfterAttributes" })
        {
          var corrections = 0;
          ExcelImagePlacementService.PlacementStageObserved = (stage, shape) =>
          {
            if (stage == injectionStage) SetDrift(shape, property);
            if (stage == "AfterSizeCorrection") corrections++;
          };
          var placed = service.PlaceImage(identity, "OtherTarget", new(row, 4), EvidenceSide.New,
            path, dimensions, expected.Width, verticalOffsetPoints: 2);
          Assert.IsTrue(placed.Succeeded, placed.Message);
          Assert.AreEqual(1, corrections);
          var raw = InvokeMethod(shapes, "Item", placed.ShapeName)!;
          try
          {
            Assert.AreEqual(expected.Width, Read(raw, "Width"), 0.05);
            Assert.AreEqual(expected.Height, Read(raw, "Height"), 0.05);
            Assert.AreEqual(-1d, Read(raw, "LockAspectRatio"));
            Assert.AreEqual(placed.Target!.WidthPoints, Read(raw, "Width"));
            Assert.AreEqual(placed.Target.HeightPoints, Read(raw, "Height"));
            Assert.AreEqual(placed.Target.TopPoints, Read(raw, "Top"));
          }
          finally { Release(raw); }
          Assert.IsTrue(service.DeletePlacedImage(identity, "OtherTarget", placed.ShapeName).Succeeded);
          Assert.AreEqual(1, Count());
        }
      }

      // Persisting size drift is accepted through 1pt, with final Excel geometry retained.
      using (var bitmap = new Bitmap(994, 994)) bitmap.Save(path, ImageFormat.Png);
      foreach (var property in new[] { "Width", "Height" })
      foreach (var (delta, accepted) in new[] { (-1d, true), (1d, true), (-1.25, false), (1.25, false), (-0.3, true) })
      {
        var corrections = 0;
        void ApplyDrift(object shape)
        {
          SetProperty(shape, "LockAspectRatio", 0);
          SetProperty(shape, property, 745.5 + delta);
          SetProperty(shape, "LockAspectRatio", -1);
        }
        ExcelImagePlacementService.PlacementStageObserved = (stage, shape) =>
        {
          if (stage == "AfterAttributes") ApplyDrift(shape);
          if (stage == "AfterSizeCorrection") { corrections++; ApplyDrift(shape); }
        };
        var result = service.PlaceImage(identity, "OtherTarget", new(6, 4), EvidenceSide.New,
          path, new(745.5, 745.5), 750);
        Assert.AreEqual(accepted, result.Succeeded, result.Message);
        Assert.AreEqual(delta == -0.3 ? 1 : 0, corrections);
        if (accepted)
        {
          var raw = InvokeMethod(shapes, "Item", result.ShapeName)!;
          try
          {
            Assert.AreEqual(745.5 + delta, Read(raw, property), 0.001);
            Assert.AreEqual(Read(raw, "Width"), result.Target!.WidthPoints);
            Assert.AreEqual(Read(raw, "Height"), result.Target.HeightPoints);
          }
          finally { Release(raw); }
          var managedImage = new ExcelManagedShapeService();
          if (delta == -0.3)
          {
            var edited = InvokeMethod(shapes, "Item", result.ShapeName)!;
            try
            {
              SetProperty(edited, "LockAspectRatio", 0);
              SetProperty(edited, "Width", result.Target!.WidthPoints + 0.3);
              Assert.IsFalse(managedImage.Delete(identity, result.Target).Succeeded,
                "A tolerated insertion must not relax external-edit detection during Undo.");
              SetProperty(edited, "Width", result.Target.WidthPoints);
              SetProperty(edited, "LockAspectRatio", -1);
            }
            finally { Release(edited); }
          }
          Assert.IsTrue(managedImage.Delete(identity, result.Target!).Succeeded);
        }
        else Assert.IsFalse(result.MutationMayHaveOccurred);
        Assert.AreEqual(1, Count());
      }

      // An otherwise permitted width change must not exceed the available side width.
      ExcelImagePlacementService.PlacementStageObserved = (stage, shape) =>
      {
        if (stage == "AfterAttributes")
        { SetProperty(shape, "LockAspectRatio", 0); SetProperty(shape, "Width", 746.5); }
      };
      var overflow = service.PlaceImage(identity, "OtherTarget", new(6, 4), EvidenceSide.New,
        path, new(745.5, 745.5), 745.5);
      Assert.IsFalse(overflow.Succeeded);
      StringAssert.Contains(overflow.Message, "配置可能幅");
      Assert.AreEqual(1, Count());

      // Greater drift and correction exceptions still fail without another retry.
      using (var bitmap = new Bitmap(994, 500)) bitmap.Save(path, ImageFormat.Png);
      foreach (var fault in new[] { "PersistentDrift", "CorrectionThrows", "Name", "Position", "Rollback" })
      {
        var corrections = 0;
        ExcelImagePlacementService.PlacementStageObserved = (stage, shape) =>
        {
          if (stage == "AfterAttributes") SetDrift(shape, "Width");
          if (stage == "AfterSizeCorrection")
          {
            corrections++;
            if (fault == "CorrectionThrows") throw new InvalidOperationException("Injected correction failure");
            if (fault == "Name") SetProperty(shape, "Name", "WrongName");
            else if (fault == "Position") SetProperty(shape, "Left", Read(shape, "Left") + 10);
            else
            { SetProperty(shape, "LockAspectRatio", 0); SetProperty(shape, "Width", 744.25); }
          }
          if (stage == "BeforeRollback" && fault == "Rollback") throw new InvalidOperationException("Injected deletion failure");
        };
        var failed = service.PlaceImage(identity, "OtherTarget", new(6, 4), EvidenceSide.New,
          path, new(745.5, 375), 745.5);
        Assert.IsFalse(failed.Succeeded, fault);
        Assert.AreEqual(1, corrections);
        Assert.IsTrue(failed.Diagnostic!.SizeCorrectionAttempted);
        Assert.AreEqual(reportedActual, failed.Diagnostic.BeforeSizeCorrection!.Width);
        Assert.AreEqual(fault == "Rollback", failed.MutationMayHaveOccurred);
        Assert.AreEqual(fault == "Rollback" ? 2 : 1, Count());
        if (fault == "PersistentDrift") Assert.AreEqual(744.25, failed.Diagnostic.Actual!.Width);
        if (fault == "CorrectionThrows") Assert.IsNull(failed.Diagnostic.Actual, "No stale geometry after a partially applied correction.");
        if (fault == "Rollback")
        {
          var retained = InvokeMethod(shapes, "Item", 2)!;
          try { _ = InvokeMethod(retained, "Delete"); }
          finally { Release(retained); }
        }
      }
      ExcelImagePlacementService.PlacementStageObserved = null;
      var cleanCorrections = 0;
      ExcelImagePlacementService.PlacementStageObserved = (stage, _) =>
      { if (stage == "AfterSizeCorrection") cleanCorrections++; };
      var clean = service.PlaceImage(identity, "OtherTarget", new(6, 4), EvidenceSide.New,
        path, new(745.5, 375), 745.5);
      Assert.IsTrue(clean.Succeeded, clean.Message);
      Assert.AreEqual(0, cleanCorrections);
      Assert.IsTrue(service.DeletePlacedImage(identity, "OtherTarget", clean.ShapeName).Succeeded);

      SetRangeProperty(sheet, "A1:AF202", "NumberFormat", "0.00");
      SetCellValue(sheet, 2, 3, "NEW"); SetCellValue(sheet, 2, 18, "OLD");
      SetCellValue(sheet, 3, 1, 1); SetCellValue(sheet, 3, 2, 1);
      SetCellValue(sheet, 13, 1, 1); SetCellValue(sheet, 13, 2, 2);
      using (var bitmap = new Bitmap(120, 80)) bitmap.Save(path, ImageFormat.Png);
      var automatic = new ExcelAutomaticPlacementService();
      var images = new[] { new AutomaticPlacementImage(path, new(120, 80)) };
      var automaticCorrections = 0;
      ExcelImagePlacementService.PlacementStageObserved = (stage, shape) =>
      {
        if (stage == "AfterAttributes")
        { SetProperty(shape, "LockAspectRatio", 0); SetProperty(shape, "Height", Read(shape, "Height") - 0.3); }
        if (stage == "AfterSizeCorrection")
        {
          automaticCorrections++;
          SetProperty(shape, "LockAspectRatio", 0);
          SetProperty(shape, "Height", Read(shape, "Height") - 0.3);
          SetProperty(shape, "LockAspectRatio", -1);
        }
      };
      var first = automatic.PlaceImages(identity, "OtherTarget", EvidenceSide.New, images, requestedCaseLabel: "1-1");
      Assert.IsTrue(first.Succeeded, first.Message);
      var pair = automatic.PlaceImages(identity, "OtherTarget", EvidenceSide.Old, images, requestedCaseLabel: "1-1");
      Assert.IsTrue(pair.Succeeded, pair.Message);
      Assert.AreEqual(2, automaticCorrections);
      var before = first.PlacedImages[0].Target;
      var old = pair.PlacedImages[0];
      Assert.AreEqual(before.TopPoints, old.Target.TopPoints, 0.05);
      Assert.AreEqual(old.Plan.Image.WidthPoints, old.Target.WidthPoints, 0.05);
      Assert.AreEqual(old.Plan.Image.HeightPoints, old.Target.HeightPoints, 1);
      var managed = new ExcelManagedShapeService();
      Assert.IsTrue(managed.Delete(identity, old.Target).Succeeded);
      Assert.IsTrue(pair.ReferenceResize!.SetApplied(identity, false).Succeeded);
      Assert.IsTrue(pair.ReferenceResize.SetApplied(identity, true).Succeeded);
      var redone = service.PlaceImage(identity, "OtherTarget", old.FocusCell, EvidenceSide.Old,
        path, images[0].Dimensions, old.AvailableWidthPoints,
        scaleOverride: old.Plan.Image.Scale, verticalOffsetPoints: old.Plan.VerticalOffsetPoints);
      Assert.IsTrue(redone.Succeeded, redone.Message);
      Assert.AreEqual(3, automaticCorrections);
      Assert.AreEqual(before.TopPoints, redone.Target!.TopPoints, 0.05);
      Assert.AreEqual(old.Target.WidthPoints, redone.Target.WidthPoints, 0.05);
      Assert.AreEqual(old.Target.HeightPoints, redone.Target.HeightPoints, 0.05);
      Assert.IsTrue(managed.Delete(identity, redone.Target).Succeeded);
      Assert.IsTrue(managed.Delete(identity, managed.Inspect(identity, "OtherTarget", before.ShapeName).Shape!).Succeeded);

      // Drift above 1pt in the second image compensates the first image and added rows.
      var snapshots = new ExcelSheetSnapshotService();
      var previous = snapshots.Capture(identity, "OtherTarget").Snapshot!;
      var inserted = 0;
      ExcelImagePlacementService.PlacementStageObserved = (stage, shape) =>
      {
        if (stage == "AfterAttributes")
        {
          inserted++;
          SetProperty(shape, "LockAspectRatio", 0);
          SetProperty(shape, "Height", Read(shape, "Height") - 0.3);
        }
        if (stage == "AfterSizeCorrection" && inserted == 2)
        { SetProperty(shape, "LockAspectRatio", 0); SetProperty(shape, "Height", Read(shape, "Height") - 1.25); }
      };
      var compensated = automatic.PlaceImages(identity, "OtherTarget", EvidenceSide.New,
        [images[0], images[0]], requestedCaseLabel: "1-1");
      Assert.IsFalse(compensated.Succeeded);
      Assert.IsTrue(compensated.CompensationSucceeded, compensated.Message);
      Assert.IsTrue(compensated.ImageFailure!.Diagnostic!.SizeCorrectionAttempted);
      Assert.IsFalse(compensated.ImageFailure.MutationMayHaveOccurred);
      var after = snapshots.Capture(identity, "OtherTarget").Snapshot!;
      CollectionAssert.AreEqual(previous.LayoutSignals.Anchors.ToArray(), after.LayoutSignals.Anchors.ToArray());
      Assert.AreEqual(1, Count());
      var cell = GetRequiredProperty(sheet, "Cells", 1, 1);
      try { Assert.AreEqual("keep existing content", GetRequiredProperty(cell, "Value2")); }
      finally { Release(cell); }
      Assert.AreEqual(1, Count());
    }
    finally { ExcelImagePlacementService.PlacementStageObserved = null; Release(shapes); }
  }
}
