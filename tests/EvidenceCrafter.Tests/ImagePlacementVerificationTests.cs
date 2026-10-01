using System.Globalization;
using System.Text.Json;
using EvidenceCrafter.App;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class ImagePlacementVerificationTests
{
  [TestMethod]
  public void PositionComparison_PreservesZeroAndRejectsInvalidCoordinates()
  {
    Assert.IsTrue(ExcelImagePlacementService.PositionMatches(0, 0));
    Assert.IsFalse(ExcelImagePlacementService.PositionMatches(0, 64));
    Assert.IsFalse(ExcelImagePlacementService.PositionMatches(-1, -1));
    Assert.IsFalse(ExcelImagePlacementService.PositionMatches(double.NaN, 0));
    Assert.IsFalse(ExcelImagePlacementService.PositionMatches(0, double.PositiveInfinity));
    Assert.IsFalse(ExcelImagePlacementService.PositionMatches(double.MaxValue, double.MaxValue));
  }

  [TestMethod]
  public void PositionComparison_AcceptsNearestExcelCoordinateButRejectsAnotherPosition()
  {
    const double planned = 1874212.56;
    Assert.IsTrue(ExcelImagePlacementService.PositionMatches(planned, 1874212.5));
    Assert.IsFalse(ExcelImagePlacementService.PositionMatches(planned, 1874212.375));
    Assert.IsFalse(ExcelImagePlacementService.PositionMatches(77, 77.1));
  }

  [TestMethod]
  public void GeometryVerification_KeepsSizesStrictAndReportsRawZero()
  {
    var planned = new ImagePlacementGeometry(0, 0, 120, 80);
    Assert.IsNull(ExcelImagePlacementService.VerifyGeometry(planned, planned));
    foreach (var actual in new[] { planned with { Width = 0 }, planned with { Height = 0 },
      planned with { Width = 120.06 }, planned with { Height = double.NaN } })
      Assert.IsNotNull(ExcelImagePlacementService.VerifyGeometry(planned, actual));
    StringAssert.Contains(ExcelImagePlacementService.VerifyGeometry(planned, planned with { Width = 0 })!, "実際 0pt");
  }

  [TestMethod]
  public void FailureDiagnostic_RecordsGeometryWithoutWorkbookOrImageContents()
  {
    var folder = Path.Combine(Path.GetTempPath(), "EvidenceCrafter.Tests", Guid.NewGuid().ToString("N"));
    try
    {
      var diagnostic = new ImagePlacementDiagnostic("Verify", "16.0", "19127", new CellReference(6, 19),
        EvidenceSide.Old, new ImageDimensions(120, 80), 1, new(978, 77, 120, 80), new(978, 77, 0, 80), false, true);
      new DiagnosticLog(folder).Write(DiagnosticEventKind.MutationResult, DiagnosticOutcome.Failed, placement: diagnostic);
      using var entry = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "diagnostic.jsonl")));
      Assert.AreEqual(0d, entry.RootElement.GetProperty("placement").GetProperty("actual").GetProperty("width").GetDouble());
      Assert.AreEqual(19, entry.RootElement.GetProperty("placement").GetProperty("cell").GetProperty("column").GetInt32());
      Assert.IsFalse(entry.RootElement.TryGetProperty("workbookPath", out _));
    }
    finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
  }
}

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  private static void VerifyInsertedImageGeometry(object sheet, WorkbookIdentity identity, string imagePath)
  {
    using (var bitmap = new System.Drawing.Bitmap(120, 80)) bitmap.Save(imagePath, System.Drawing.Imaging.ImageFormat.Png);
    var service = new ExcelImagePlacementService();
    var shapes = GetRequiredProperty(sheet, "Shapes");
    int ShapeCount() => Convert.ToInt32(GetRequiredProperty(shapes, "Count"), CultureInfo.InvariantCulture);
    var sentinel = InvokeMethod(shapes, "AddShape", 1, 1800f, 0f, 30f, 30f)!;
    SetProperty(sentinel, "Name", "ExistingImage");
    Release(sentinel);
    SetCellValue(sheet, 1, 1, "existing cell");
    try
    {
      SetRangeProperty(sheet, "A1:AF202", "RowHeight", 15.75);
      SetRangeProperty(sheet, "A1:AF202", "NumberFormat", "0.00");
      SetRangeProperty(sheet, "10:10", "Hidden", true);
      foreach (var (row, column, offset) in new[] { (1, 1, 0d), (6, 19, 2d), (100000, 4, 3.06) })
      {
        var cell = GetRequiredProperty(sheet, "Cells", row, column);
        ImagePlacementResult placed;
        try
        {
          var top = Convert.ToDouble(GetRequiredProperty(cell, "Top"), CultureInfo.InvariantCulture) + offset;
          var left = Convert.ToDouble(GetRequiredProperty(cell, "Left"), CultureInfo.InvariantCulture);
          placed = service.PlaceImage(identity, "OtherTarget", new(row, column), EvidenceSide.New,
            imagePath, new(120, 80), 120, horizontalMarginPoints: 0, verticalOffsetPoints: offset);
          Assert.IsTrue(placed.Succeeded, placed.Message);
          var raw = InvokeMethod(shapes, "Item", placed.ShapeName)!;
          try
          {
            var actual = Convert.ToDouble(GetRequiredProperty(raw, "Top"), CultureInfo.InvariantCulture);
            Assert.AreEqual(placed.Target!.TopPoints, actual);
            Assert.AreEqual((double)(float)top, actual, 0.05);
            Assert.AreEqual(left, placed.Target.LeftPoints, 0.05);
            Assert.AreEqual(120d, placed.Target.WidthPoints, 0.05);
            Assert.AreEqual(80d, placed.Target.HeightPoints, 0.05);
          }
          finally { Release(raw); }
        }
        finally { Release(cell); }
        Assert.IsTrue(service.DeletePlacedImage(identity, "OtherTarget", placed.ShapeName).Succeeded);
        var redone = service.PlaceImage(identity, "OtherTarget", new(row, column), EvidenceSide.New,
          imagePath, new(120, 80), 120, horizontalMarginPoints: 0, verticalOffsetPoints: offset);
        Assert.IsTrue(redone.Succeeded, redone.Message);
        Assert.AreEqual(placed.Target!.TopPoints, redone.Target!.TopPoints);
        Assert.IsTrue(service.DeletePlacedImage(identity, "OtherTarget", redone.ShapeName).Succeeded);
      }

      SetRangeProperty(sheet, "A:XFD", "Hidden", true);
      var hidden = service.PlaceImage(identity, "OtherTarget", new(6, 19), EvidenceSide.Old,
        imagePath, new(120, 80), 120);
      Assert.IsFalse(hidden.Succeeded);
      Assert.IsFalse(hidden.MutationMayHaveOccurred);
      Assert.IsTrue(hidden.Diagnostic!.ColumnHidden);
      StringAssert.Contains(hidden.Message, "非表示");
      Assert.AreEqual(1, ShapeCount());
      SetRangeProperty(sheet, "A:XFD", "Hidden", false);

      foreach (var fault in new[] { "Width", "Height", "WidthZero", "HeightZero", "Left", "Top", "Name", "BeforeInsert", "Rollback", "RollbackUnknown" })
      {
        ExcelImagePlacementService.PlacementStageObserved = (stage, shape) =>
        {
          if (stage == "BeforeInsert" && fault == "BeforeInsert") throw new InvalidOperationException("Injected before insert");
          if (stage == "AfterAttributes")
          {
            if (fault == "Name") SetProperty(shape, "Name", "UnexpectedName");
            else if (fault != "BeforeInsert")
            {
              SetProperty(shape, "LockAspectRatio", 0);
              SetProperty(shape, fault is "Rollback" or "RollbackUnknown" ? "Width" : fault.Replace("Zero", "", StringComparison.Ordinal),
                fault.EndsWith("Zero", StringComparison.Ordinal) ? 0d : fault is "Top" or "Left" ? 300d : 1d);
            }
          }
          if (stage == "BeforeRollback" && fault == "Rollback") throw new InvalidOperationException("Injected delete failure");
          if (stage == "AfterRollback" && fault == "RollbackUnknown") throw new InvalidOperationException("Injected unverified delete");
        };
        var failed = service.PlaceImage(identity, "OtherTarget", new(6, 19), EvidenceSide.Old,
          imagePath, new(120, 80), 120);
        ExcelImagePlacementService.PlacementStageObserved = null;
        Assert.IsFalse(failed.Succeeded);
        Assert.IsNotNull(failed.Diagnostic);
        if (fault == "WidthZero") Assert.AreEqual(0d, failed.Diagnostic.Actual!.Width);
        if (fault == "HeightZero") Assert.AreEqual(0d, failed.Diagnostic.Actual!.Height);
        Assert.AreEqual(fault is "Rollback" or "RollbackUnknown", failed.MutationMayHaveOccurred, fault);
        Assert.AreEqual(fault == "Rollback" ? 2 : 1, ShapeCount(), fault);
        if (fault == "Rollback")
        {
          var retained = InvokeMethod(shapes, "Item", 2)!;
          try { _ = InvokeMethod(retained, "Delete"); }
          finally { Release(retained); }
        }
      }

      SetCellValue(sheet, 2, 3, "NEW"); SetCellValue(sheet, 2, 18, "OLD");
      SetCellValue(sheet, 3, 1, 1); SetCellValue(sheet, 3, 2, 1);
      SetCellValue(sheet, 13, 1, 1); SetCellValue(sheet, 13, 2, 2);
      SetCellValue(sheet, 15, 3, "next case marker");
      var stageCount = 0;
      ExcelImagePlacementService.PlacementStageObserved = (stage, shape) =>
      {
        if (stage == "AfterAttributes" && ++stageCount == 2)
        { SetProperty(shape, "LockAspectRatio", 0); SetProperty(shape, "Width", 1d); }
        if (stage == "BeforeRollback") throw new InvalidOperationException("Injected rollback failure");
      };
      var automatic = new ExcelAutomaticPlacementService().PlaceImages(identity, "OtherTarget", EvidenceSide.New,
        [new(imagePath, new(120, 80)), new(imagePath, new(120, 80))], requestedCaseLabel: "1-1");
      Assert.IsFalse(automatic.Succeeded, automatic.Message);
      Assert.IsFalse(automatic.CompensationSucceeded, automatic.Message);
      Assert.IsTrue(automatic.ImageFailure!.MutationMayHaveOccurred);
      Assert.HasCount(1, automatic.PlacedImages);
      Assert.IsNotEmpty(automatic.AppliedInsertions);
      Assert.AreEqual(3, ShapeCount(), "Do not delete a prior image or compensate rows after unverified rollback.");
      Assert.IsFalse(automatic.Message.Contains("変更を取り消しました", StringComparison.Ordinal));
      var shifted = GetRequiredProperty(sheet, "Cells", 15 + automatic.AppliedInsertions.Sum(row => row.Count), 3);
      try { Assert.AreEqual("next case marker", GetRequiredProperty(shifted, "Value2")); }
      finally { Release(shifted); }
      var original = GetRequiredProperty(sheet, "Cells", 1, 1);
      try { Assert.AreEqual("existing cell", GetRequiredProperty(original, "Value2")); }
      finally { Release(original); }
    }
    finally { ExcelImagePlacementService.PlacementStageObserved = null; Release(shapes); }
  }
}
