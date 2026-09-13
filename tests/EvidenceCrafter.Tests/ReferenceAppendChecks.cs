using System.Drawing.Imaging;
using System.Globalization;
using System.Security.Cryptography;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  private static void VerifyReferenceAppend(object workbooks, string directory, string imagePath)
  {
    const string referenceDirectory = @"C:\work\Macro\Case&Evidence\refer";
    if (!Directory.Exists(referenceDirectory)) throw new OfficeUnavailableException("Reference workbooks are unavailable.");
    using (var bitmap = new Bitmap(120, 80))
    {
      using var graphics = Graphics.FromImage(bitmap);
      graphics.Clear(Color.CornflowerBlue);
      bitmap.Save(imagePath, ImageFormat.Png);
    }
    foreach (var source in Directory.GetFiles(referenceDirectory, "*.xlsx"))
    {
      var originalHash = SHA256.HashData(File.ReadAllBytes(source));
      var copy = Path.Combine(directory, "reference-copy.xlsx");
      File.Copy(source, copy);
      object? workbook = null;
      object? worksheets = null;
      object? sheet = null;
      object? shapes = null;
      try
      {
        workbook = InvokeMethod(workbooks, "Open", copy)!;
        worksheets = GetRequiredProperty(workbook, "Worksheets");
        sheet = GetRequiredProperty(worksheets, "Item", "B1");
        _ = InvokeMethod(workbook, "Activate");
        _ = InvokeMethod(sheet, "Activate");
        var discovery = new ExcelSessionCatalog().Discover();
        var identity = discovery.Workbooks.SingleOrDefault(item => string.Equals(item.FullPath, copy, StringComparison.OrdinalIgnoreCase));
        Assert.IsNotNull(identity, $"Open={GetRequiredProperty(workbook, "FullName")}; found={string.Join(";", discovery.Workbooks.Select(item => item.FullPath))}; warnings={string.Join(";", discovery.Warnings)}");
        var service = new ExcelAutomaticPlacementService();
        var request = new[] { new AutomaticPlacementImage(imagePath, new ImageDimensions(120, 80)) };
        shapes = GetRequiredProperty(sheet, "Shapes");
        var oldAnalysis = service.Analyze(identity, "B1", EvidenceSide.Old, request, requestedCaseLabel: "1-1");
        if (oldAnalysis.Succeeded)
        {
          var tallRequest = new[] { new AutomaticPlacementImage(imagePath, new ImageDimensions(120, 800)) };
          var backfilledShapes = new List<(string Name, string CaseLabel)>();
          var boundaryRow = 0;
          foreach (var (label, side) in new[] { ("1-1", EvidenceSide.New), ("1-2", EvidenceSide.New), ("1-2", EvidenceSide.Old), ("1-1", EvidenceSide.Old) })
          {
            if (label == "1-1" && side == EvidenceSide.Old)
            {
              var before = service.Analyze(identity, "B1", side, tallRequest, requestedCaseLabel: label);
              boundaryRow = before.LayoutAnalysis!.Layout!.EndRow;
              SetRangeProperty(sheet, $"D{boundaryRow}", "Value2", "Existing boundary content");
            }
            var placed = service.PlaceImages(identity, "B1", side, tallRequest, requestedCaseLabel: label);
            Assert.IsTrue(placed.Succeeded, $"{Path.GetFileName(source)} {label} {side}: {placed.Message}");
            backfilledShapes.Add((placed.PlacedImages[0].ShapeName, label));
            if (placed.Analysis!.CompletesCaseAfterPlacement)
            {
              var trim = new ExcelCaseMaintenanceService().TrimCompletedCaseTail(identity, "B1", placed.PlacedImages[0].FocusCell.Row);
              trim.DeletionSnapshot?.Dispose();
            }
          }
          var boundary = GetRequiredProperty(sheet, "Range", $"D{boundaryRow}");
          try
          {
            object? boundaryValue = null;
            try { boundaryValue = GetRequiredProperty(boundary, "Value2"); } catch (InvalidOperationException) { }
            if (!Equals(boundaryValue, "Existing boundary content"))
            {
              object? used = null;
              object? found = null;
              var foundContent = false;
              try
              {
                used = GetRequiredProperty(sheet, "UsedRange");
                found = InvokeMethod(used, "Find", "Existing boundary content");
                foundContent = found is not null;
              }
              finally { Release(found); Release(used); }
              Assert.IsTrue(foundContent, $"{Path.GetFileName(source)}: boundary content was lost (original row {boundaryRow}).");
            }
          }
          finally { Release(boundary); }
          var bounds = new List<RectangleF>();
          var finalSnapshot = new ExcelSheetSnapshotService().Capture(identity, "B1", 3).Snapshot!;
          var finalAnchors = ExcelAutomaticPlacementService.ConfirmedAnchors(finalSnapshot.LayoutSignals);
          foreach (var (name, caseLabel) in backfilledShapes)
          {
            var shape = InvokeMethod(shapes, "Item", name)!;
            try
            {
              var rectangle = ShapeBounds(shape);
              Assert.IsFalse(bounds.Any(rectangle.IntersectsWith), "Backfilled NEW/OLD images must not overlap.");
              bounds.Add(rectangle);
              var anchorIndex = Array.FindIndex(finalAnchors,
                anchor => ExcelAutomaticPlacementService.FormatCaseLabel(anchor) == caseLabel);
              Assert.IsGreaterThanOrEqualTo(0, anchorIndex, $"CASE {caseLabel} must still exist after row mutations.");
              var nextAnchorRow = anchorIndex + 1 < finalAnchors.Length
                ? finalAnchors[anchorIndex + 1].Row
                : ExcelWorksheetLimits.MaximumRow + 1;
              var snapshotShape = finalSnapshot.Shapes.Single(item => item.Name == name);
              Assert.IsGreaterThanOrEqualTo(finalAnchors[anchorIndex].Row + 1, snapshotShape.StartRow,
                $"{name} starts outside CASE {caseLabel}.");
              Assert.IsLessThan(nextAnchorRow, snapshotShape.EndRow,
                $"{name} intrudes from CASE {caseLabel} into the next CASE at row {nextAnchorRow}.");
            }
            finally { Release(shape); }
          }
          Console.WriteLine($"BACKFILL verified: {Path.GetFileName(source)}");
        }
        for (var pass = 0; pass < 3; pass++)
        {
          var result = service.PlaceImages(identity, "B1", EvidenceSide.New, request, requestedCaseLabel: "1-1");
          Assert.IsTrue(result.Succeeded, $"{Path.GetFileName(source)} pass {pass}: {result.Message}");
          var added = InvokeMethod(shapes, "Item", result.PlacedImages[0].ShapeName)!;
          try
          {
            var bounds = ShapeBounds(added);
            var count = Convert.ToInt32(GetRequiredProperty(shapes, "Count"), CultureInfo.InvariantCulture);
            for (var i = 1; i <= count; i++)
            {
              var existing = InvokeMethod(shapes, "Item", i)!;
              try
              {
                if (Equals(GetRequiredProperty(existing, "Name"), result.PlacedImages[0].ShapeName)) continue;
                Assert.IsFalse(bounds.IntersectsWith(ShapeBounds(existing)), $"{Path.GetFileName(source)}: added image overlaps an existing shape.");
              }
              finally { ReleaseOnce(existing); }
            }
            // Exercise a user-moved, enlarged and no longer managed picture on the next append.
            if (pass == 0)
            {
              SetProperty(added, "Name", "ManuallyMovedPicture");
              SetProperty(added, "Top", bounds.Top + 15);
              SetProperty(added, "Height", bounds.Height + 20);
              SetRangeProperty(sheet, "8:8", "RowHeight", 30);
              var hiddenRow = GetRequiredProperty(sheet, "Rows", 9);
              try { SetProperty(hiddenRow, "Hidden", true); }
              finally { Release(hiddenRow); }
            }
          }
          finally { Release(added); }
        }
        Console.WriteLine($"APPEND verified: {Path.GetFileName(source)}, 3 images, no rectangle intersections.");
      }
      finally
      {
        Release(shapes); Release(sheet); Release(worksheets);
        if (workbook is not null) { TryInvoke(workbook, "Close", false); Release(workbook); }
        File.Delete(copy);
        CollectionAssert.AreEqual(originalHash, SHA256.HashData(File.ReadAllBytes(source)), "Reference source was changed.");
      }
    }
  }

  private static RectangleF ShapeBounds(object shape) => new(
    Convert.ToSingle(GetRequiredProperty(shape, "Left"), CultureInfo.InvariantCulture),
    Convert.ToSingle(GetRequiredProperty(shape, "Top"), CultureInfo.InvariantCulture),
    Convert.ToSingle(GetRequiredProperty(shape, "Width"), CultureInfo.InvariantCulture),
    Convert.ToSingle(GetRequiredProperty(shape, "Height"), CultureInfo.InvariantCulture));
}
