using System.Diagnostics;
using System.Drawing.Imaging;
using System.Reflection;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  private static void VerifySameSidePerformance(object sheet, WorkbookIdentity identity, string imagePath)
  {
    var logPath = Environment.GetEnvironmentVariable("EVIDENCECRAFTER_PERF_LOG");
    void Report(string line)
    {
      Console.WriteLine(line);
      if (!string.IsNullOrWhiteSpace(logPath)) File.AppendAllText(logPath, line + Environment.NewLine);
    }
    SetRangeProperty(sheet, "A1:AF1602", "NumberFormat", "0.00");
    SetRangeProperty(sheet, "A1:AF1602", "RowHeight", 15);
    SetCellValue(sheet, 2, 3, "NEW");
    SetCellValue(sheet, 2, 18, "OLD");
    SetCellValue(sheet, 3, 1, 1);
    SetCellValue(sheet, 3, 2, 1);
    SetCellValue(sheet, 13, 1, 1);
    SetCellValue(sheet, 13, 2, 2);
    using (var bitmap = new Bitmap(120, 80)) bitmap.Save(imagePath, ImageFormat.Png);
    var images = new[] { new AutomaticPlacementImage(imagePath, new ImageDimensions(120, 80)) };
    var service = new ExcelAutomaticPlacementService();
    var first = service.PlaceImages(identity, "OtherTarget", EvidenceSide.New, images, requestedCaseLabel: "1-1");
    Assert.IsTrue(first.Succeeded, first.Message);
    object? collection = null, application = null;
    // The baseline binary has no observer; keep the same workload compilable there.
    var observer = typeof(ExcelSheetSnapshotService).GetField("ShapeReadObserved", BindingFlags.Static | BindingFlags.NonPublic);
    var scans = 0;
    var visited = 0;
    var detailed = 0;
    try
    {
      observer?.SetValue(null, (Action<int, int>)((all, details) => { scans++; visited += all; detailed += details; }));
      collection = GetRequiredProperty(sheet, "Shapes");
      application = GetRequiredProperty(sheet, "Application");
      var added = 0;
      foreach (var count in new[] { 0, 50, 200, 500 })
      {
        for (; added < count; added++)
        {
          var other = InvokeMethod(collection, "AddShape", 1, 100f, 2000f + added * 30, 120f, 20f)!;
          try { SetProperty(other, "Name", $"Unrelated_{added}"); }
          finally { Release(other); }
        }
        var previewTimes = new List<double>();
        var commitTimes = new List<double>();
        var placementTimes = new List<double>();
        _ = InvokeMethod(sheet, "Copy", Type.Missing, sheet);
        var copy = GetRequiredProperty(application, "ActiveSheet");
        RowDeletionSnapshot? saved = null;
        try
        {
          var target = $"Perf_{count}";
          SetProperty(copy, "Name", target);
          for (var iteration = 0; iteration < 11; iteration++)
          {
            var timer = Stopwatch.StartNew();
            var preview = service.Analyze(identity, target, EvidenceSide.Old, images, requestedCaseLabel: "1-1");
            Assert.IsTrue(preview.Succeeded, preview.Message);
            var previewMs = timer.Elapsed.TotalMilliseconds;
            scans = visited = detailed = 0;
            timer.Restart();
            var placed = service.PlaceImages(identity, target, EvidenceSide.Old, images,
              preparedAnalysis: preview, requestedCaseLabel: "1-1");
            Assert.IsTrue(placed.Succeeded, placed.Message);
            var placementMs = timer.Elapsed.TotalMilliseconds;
            // Measure placement repeatedly; exercise row-backup cleanup once per size.
            // Repeated native backup sheets destabilize long Excel automation runs.
            if (iteration == 10)
            {
              var cleanup = new ExcelCaseMaintenanceService().TrimCompletedCaseTail(identity, target, 3);
              saved = cleanup.DeletionSnapshot;
              Assert.IsTrue(cleanup.Succeeded, cleanup.Message);
            }
            var cleanupMs = timer.Elapsed.TotalMilliseconds - placementMs;
            var next = new ExcelCaseNavigationService().Navigate(identity, target,
              CaseNavigationDirection.Next, "1-1", EvidenceSide.Old, sameCaseThenNext: false);
            Assert.IsTrue(next.Succeeded, next.Message);
            var totalMs = timer.Elapsed.TotalMilliseconds;
            if (iteration > 0)
            {
              previewTimes.Add(previewMs);
              commitTimes.Add(totalMs - cleanupMs);
              placementTimes.Add(placementMs);
              Report($"PERF backfill/{count}/{iteration}: preview={previewMs:F2}; place={placementMs:F2}; cleanup={cleanupMs:F2}; navigation={totalMs - placementMs - cleanupMs:F2}; total={totalMs:F2} ms");
              if (observer is not null) Report($"READS backfill/{count}: snapshots={scans}; visited={visited}; details={detailed}");
            }
            if (iteration < 10)
            {
              Assert.IsTrue(new ExcelManagedShapeService().Delete(identity, placed.PlacedImages[0].Target).Succeeded);
              if (placed.ReferenceResize is not null)
                Assert.IsTrue(placed.ReferenceResize.SetApplied(identity, false).Succeeded);
            }
          }
        }
        finally
        {
          saved?.Dispose();
          _ = InvokeMethod(copy, "Delete");
          Release(copy);
        }
        Report($"PERF backfill/{count}: preview median={(previewTimes.Order().ElementAt(4) + previewTimes.Order().ElementAt(5)) / 2:F2}; place median={(placementTimes.Order().ElementAt(4) + placementTimes.Order().ElementAt(5)) / 2:F2}; place+navigation median={(commitTimes.Order().ElementAt(4) + commitTimes.Order().ElementAt(5)) / 2:F2}; max={commitTimes.Max():F2} ms");
      }
    }
    finally { observer?.SetValue(null, null); Release(collection); ReleaseOnce(application); }
  }
}
