using System.Diagnostics;
using System.Drawing.Imaging;
using System.Globalization;
using System.Reflection;
using System.Text.Json;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  private sealed record BackfillMeasurement(double Preview, double Placement, double Cleanup, double Navigation,
    double Total, int Snapshots, int Visited, int Detailed, int TopBounds, int BottomBounds,
    Dictionary<string, double> Stages, Dictionary<string, int> StageCalls, bool HasSnapshotObserver,
    bool HasSnapshotBoundsObserver, bool HasSafetyBoundsObserver, bool HasStageObserver);

  private static void RunSameSidePerformanceSamples()
  {
    var counts = (Environment.GetEnvironmentVariable("EVIDENCECRAFTER_PERF_COUNTS") ?? "0,50,200,500")
      .Split(',').Select(value => int.Parse(value, CultureInfo.InvariantCulture)).ToArray();
    var samples = int.Parse(Environment.GetEnvironmentVariable("EVIDENCECRAFTER_PERF_SAMPLES") ?? "10", CultureInfo.InvariantCulture);
    var warmups = int.Parse(Environment.GetEnvironmentVariable("EVIDENCECRAFTER_PERF_WARMUPS") ?? "1", CultureInfo.InvariantCulture);
    var position = Environment.GetEnvironmentVariable("EVIDENCECRAFTER_PERF_POSITION") ?? "front";
    var format = Environment.GetEnvironmentVariable("EVIDENCECRAFTER_PERF_FORMAT") ?? "uniform";
    var deletion = Environment.GetEnvironmentVariable("EVIDENCECRAFTER_PERF_DELETION") ?? "few";
    Assert.IsTrue(counts.Length > 0 && counts.All(count => count is >= 0 and <= 500));
    Assert.IsTrue(samples is >= 1 and <= 100 && warmups is >= 0 and <= 10);
    Assert.IsTrue(position is "front" or "middle" or "back");
    Assert.IsTrue(format is "uniform" or "mixed");
    Assert.IsTrue(deletion is "none" or "few" or "many");
    foreach (var count in counts)
    {
      var measurements = new List<BackfillMeasurement>();
      for (var iteration = 0; iteration < warmups + samples; iteration++)
      {
        BackfillMeasurement? measurement = null;
        try
        {
          // Every sample owns a fresh process, workbook and fixture. Startup and
          // teardown remain supervised, but are outside the operation stopwatch.
          RunSupervisedScenario(Scenario.SameSidePerformance, (sheet, identity, path) =>
            measurement = MeasureSameSidePerformance(sheet, identity, path, count, position, format, deletion));
          Assert.IsNotNull(measurement);
          ReportPerformance(JsonSerializer.Serialize(new
          {
            Kind = "sample", Count = count, Iteration = iteration, Warmup = iteration < warmups,
            Position = position, Format = format, Deletion = deletion, Measurement = measurement,
          }));
          if (iteration >= warmups) measurements.Add(measurement);
        }
        catch (Exception exception)
        {
          ReportPerformance(JsonSerializer.Serialize(new
          {
            Kind = "failure", Count = count, Iteration = iteration, Position = position,
            Format = format, Deletion = deletion, Error = exception.ToString(),
          }));
          throw;
        }
      }
      double Median(Func<BackfillMeasurement, double> select)
      {
        var ordered = measurements.Select(select).Order().ToArray();
        return (ordered[(ordered.Length - 1) / 2] + ordered[ordered.Length / 2]) / 2;
      }
      ReportPerformance(JsonSerializer.Serialize(new
      {
        Kind = "summary", Count = count, Samples = samples, Position = position, Format = format, Deletion = deletion,
        PreviewMedian = Median(item => item.Preview), PlacementMedian = Median(item => item.Placement),
        CleanupMedian = Median(item => item.Cleanup), NavigationMedian = Median(item => item.Navigation),
        TotalMedian = Median(item => item.Total), TotalMaximum = measurements.Max(item => item.Total),
      }));
    }
  }

  private static void ReportPerformance(string line)
  {
    Console.WriteLine(line);
    var path = Environment.GetEnvironmentVariable("EVIDENCECRAFTER_PERF_LOG");
    if (!string.IsNullOrWhiteSpace(path)) File.AppendAllText(path, line + Environment.NewLine);
  }

  private static BackfillMeasurement MeasureSameSidePerformance(object sheet, WorkbookIdentity identity,
    string imagePath, int shapeCount, string position, string format, string deletion)
  {
    var caseRow = position switch { "middle" => 703, "back" => 1503, _ => 3 };
    var nextRow = caseRow + (deletion == "many" ? 50 : 10);
    var caseLabel = caseRow == 3 ? "1-1" : "1-2";
    SetRangeProperty(sheet, "A1:AF1602", "NumberFormat", "0.00");
    SetRangeProperty(sheet, "A1:AF1602", "RowHeight", 15);
    if (format == "mixed")
    {
      SetRangeProperty(sheet, "D1:D1602", "NumberFormat", "@");
      SetRangeProperty(sheet, "F1:F1602", "WrapText", true);
    }
    SetCellValue(sheet, 2, 3, "NEW");
    SetCellValue(sheet, 2, 18, "OLD");
    SetCellValue(sheet, 3, 1, 1);
    SetCellValue(sheet, 3, 2, 1);
    SetCellValue(sheet, caseRow, 1, 1);
    SetCellValue(sheet, caseRow, 2, caseRow == 3 ? 1 : 2);
    SetCellValue(sheet, nextRow, 1, 1);
    SetCellValue(sheet, nextRow, 2, caseRow == 3 ? 2 : 3);
    using (var bitmap = new Bitmap(120, 80)) bitmap.Save(imagePath, ImageFormat.Png);
    var images = new[] { new AutomaticPlacementImage(imagePath, new ImageDimensions(120, 80)) };
    var service = new ExcelAutomaticPlacementService();
    var first = service.PlaceImages(identity, "OtherTarget", EvidenceSide.New, images, requestedCaseLabel: caseLabel);
    Assert.IsTrue(first.Succeeded, first.Message);
    object? shapes = null;
    try
    {
      shapes = GetRequiredProperty(sheet, "Shapes");
      for (var index = 0; index < shapeCount; index++)
      {
        var top = 2000f + index * 30;
        // Preserve the baseline front fixture; alternative CASE positions keep
        // all unrelated pictures outside the target and its possible growth.
        if (position == "middle" && top >= (caseRow - 2) * 15) top += 1500;
        var shape = InvokeMethod(shapes, "AddShape", 1, 100f, top, 120f, 20f)!;
        try { SetProperty(shape, "Name", $"Unrelated_{index}"); }
        finally { Release(shape); }
      }
    }
    finally { Release(shapes); }

    var flags = BindingFlags.Static | BindingFlags.NonPublic;
    var detailsObserver = typeof(ExcelSheetSnapshotService).GetField("ShapeReadObserved", flags);
    var snapshotBoundsObserver = typeof(ExcelSheetSnapshotService).GetField("ShapeBoundsReadObserved", flags);
    var safetyBoundsObserver = typeof(ExcelRowMutationService).GetField("ShapeBoundsReadObserved", flags);
    var stageObserver = typeof(ExcelRowMutationService).GetField("StageMeasured", flags);
    var observers = new[] { detailsObserver, snapshotBoundsObserver, safetyBoundsObserver, stageObserver };
    var previous = observers.Select(field => field?.GetValue(null)).ToArray();
    var snapshots = 0;
    var visited = 0;
    var detailed = 0;
    var topBounds = 0;
    var bottomBounds = 0;
    var stages = new Dictionary<string, double>();
    var stageCalls = new Dictionary<string, int>();
    RowDeletionSnapshot? saved = null;
    try
    {
      // Reflection keeps precisely this test workload compilable against the baseline.
      detailsObserver?.SetValue(null, (Action<int, int>)((all, details) => { snapshots++; visited += all; detailed += details; }));
      Action<int, int> bounds = (top, bottom) => { topBounds += top; bottomBounds += bottom; };
      snapshotBoundsObserver?.SetValue(null, bounds);
      safetyBoundsObserver?.SetValue(null, bounds);
      stageObserver?.SetValue(null, (Action<string, double>)((name, milliseconds) =>
      {
        stages[name] = stages.GetValueOrDefault(name) + milliseconds;
        stageCalls[name] = stageCalls.GetValueOrDefault(name) + 1;
      }));
      var clock = Stopwatch.StartNew();
      var preview = service.Analyze(identity, "OtherTarget", EvidenceSide.Old, images, requestedCaseLabel: caseLabel);
      var previewMs = clock.Elapsed.TotalMilliseconds;
      Assert.IsTrue(preview.Succeeded, preview.Message);
      // Counters describe commit through navigation; preview is timed separately.
      snapshots = visited = detailed = topBounds = bottomBounds = 0;
      stages.Clear();
      stageCalls.Clear();
      clock.Restart();
      var placed = service.PlaceImages(identity, "OtherTarget", EvidenceSide.Old, images,
        preparedAnalysis: preview, requestedCaseLabel: caseLabel);
      var placementMs = clock.Elapsed.TotalMilliseconds;
      Assert.IsTrue(placed.Succeeded, placed.Message);
      // 'none' exercises the standard four-row tail maintenance without deleting
      // the already reserved tail. Other variants use the normal two-side cleanup.
      var maintenance = new ExcelCaseMaintenanceService();
      var cleanup = deletion == "none"
        ? maintenance.TrimCaseTail(identity, "OtherTarget", placed.PlacedImages[0].FocusCell.Row, tailRows: 4)
        : maintenance.TrimCompletedCaseTail(identity, "OtherTarget", placed.PlacedImages[0].FocusCell.Row);
      var cleanupEnd = clock.Elapsed.TotalMilliseconds;
      saved = cleanup.DeletionSnapshot;
      Assert.IsTrue(cleanup.Succeeded, cleanup.Message);
      if (deletion == "none") Assert.IsFalse(cleanup.Changed, "The no-deletion fixture must not remove rows.");
      else Assert.IsTrue(cleanup.Changed, "The deletion fixture must exercise native backup and row deletion.");
      var next = new ExcelCaseNavigationService().Navigate(identity, "OtherTarget",
        CaseNavigationDirection.Next, caseLabel, EvidenceSide.Old, sameCaseThenNext: false);
      var totalMs = clock.Elapsed.TotalMilliseconds;
      Assert.IsTrue(next.Succeeded, next.Message);
      return new(previewMs, placementMs, cleanupEnd - placementMs, totalMs - cleanupEnd,
        totalMs, snapshots, visited, detailed, topBounds, bottomBounds, stages, stageCalls,
        detailsObserver is not null, snapshotBoundsObserver is not null, safetyBoundsObserver is not null, stageObserver is not null);
    }
    finally
    {
      for (var index = 0; index < observers.Length; index++) observers[index]?.SetValue(null, previous[index]);
      saved?.Dispose();
    }
  }
}