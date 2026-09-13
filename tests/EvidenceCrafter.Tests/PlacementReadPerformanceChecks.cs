using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  private static void VerifyRowHeightReadPerformance(object worksheet)
  {
    var reader = typeof(ExcelSheetSnapshotService).GetMethod(
      "ReadRowHeights", BindingFlags.NonPublic | BindingFlags.Static)!;
    foreach (var scenario in new[] { "uniform", "one-exception", "bands", "hidden", "all-hidden" })
    {
      object? area = null;
      object? rows = null;
      try
      {
        area = GetRequiredProperty(worksheet, "Range", "A3:A242");
        rows = GetRequiredProperty(area, "EntireRow");
        SetProperty(rows, "Hidden", false);
        SetProperty(rows, "RowHeight", 15);
        if (scenario == "one-exception") SetRangeProperty(worksheet, "123:123", "RowHeight", 30);
        if (scenario == "bands") SetRangeProperty(worksheet, "123:242", "RowHeight", 30);
        if (scenario == "hidden")
        {
          object? hiddenArea = null;
          object? hiddenRow = null;
          try
          {
            hiddenArea = GetRequiredProperty(worksheet, "Range", "A123:A123");
            hiddenRow = GetRequiredProperty(hiddenArea, "EntireRow");
            SetProperty(hiddenRow, "Hidden", true);
          }
          finally { Release(hiddenRow); Release(hiddenArea); }
        }
        if (scenario == "all-hidden") SetProperty(rows, "Hidden", true);

        var expected = new Dictionary<int, double>();
        for (var row = 3; row <= 242; row++)
        {
          object? cell = null;
          try
          {
            cell = GetRequiredProperty(worksheet, "Cells", row, 1);
            var height = Convert.ToDouble(GetRequiredProperty(cell, "Height"), CultureInfo.InvariantCulture);
            if (double.IsFinite(height) && height >= 0) expected[row] = height;
          }
          finally { Release(cell); }
        }
        MeasurePlacementRead($"row-heights/{scenario}", () =>
        {
          var actual = (IReadOnlyDictionary<int, double>)reader.Invoke(null, [worksheet, 3, 242])!;
          CollectionAssert.AreEqual(expected.ToArray(), actual.ToArray(), scenario);
        }, samples: 3);
      }
      finally
      {
        if (rows is not null) SetProperty(rows, "Hidden", false);
        Release(rows);
        Release(area);
      }
    }
  }

  private static void VerifyPlacementAnalysisPerformance(object worksheet, WorkbookIdentity identity)
  {
    object? area = null;
    object? shapes = null;
    object? selection = null;
    try
    {
      area = GetRequiredProperty(worksheet, "Range", "A1:AF722");
      SetProperty(area, "NumberFormat", "0.00");
      SetProperty(area, "RowHeight", 15);
      SetRangeProperty(worksheet, "C:AF", "ColumnWidth", 8.5);
      SetRangeProperty(worksheet, "123:123", "RowHeight", 30);
      SetCellValue(worksheet, 2, 3, "新環境");
      SetCellValue(worksheet, 2, 18, "旧環境");
      SetCellValue(worksheet, 3, 1, 1);
      SetCellValue(worksheet, 3, 2, 1);
      SetCellValue(worksheet, 243, 2, 2);
      SetCellValue(worksheet, 483, 2, 3);
      shapes = GetRequiredProperty(worksheet, "Shapes");
      for (var index = 0; index < 60; index++)
      {
        object? shape = null;
        try
        {
          shape = InvokeMethod(shapes, "AddShape", 1, 1000f, 100f + index * 30, 120f, 20f)!;
          SetProperty(shape, "Name", index % 2 == 0 ? $"EST_IMG_{index:X32}" : $"Benchmark_{index}");
          SetProperty(shape, "AlternativeText", new ManagedShapeMetadata(1, EvidenceSide.Old, new CellReference(5, 19)).Serialize());
        }
        finally { Release(shape); }
      }
      selection = GetRequiredProperty(worksheet, "Cells", 10, 7);
      _ = InvokeMethod(selection, "Select");
      var sheetName = Convert.ToString(GetRequiredProperty(worksheet, "Name"), CultureInfo.InvariantCulture)!;
      var snapshotService = new ExcelSheetSnapshotService();
      var automatic = new ExcelAutomaticPlacementService();
      var images = new[] { new AutomaticPlacementImage("benchmark.png", new ImageDimensions(120, 30)) };
      string? fingerprint = null;
      MeasurePlacementRead("analyze/240-row-case-60-shapes", () =>
      {
        var analysis = automatic.Analyze(identity, sheetName, EvidenceSide.New, images, requestedCaseLabel: "1-1");
        Assert.IsTrue(analysis.Succeeded, analysis.Message);
        Assert.AreEqual(new CellReference(5, 4), analysis.Steps[0].Plan.FocusCell);
        Assert.AreEqual("1-1", analysis.CaseLabel);
        Assert.IsEmpty(analysis.Steps[0].Plan.Insertions);
        fingerprint ??= analysis.SnapshotFingerprint;
        Assert.AreEqual(fingerprint, analysis.SnapshotFingerprint);
      });
      Console.WriteLine($"PERF analysis fingerprint: {fingerprint}");
      MeasurePlacementRead("navigation-snapshot/60-shapes", () =>
      {
        var captured = snapshotService.CaptureForNavigation(identity, sheetName);
        Assert.IsTrue(captured.Succeeded, captured.Message);
        Assert.HasCount(60, captured.Snapshot!.Shapes);
      });
      var shapeReader = typeof(ExcelSheetSnapshotService).GetMethod(
        "ReadShapes", BindingFlags.NonPublic | BindingFlags.Static)!;
      var expectedShapes = ReadIndividualShapeBounds(shapes);
      MeasurePlacementRead("shapes/60", () =>
      {
        var read = (IReadOnlyList<SnapshotShape>)shapeReader.Invoke(null, [worksheet])!;
        CollectionAssert.AreEqual(expectedShapes, read.ToArray());
      });
      VerifyShapeCoordinatesAfterChanges(worksheet, shapes, shapeReader);
    }
    finally { Release(selection); Release(shapes); Release(area); }
  }

  private static void VerifyShapeCoordinatesAfterChanges(object worksheet, object shapes, MethodInfo reader)
  {
    object? application = null;
    object? shape = null;
    var originalStyle = 1;
    try
    {
      application = GetRequiredProperty(worksheet, "Application");
      originalStyle = Convert.ToInt32(GetRequiredProperty(application, "ReferenceStyle"), CultureInfo.InvariantCulture);
      foreach (var style in new[] { 1, -4150 })
      {
        SetProperty(application, "ReferenceStyle", style);
        foreach (var address in new[] { "A1", "AAA1024", "XFD1048576" })
        {
          object? cell = null;
          try
          {
            cell = GetRequiredProperty(worksheet, "Range", address);
            var expected = new CellReference(
              Convert.ToInt32(GetRequiredProperty(cell, "Row"), CultureInfo.InvariantCulture),
              Convert.ToInt32(GetRequiredProperty(cell, "Column"), CultureInfo.InvariantCulture));
            Assert.AreEqual(expected, ExcelSheetSnapshotService.ReadShapeCellReference(cell));
          }
          finally { Release(cell); }
        }
      }
      shape = InvokeMethod(shapes, "Item", 1)!;
      SetProperty(shape, "Left", 25f);
      SetProperty(shape, "Top", 2200f);
      SetProperty(shape, "AlternativeText", "invalid metadata");
      CollectionAssert.AreEqual(ReadIndividualShapeBounds(shapes),
        ((IReadOnlyList<SnapshotShape>)reader.Invoke(null, [worksheet])!).ToArray(),
        "Moving an image and changing its metadata must be visible on the next read.");
    }
    finally
    {
      if (application is not null) SetProperty(application, "ReferenceStyle", originalStyle);
      Release(shape);
      // The scenario supervisor still owns the Application RCW and must be able to Quit it.
      ReleaseOnce(application);
    }
  }

  // Independent oracle: the original separate Row/Column reads, with no address parsing.
  private static SnapshotShape[] ReadIndividualShapeBounds(object shapes)
  {
    var result = new List<SnapshotShape>();
    var count = Convert.ToInt32(GetRequiredProperty(shapes, "Count"), CultureInfo.InvariantCulture);
    for (var index = 1; index <= count; index++)
    {
      object? shape = null;
      object? topLeft = null;
      object? bottomRight = null;
      try
      {
        shape = InvokeMethod(shapes, "Item", index)!;
        topLeft = GetRequiredProperty(shape, "TopLeftCell");
        bottomRight = GetRequiredProperty(shape, "BottomRightCell");
        var name = Convert.ToString(GetRequiredProperty(shape, "Name"), CultureInfo.InvariantCulture)!;
        var metadata = Convert.ToString(GetRequiredProperty(shape, "AlternativeText"), CultureInfo.InvariantCulture);
        var parsed = ManagedShapeMetadata.TryParse(metadata, out var managedMetadata) ? managedMetadata : null;
        result.Add(new SnapshotShape(name,
          Convert.ToInt32(GetRequiredProperty(topLeft, "Row"), CultureInfo.InvariantCulture),
          Convert.ToInt32(GetRequiredProperty(bottomRight, "Row"), CultureInfo.InvariantCulture),
          Convert.ToInt32(GetRequiredProperty(topLeft, "Column"), CultureInfo.InvariantCulture),
          Convert.ToInt32(GetRequiredProperty(bottomRight, "Column"), CultureInfo.InvariantCulture),
          ManagedShapeMetadata.IsManagedName(name) && parsed is not null)
        {
          TopPoints = Convert.ToDouble(GetRequiredProperty(shape, "Top"), CultureInfo.InvariantCulture),
          WidthPoints = Convert.ToDouble(GetRequiredProperty(shape, "Width"), CultureInfo.InvariantCulture),
          HeightPoints = Convert.ToDouble(GetRequiredProperty(shape, "Height"), CultureInfo.InvariantCulture),
          HorizontalOffsetPoints = ManagedShapeMetadata.IsManagedName(name) && parsed is not null
            ? Convert.ToDouble(GetRequiredProperty(shape, "Left"), CultureInfo.InvariantCulture) -
              Convert.ToDouble(GetRequiredProperty(topLeft, "Left"), CultureInfo.InvariantCulture) : 0,
          SourceDimensions = parsed?.SourceDimensions,
        });
      }
      finally { Release(bottomRight); Release(topLeft); Release(shape); }
    }
    return result.ToArray();
  }

  private static void MeasurePlacementRead(string name, Action action, int samples = 5)
  {
    action(); // Warm-up is excluded from the sample set.
    var timings = new List<double>();
    for (var iteration = 0; iteration < samples; iteration++)
    {
      var watch = Stopwatch.StartNew();
      action();
      timings.Add(watch.Elapsed.TotalMilliseconds);
    }
    var ordered = timings.Order().ToArray();
    Console.WriteLine($"PERF {name}: median={ordered[samples / 2]:F2} ms; samples=[{string.Join(", ", timings.Select(value => value.ToString("F2", CultureInfo.InvariantCulture)))}]");
  }
}
