using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  // Characterize the live Excel reads against the pre-optimization, cell-by-cell behavior.
  // Use the existing isolated Excel process and its timeout/cleanup supervisor.
  private static void VerifySnapshotReadPerformance(object worksheet)
  {
    var verticalReader = typeof(ExcelSheetSnapshotService).GetMethod(
      "ReadVerticalBoundaries", BindingFlags.NonPublic | BindingFlags.Static)!;
    var widthReader = typeof(ExcelSheetSnapshotService).GetMethod(
      "ReadColumnWidths", BindingFlags.NonPublic | BindingFlags.Static)!;
    object? area = null;
    try
    {
      area = GetRequiredProperty(worksheet, "Range", "A1:AF1500");
      _ = InvokeMethod(area, "Clear");
      SetRangeBorder(worksheet, 3, 17, 1202, 17, 10);
      var timings = new List<double>();
      IReadOnlyList<VerticalBoundarySignal>? boundaries = null;
      for (var iteration = 0; iteration < 4; iteration++)
      {
        var timer = Stopwatch.StartNew();
        boundaries = (IReadOnlyList<VerticalBoundarySignal>)verticalReader.Invoke(
          null, [worksheet, 3, 1500, 32, new[] { 18 }, 1100])!;
        if (iteration > 0) timings.Add(timer.Elapsed.TotalMilliseconds);
        Assert.IsEmpty(boundaries, "Partial borders are decorative and must not trigger a row-by-row scan.");
      }
      timings.Sort();
      Console.WriteLine($"Snapshot vertical boundary, 1200 bordered rows + trailing blank rows: median {timings[1]:F2} ms (3 warmed runs)");

      SetRangeBorder(worksheet, 3, 17, 1500, 17, 10);
      var continuous = (IReadOnlyList<VerticalBoundarySignal>)verticalReader.Invoke(
        null, [worksheet, 3, 1500, 32, new[] { 18 }, 1100])!;
      CollectionAssert.AreEqual(new[] { new VerticalBoundarySignal(17, 3, 1500) }, continuous.ToArray());

      SetRangeProperty(worksheet, "C:AF", "ColumnWidth", 8.5);
      for (var scenario = 0; scenario < 5; scenario++)
      {
        if (scenario == 1) SetRangeProperty(worksheet, "Q:Q", "ColumnWidth", 12.75);
        if (scenario == 2) SetRangeProperty(worksheet, "R:R", "Hidden", true);
        if (scenario == 3) SetRangeProperty(worksheet, "Q:Q", "ColumnWidth", 8.5);
        if (scenario == 4) SetRangeProperty(worksheet, "C:AF", "Hidden", true);
        var expected = new Dictionary<int, double>();
        for (var column = 3; column <= 32; column++)
        {
          object? cell = null;
          try
          {
            cell = GetRequiredProperty(worksheet, "Cells", 1, column);
            var width = Convert.ToDouble(GetRequiredProperty(cell, "Width"), CultureInfo.InvariantCulture);
            if (double.IsFinite(width) && width > 0) expected[column] = width;
          }
          finally { Release(cell); }
        }
        timings.Clear();
        for (var iteration = 0; iteration < 4; iteration++)
        {
          var timer = Stopwatch.StartNew();
          var actual = (IReadOnlyDictionary<int, double>)widthReader.Invoke(null, [worksheet, 3, 32])!;
          if (iteration > 0) timings.Add(timer.Elapsed.TotalMilliseconds);
          CollectionAssert.AreEqual(expected.ToArray(), actual.ToArray(), $"Column scenario {scenario}");
        }
        timings.Sort();
        Console.WriteLine($"Snapshot column widths, scenario {scenario} (0=uniform, 1=mixed, 2=mixed/hidden, 3=uniform/hidden, 4=all hidden): median {timings[1]:F2} ms (3 warmed runs)");
      }
    }
    finally
    {
      if (area is not null) _ = InvokeMethod(area, "Clear");
      Release(area);
      SetRangeProperty(worksheet, "C:AF", "Hidden", false);
      SetRangeProperty(worksheet, "C:AF", "ColumnWidth", 8.43);
    }
  }

  private static void SetRangeProperty(object worksheet, string address, string property, object value)
  {
    object? range = null;
    object? entireRange = null;
    try
    {
      range = GetRequiredProperty(worksheet, "Range", address);
      if (property == "Hidden")
      {
        entireRange = GetRequiredProperty(range, char.IsDigit(address[0]) ? "EntireRow" : "EntireColumn");
        SetProperty(entireRange, property, value);
      }
      else
      {
        SetProperty(range, property, value);
      }
    }
    finally { Release(entireRange); Release(range); }
  }
}
