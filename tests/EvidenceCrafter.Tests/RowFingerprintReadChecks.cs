using System.Reflection;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  private static void VerifyRowFingerprintReads(object sheet)
  {
    var capture = typeof(ExcelRowMutationService).GetMethod("CaptureRowFingerprint", BindingFlags.Static | BindingFlags.NonPublic)!;
    var reads = new Dictionary<string, int>();
    string Compare(string scenario, int count = 2)
    {
      var expected = CaptureLegacyFingerprint(sheet, 7, count, true);
      reads.Clear();
      var previousObserver = ExcelRowMutationService.StageMeasured;
      string actual;
      try
      {
        ExcelRowMutationService.StageMeasured = (stage, _) => reads[stage] = reads.GetValueOrDefault(stage) + 1;
        actual = (string)capture.Invoke(null, [sheet, 7, count, true])!;
      }
      finally { ExcelRowMutationService.StageMeasured = previousObserver; }
      Assert.AreEqual(expected, actual, scenario);
      return actual;
    }

    var range = GetRequiredProperty(sheet, "Range", "Z7:AC8");
    var font = GetRequiredProperty(range, "Font");
    var cell = GetRequiredProperty(sheet, "Range", "AA8");
    var cellFont = GetRequiredProperty(cell, "Font");
    try
    {
      SetProperty(range, "Value2", new object[,] { { "text", 0, false, 12.75 }, { "rich text", "cell", true, "" } });
      SetProperty(range, "HorizontalAlignment", -4131);
      SetProperty(range, "VerticalAlignment", -4108);
      SetProperty(font, "Name", "Calibri");
      SetProperty(font, "Size", 11d);
      SetProperty(font, "Bold", false);
      SetProperty(font, "Italic", false);
      SetProperty(font, "Color", 0);
      var baseline = Compare("uniform formatting");

      foreach (var (property, original, changed) in new (string, object, object)[]
      {
        ("Name", "Calibri", "Arial"), ("Size", 11d, 18d), ("Bold", false, true),
        ("Italic", false, true), ("Color", 0, 255),
      })
      {
        SetProperty(cellFont, property, changed);
        Assert.AreNotEqual(baseline, Compare("mixed font " + property), property);
        SetProperty(cellFont, property, original);
        Assert.AreEqual(baseline, Compare("restored font " + property));
      }

      foreach (var (property, original, changed) in new (string, object, object)[]
      {
        ("HorizontalAlignment", -4131, -4152), ("VerticalAlignment", -4108, -4160),
      })
      {
        SetProperty(cell, property, changed);
        Assert.AreNotEqual(baseline, Compare("mixed " + property), property);
        SetProperty(cell, property, original);
        Assert.AreEqual(baseline, Compare("restored " + property));
      }

      var characters = GetRequiredProperty(cell, "Characters", 1, 1);
      var characterFont = GetRequiredProperty(characters, "Font");
      try
      {
        SetProperty(characterFont, "Bold", true);
        Assert.AreNotEqual(baseline, Compare("mixed font within a cell"));
        SetProperty(characterFont, "Bold", false);
        Assert.AreEqual(baseline, Compare("restored rich text"));
      }
      finally { Release(characterFont); Release(characters); }

      var conditions = GetRequiredProperty(cell, "FormatConditions");
      var validation = GetRequiredProperty(cell, "Validation");
      try
      {
        Release(InvokeMethod(conditions, "Add", 2, Type.Missing, "=1=1"));
        Assert.AreNotEqual(baseline, Compare("conditional formatting on one cell"));
        _ = InvokeMethod(conditions, "Delete");
        Assert.AreEqual(baseline, Compare("removed conditional formatting"));
        _ = InvokeMethod(validation, "Add", 1, 1, 1, "1", "10");
        Assert.AreNotEqual(baseline, Compare("validation on one cell"));
        _ = InvokeMethod(validation, "Delete");
        Assert.AreEqual(baseline, Compare("removed validation"));
      }
      finally { Release(validation); Release(conditions); }

      SetProperty(range, "NumberFormat", "0.00");
      SetProperty(range, "WrapText", false);
      SetRangeProperty(sheet, "AA7:AA8", "NumberFormat", "@");
      SetRangeProperty(sheet, "AA7:AA8", "WrapText", true);
      SetRangeProperty(sheet, "AA9", "NumberFormat", "0%");
      SetRangeProperty(sheet, "AA9", "WrapText", false);
      var columnUniform = Compare("formats uniform within each column, different outside fingerprint rows");
      Assert.AreEqual(4, reads.GetValueOrDefault("Fingerprint.NumberFormat.Bulk"));
      Assert.AreEqual(4, reads.GetValueOrDefault("Fingerprint.WrapText.Bulk"));
      Assert.AreEqual(0, reads.GetValueOrDefault("Fingerprint.NumberFormat.CellFallback"));
      Assert.AreEqual(0, reads.GetValueOrDefault("Fingerprint.WrapText.CellFallback"));
      SetProperty(cell, "NumberFormat", "0.0");
      SetProperty(cell, "WrapText", false);
      Assert.AreNotEqual(columnUniform, Compare("mixed values within one column"));
      Assert.AreEqual(2, reads.GetValueOrDefault("Fingerprint.NumberFormat.CellFallback"));
      Assert.AreEqual(2, reads.GetValueOrDefault("Fingerprint.WrapText.CellFallback"));
      SetRangeProperty(sheet, "Z7:AC7", "NumberFormat", "0.00");
      SetRangeProperty(sheet, "Z7:AC7", "WrapText", false);
      SetRangeProperty(sheet, "Z8:AC8", "NumberFormat", "@");
      SetRangeProperty(sheet, "Z8:AC8", "WrapText", true);
      _ = Compare("every column has mixed formats");
      Assert.AreEqual(8, reads.GetValueOrDefault("Fingerprint.NumberFormat.CellFallback"));
      Assert.AreEqual(8, reads.GetValueOrDefault("Fingerprint.WrapText.CellFallback"));
      _ = Compare("single row", count: 1);

      var merged = GetRequiredProperty(sheet, "Range", "Z7:AA7");
      try
      {
        _ = InvokeMethod(merged, "Merge");
        _ = Compare("merged cells");
      }
      finally { Release(merged); }
    }
    finally { Release(cellFont); Release(cell); Release(font); Release(range); }
  }
}
