using System.Collections;
using System.Diagnostics;
using System.Reflection;
using EvidenceCrafter.App;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  private static void VerifyAutomaticPlacementAdvance(object sheet, WorkbookIdentity identity, string imagePath)
  {
    SetRangeProperty(sheet, "A1:AF2548", "NumberFormat", "0.00");
    SetRangeProperty(sheet, "A1:AF2499", "RowHeight", 375);
    SetRangeProperty(sheet, "A2500:AF2500", "RowHeight", 178.5);
    SetRangeProperty(sheet, "A2501:AF2501", "RowHeight", 177.75);
    SetRangeProperty(sheet, "A2502:AF2548", "RowHeight", 18.75);
    SetCellValue(sheet, 2499, 3, "NEW"); SetCellValue(sheet, 2499, 18, "OLD");
    for (var index = 0; index < 3; index++)
    {
      SetCellValue(sheet, 2500 + index * 16, 1, 1);
      SetCellValue(sheet, 2500 + index * 16, 2, index + 1);
    }
    var workbook = GetRequiredProperty(sheet, "Parent");
    var sheets = GetRequiredProperty(workbook, "Worksheets");
    var dependencySheet = InvokeMethod(sheets, "Add")!;
    var pageSetup = GetRequiredProperty(sheet, "PageSetup");
    var directory = Path.Combine(Path.GetDirectoryName(imagePath)!, "advance-ui");
    Directory.CreateDirectory(directory);
    var previousContext = SynchronizationContext.Current;
    var form = new MainForm(new ExcelSessionCatalog());
    var formType = typeof(MainForm);
    object Field(string name) => formType.GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
    void SetField(string name, object value) => formType.GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(form, value);
    try
    {
      SetProperty(pageSetup, "PrintArea", "$A$2500:$AF$2548");
      SetField("settingsStore", new AppSettingsStore(Path.Combine(directory, "settings.json")));
      SetField("diagnosticLog", new DiagnosticLog(directory));
      SetField("settings", new EvidenceCrafterSettings
      {
        GlobalShortcutEnabled = false, FollowExcelSelection = false,
        AdvanceMode = PlacementAdvanceMode.NextCaseSameSide,
      });
      _ = form.Handle;
      SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
      using var image = new Bitmap(120, 80);
      image.Save(imagePath, System.Drawing.Imaging.ImageFormat.Png);
      var automatic = new ExcelAutomaticPlacementService();
      var place = formType.GetMethod("PlaceClipboardImageAutomaticallyAsync", BindingFlags.Instance | BindingFlags.NonPublic)!;
      var snapshots = new ExcelSheetSnapshotService();
      void Complete(Task task, string stage = "placement")
      {
        Console.WriteLine($"UI {stage} started");
        var clock = Stopwatch.StartNew();
        while (!task.IsCompleted && clock.Elapsed < TimeSpan.FromSeconds(65))
        {
          Application.DoEvents(); Thread.Sleep(10);
        }
        Assert.IsTrue(task.IsCompleted, $"UI {stage} did not finish: {((Label)Field("statusLabel")).Text}");
        task.GetAwaiter().GetResult();
      }
      foreach (var label in new[] { "1-1", "1-2" })
      {
        var first = automatic.PlaceImages(identity, "OtherTarget", EvidenceSide.New,
          [new AutomaticPlacementImage(imagePath, new ImageDimensions(120, 80))], requestedCaseLabel: label);
        Assert.IsTrue(first.Succeeded, first.Message);
        if (label == "1-1") Assert.AreEqual(937483.3125, first.PlacedImages[0].Target.TopPoints);
        if (label == "1-2")
        {
          var formula = GetRequiredProperty(dependencySheet, "Cells", 1, 1);
          try { SetProperty(formula, "Formula", "=1+1"); }
          finally { Release(formula); }
        }
        var before = snapshots.Capture(identity, "OtherTarget").Snapshot!;
        var task = (Task)place.Invoke(form, [identity, "OtherTarget", EvidenceSide.Old,
          image, label, null, null, false])!;
        Complete(task);
        Assert.AreEqual(label == "1-1" ? "1-2" : "1-3", ((ComboBox)Field("caseLabelBox")).Text,
          ((Label)Field("statusLabel")).Text);
        Assert.IsTrue(((RadioButton)Field("oldSideButton")).Checked);
        Assert.HasCount(label == "1-1" ? 1 : 2, (ICollection)Field("undoHistory"));
        Assert.IsNull(formType.GetField("rowRecoveryMessage", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form));
        var after = snapshots.Capture(identity, "OtherTarget").Snapshot!;
        Assert.HasCount(before.Shapes.Count + 1, after.Shapes);
        var reference = after.Shapes.Single(shape => shape.Name == first.PlacedImages[0].ShapeName);
        var added = after.Shapes.Single(shape => before.Shapes.All(old => old.Name != shape.Name));
        Assert.AreEqual(reference.TopPoints, added.TopPoints, 0.05);
        if (label == "1-1")
        {
          Complete((Task)formType.GetMethod("UndoAsync", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(form, null)!, "Undo");
          Assert.HasCount(before.Shapes.Count, snapshots.Capture(identity, "OtherTarget").Snapshot!.Shapes);
          Complete((Task)formType.GetMethod("RedoAsync", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(form, null)!, "Redo");
          var restored = snapshots.Capture(identity, "OtherTarget").Snapshot!;
          Assert.HasCount(after.Shapes.Count, restored.Shapes);
          Assert.AreEqual(reference.TopPoints, restored.Shapes.Single(shape => shape.Name != reference.Name).TopPoints, 0.05);
        }
        if (label == "1-2")
        {
          Assert.IsTrue(((Label)Field("statusLabel")).Text.Contains("行未変更", StringComparison.Ordinal));
          CollectionAssert.AreEqual(before.LayoutSignals.Anchors.ToArray(), after.LayoutSignals.Anchors.ToArray());
        }
        Console.WriteLine($"Automatic UI: {label} -> {((ComboBox)Field("caseLabelBox")).Text}: {((Label)Field("statusLabel")).Text}");
      }
      // The UI must still block subsequent mutations when a row result is uncertain.
      var uncertain = new RowMutationResult(false, false, RowMutationOperation.Delete,
        "OtherTarget", 10, 2, "Injected uncertain row mutation", null) { MutationMayHaveOccurred = true };
      var stop = formType.GetMethod("StopAfterRowFailure", BindingFlags.Instance | BindingFlags.NonPublic)!;
      Assert.IsTrue((bool)stop.Invoke(form, [uncertain, "Recovery check"] )!);
      var shapeCount = snapshots.Capture(identity, "OtherTarget").Snapshot!.Shapes.Count;
      ((Task)place.Invoke(form, [identity, "OtherTarget", EvidenceSide.Old,
        image, "1-3", null, null, false])!).GetAwaiter().GetResult();
      Assert.HasCount(shapeCount, snapshots.Capture(identity, "OtherTarget").Snapshot!.Shapes);
      Assert.AreEqual("1-3", ((ComboBox)Field("caseLabelBox")).Text);
    }
    finally
    {
      form.Dispose();
      SynchronizationContext.SetSynchronizationContext(previousContext);
      _ = InvokeMethod(dependencySheet, "Delete");
      Release(dependencySheet); Release(pageSetup); Release(sheets); ReleaseOnce(workbook);
      foreach (var file in Directory.GetFiles(directory)) File.Delete(file);
      Directory.Delete(directory);
    }
  }
}
