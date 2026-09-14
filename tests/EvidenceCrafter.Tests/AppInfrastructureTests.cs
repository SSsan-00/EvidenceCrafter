using System.Text.Json;
using EvidenceCrafter.App;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class AppInfrastructureTests
{
  [TestMethod]
  public void ImageWorkflow_RejectsReentryUntilFinallyReleasesIt()
  {
    var gate = new ImageWorkflowGate();
    Assert.IsTrue(gate.TryBegin());
    try
    {
      foreach (var phase in new[] { "analysis", "preview", "editor", "placement" })
        Assert.IsFalse(gate.TryBegin(), phase);
    }
    finally { gate.End(); }
    Assert.IsTrue(gate.TryBegin(), "The next capture must be accepted after completion or cancellation.");
    gate.End();
  }

  [TestMethod]
  public void EnterShortcut_RecognizesNativeKeyRepeat()
  {
    Assert.IsFalse(EnterShortcut.IsRepeat(Message.Create(0, 0x100, (nint)13, 1)));
    Assert.IsTrue(EnterShortcut.IsRepeat(Message.Create(0, 0x100, (nint)13, (nint)((1L << 30) | 1))));
  }
  [TestMethod]
  public void NormalizeScreenSelection_SupportsEveryDragDirection()
  {
    var expected = new Rectangle(10, 20, 21, 31);

    Assert.AreEqual(expected, ScreenCaptureDialog.NormalizeSelection(new Point(10, 20), new Point(30, 50)));
    Assert.AreEqual(expected, ScreenCaptureDialog.NormalizeSelection(new Point(30, 50), new Point(10, 20)));
  }

  [TestMethod]
  public void SettingsStore_RoundTripsNormalizedValuesWithoutRuntimeSelection()
  {
    using var directory = new TemporaryDirectory();
    var path = Path.Combine(directory.Path, "settings.json");
    var store = new AppSettingsStore(path);

    store.Save(new EvidenceCrafterSettings
    {
      HorizontalMarginPoints = 200,
      DiagnosticLoggingEnabled = false,
      GlobalShortcutEnabled = false,
      AlwaysOnTop = true,
      DarkMode = true,
      ThemeColorArgb = Color.FromArgb(48, 96, 160).ToArgb(),
      WindowOpacityPercent = 3,
    });

    var loaded = store.Load();
    Assert.AreEqual(72, loaded.HorizontalMarginPoints);
    Assert.IsFalse(loaded.DiagnosticLoggingEnabled);
    Assert.IsFalse(loaded.GlobalShortcutEnabled);
    Assert.IsTrue(loaded.AlwaysOnTop);
    Assert.IsTrue(loaded.DarkMode);
    Assert.AreEqual(100, loaded.EffectiveThemeIntensity);
    Assert.AreEqual(Color.FromArgb(48, 96, 160).ToArgb(), loaded.EffectiveThemeColor.ToArgb());
    Assert.AreEqual(50, loaded.EffectiveWindowOpacityPercent);
    Assert.AreEqual(100, new EvidenceCrafterSettings().EffectiveWindowOpacityPercent);
    var savedJson = File.ReadAllText(path);
    Assert.IsFalse(savedJson.Contains("Workbook", StringComparison.Ordinal));
    Assert.IsFalse(savedJson.Contains("Side", StringComparison.Ordinal));
  }

  [TestMethod]
  public void SettingsStore_PreservesIntermediateThemeIntensity()
  {
    using var directory = new TemporaryDirectory();
    var path = Path.Combine(directory.Path, "settings.json");
    var store = new AppSettingsStore(path);

    store.Save(new EvidenceCrafterSettings { ThemeIntensity = 35 });

    var loaded = store.Load();
    Assert.AreEqual(35, loaded.EffectiveThemeIntensity);
    Assert.IsFalse(loaded.DarkMode);
  }

  [TestMethod]
  public void SettingsStore_InvalidJson_ReturnsDefaults()
  {
    using var directory = new TemporaryDirectory();
    var path = Path.Combine(directory.Path, "settings.json");
    File.WriteAllText(path, "not-json");

    var loaded = new AppSettingsStore(path).Load();

    Assert.AreEqual(EvidenceCrafterSettings.DefaultHorizontalMarginPoints, loaded.HorizontalMarginPoints);
  }

  [TestMethod]
  public void DiagnosticLog_WritesOnlyStructuralFields()
  {
    using var directory = new TemporaryDirectory();
    var log = new DiagnosticLog(directory.Path);
    var exception = new InvalidOperationException("secret cell text and image path");

    log.Write(
      DiagnosticEventKind.MutationResult,
      DiagnosticOutcome.Failed,
      processId: 42,
      itemCount: 3,
      duration: TimeSpan.FromMilliseconds(12),
      exception);

    var text = File.ReadAllText(Path.Combine(directory.Path, "diagnostic.jsonl"));
    using var document = JsonDocument.Parse(text);
    Assert.AreEqual("MutationResult", document.RootElement.GetProperty("event").GetString());
    Assert.AreEqual(42, document.RootElement.GetProperty("processId").GetInt32());
    StringAssert.Contains(text, nameof(InvalidOperationException));
    Assert.IsFalse(text.Contains(exception.Message, StringComparison.Ordinal));
  }

  [TestMethod]
  public void GlobalShortcut_InvalidRegistration_DoesNotCallWindowsOrThrow()
  {
    var succeeded = GlobalShortcutRegistration.TryRegister(
      windowHandle: 0,
      id: 1,
      ShortcutModifiers.Control,
      Keys.V,
      out var registration,
      out var error);

    Assert.IsFalse(succeeded);
    Assert.IsNull(registration);
    Assert.IsNotNull(error);
  }

  private sealed class TemporaryDirectory : IDisposable
  {
    internal TemporaryDirectory()
    {
      Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"EvidenceCrafterTests-{Guid.NewGuid():N}");
      Directory.CreateDirectory(Path);
    }

    internal string Path { get; }

    public void Dispose() => Directory.Delete(Path, recursive: true);
  }
}
