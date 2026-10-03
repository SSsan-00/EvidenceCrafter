using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.App;

internal enum DiagnosticEventKind
{
  ExcelDiscovery,
  SnapshotAnalysis,
  MutationPlan,
  MutationResult,
  Clipboard,
  Application,
}

internal enum DiagnosticOutcome
{
  Started,
  Succeeded,
  Rejected,
  Failed,
  Finished,
}

internal sealed class DiagnosticLog
{
  private const long MaximumFileBytes = 2 * 1024 * 1024;
  private const int MaximumFiles = 5;
  private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web)
  {
    Converters = { new JsonStringEnumConverter() },
    NumberHandling = JsonNumberHandling.AllowNamedFloatingPointLiterals,
  };
  private readonly string logDirectory;
  private readonly Lock sync = new();

  internal DiagnosticLog(string? logDirectory = null)
  {
    this.logDirectory = logDirectory ?? Path.Combine(
      Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
      "EvidenceCrafter",
      "logs");
  }

  internal void Write(
    DiagnosticEventKind eventKind,
    DiagnosticOutcome outcome,
    int? processId = null,
    int? itemCount = null,
    TimeSpan? duration = null,
    Exception? exception = null,
    uint? clipboardSequence = null,
    ImagePlacementDiagnostic? placement = null)
  {
    var entry = new DiagnosticEntry(
      DateTimeOffset.UtcNow,
      eventKind,
      outcome,
      processId,
      itemCount,
      duration is null ? null : checked((long)duration.Value.TotalMilliseconds),
      exception?.GetType().FullName,
      exception?.HResult,
      clipboardSequence,
      placement,
      typeof(DiagnosticLog).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion);

    try
    {
      lock (sync)
      {
        Directory.CreateDirectory(logDirectory);
        var path = Path.Combine(logDirectory, "diagnostic.jsonl");
        RotateIfNeeded(path);
        File.AppendAllText(path, JsonSerializer.Serialize(entry, JsonOptions) + Environment.NewLine);
      }
    }
    catch (Exception writeFailure) when (writeFailure is IOException or UnauthorizedAccessException)
    {
      // Diagnostics must never interrupt an Excel operation or application shutdown.
    }
  }

  private void RotateIfNeeded(string path)
  {
    if (!File.Exists(path) || new FileInfo(path).Length < MaximumFileBytes)
    {
      return;
    }

    var oldest = $"{path}.{MaximumFiles - 1}";
    if (File.Exists(oldest))
    {
      File.Delete(oldest);
    }

    for (var index = MaximumFiles - 2; index >= 1; index--)
    {
      var source = $"{path}.{index}";
      if (File.Exists(source))
      {
        File.Move(source, $"{path}.{index + 1}");
      }
    }

    File.Move(path, $"{path}.1");
  }

  private sealed record DiagnosticEntry(
    DateTimeOffset TimestampUtc,
    DiagnosticEventKind Event,
    DiagnosticOutcome Outcome,
    int? ProcessId,
    int? ItemCount,
    long? DurationMilliseconds,
    string? ExceptionType,
    int? ExceptionHResult,
    uint? ClipboardSequence,
    ImagePlacementDiagnostic? Placement,
    string? AppVersion);
}
