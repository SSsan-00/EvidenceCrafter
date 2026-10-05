using System.Text;
using EvidenceCrafter.Core.Models;

namespace EvidenceCrafter.Core.Services;

/// <summary>Shares the numbering contract between snapshot capture, navigation, and layout analysis.</summary>
public static class CaseAnchorNormalizer
{
  public static CaseAnchorSignal[] ConfirmedAnchors(SheetLayoutSignals signals) => Normalize(signals.Anchors);

  public static CaseAnchorSignal[] Normalize(IEnumerable<CaseAnchorSignal> anchors)
  {
    ArgumentNullException.ThrowIfNull(anchors);
    var confirmed = new List<CaseAnchorSignal>();
    string? inheritedMajor = null;
    foreach (var anchor in anchors.OrderBy(anchor => anchor.Row))
    {
      var major = string.IsNullOrWhiteSpace(anchor.ColumnAValue)
        ? inheritedMajor
        : NormalizeNumber(anchor.ColumnAValue);
      var minor = NormalizeNumber(anchor.ColumnBValue);
      if (major is not null) inheritedMajor = major;
      if (major is null || minor is null)
      {
        continue;
      }

      confirmed.Add(anchor with { ColumnAValue = major, ColumnBValue = minor });
    }

    return confirmed.ToArray();
  }

  public static string? NormalizeCaseLabel(string label)
  {
    ArgumentNullException.ThrowIfNull(label);
    var normalized = label.Normalize(NormalizationForm.FormKC).Trim()
      .Replace('ー', '-').Replace('―', '-').Replace('‐', '-')
      .Replace('‑', '-').Replace('–', '-').Replace('—', '-').Replace('−', '-');
    var parts = normalized.Split('-', StringSplitOptions.TrimEntries);
    if (parts.Length != 2)
    {
      return null;
    }

    var major = NormalizeNumber(parts[0]);
    var minor = NormalizeNumber(parts[1]);
    return major is null || minor is null ? null : $"{major}-{minor}";
  }

  private static string? NormalizeNumber(string? value)
  {
    var number = value?.Normalize(NormalizationForm.FormKC).Trim();
    return !string.IsNullOrEmpty(number) && number.All(char.IsAsciiDigit) ? number : null;
  }
}
