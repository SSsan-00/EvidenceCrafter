using System.Drawing;

namespace EvidenceCrafter.App;

internal sealed record EvidenceCrafterSettings
{
  internal const double DefaultHorizontalMarginPoints = 6;
  internal const int DefaultThemeColorArgb = unchecked((int)0xFF0D1117);
  internal const int DefaultWindowOpacityPercent = 100;
  internal const int MinimumWindowOpacityPercent = 50;
  public double HorizontalMarginPoints { get; init; } = DefaultHorizontalMarginPoints;

  public bool DiagnosticLoggingEnabled { get; init; } = true;

  public bool GlobalShortcutEnabled { get; init; } = true;

  public bool FollowExcelSelection { get; init; } = true;

  public bool AlwaysOnTop { get; init; }

  // Kept nullable so settings files written by older versions can still use DarkMode.
  public int? ThemeIntensity { get; init; }

  // Kept nullable so settings files written before custom colors retain the original dark theme.
  public int? ThemeColorArgb { get; init; }

  // Windows' color dialog exposes up to 16 user-created palette entries.
  public int[] CustomColors { get; init; } = [];

  // Null keeps settings files written before window transparency fully opaque.
  public int? WindowOpacityPercent { get; init; }

  // Legacy compatibility for settings files created before the gradient theme control.
  public bool DarkMode { get; init; }

  public PlacementAdvanceMode AdvanceMode { get; init; } = PlacementAdvanceMode.SameCaseThenNext;

  internal int EffectiveThemeIntensity => Math.Clamp(
    ThemeIntensity ?? (DarkMode ? 100 : 0),
    0,
    100);

  internal Color EffectiveThemeColor => Color.FromArgb(ThemeColorArgb ?? DefaultThemeColorArgb);

  internal int EffectiveWindowOpacityPercent => Math.Clamp(
    WindowOpacityPercent ?? DefaultWindowOpacityPercent,
    MinimumWindowOpacityPercent,
    100);

  internal EvidenceCrafterSettings Normalize() => this with
  {
    HorizontalMarginPoints = double.IsFinite(HorizontalMarginPoints)
      ? Math.Clamp(HorizontalMarginPoints, 0, 72)
      : DefaultHorizontalMarginPoints,
    ThemeIntensity = EffectiveThemeIntensity,
    ThemeColorArgb = EffectiveThemeColor.ToArgb(),
    CustomColors = (CustomColors ?? []).Take(16).ToArray(),
    WindowOpacityPercent = EffectiveWindowOpacityPercent,
    DarkMode = EffectiveThemeIntensity >= 50,
  };
}

internal enum PlacementAdvanceMode
{
  SameCaseThenNext,
  NextCaseSameSide,
}
