namespace EvidenceCrafter.Core.Services;

/// <summary>Compares placement geometry rounded to 0.1pt, allowing a 1pt difference.</summary>
public static class PlacementGeometryComparison
{
  public const double TolerancePoints = 1;
  public const double PrecisionPoints = 0.1;

  public static double Round(double value) => Math.Round(value, 1, MidpointRounding.AwayFromZero);

  public static bool Matches(double expected, double actual) =>
    double.IsFinite(expected) && double.IsFinite(actual) &&
    Math.Abs(Round(Round(actual) - Round(expected))) <= TolerancePoints;

  public static bool Exceeds(double actual, double limit) =>
    !double.IsFinite(actual) || !double.IsFinite(limit) ||
    Round(Round(actual) - Round(limit)) > TolerancePoints;
}
