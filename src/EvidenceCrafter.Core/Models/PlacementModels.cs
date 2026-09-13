namespace EvidenceCrafter.Core.Models;

public readonly record struct ImageDimensions(double WidthPoints, double HeightPoints);

public sealed record FittedImage(double WidthPoints, double HeightPoints, double Scale);

public enum ContentKind
{
  Cell,
  Shape,
  ManagedImage,
  Merge,
}

public sealed record ContentSpan(
  EvidenceSide? Side,
  int StartRow,
  int EndRow,
  ContentKind Kind)
{
  public bool IsImage => Kind is ContentKind.ManagedImage;
}

public enum PlacementMode
{
  CaseStart,
  Gap,
  Tail,
}

public sealed record RowInsertion(int AtRow, int Count, string Reason);

/// <summary>The worksheet cell that must be revealed after a placement is verified.</summary>
public readonly record struct CellReference(int Row, int Column);

public sealed record PlacementRequest(
  EvidenceCaseLayout Layout,
  EvidenceSide Side,
  ImageDimensions Image,
  double AvailableWidthPoints,
  int ActiveRow,
  bool PreferActiveGap,
  IReadOnlyList<ContentSpan> Contents,
  IReadOnlyDictionary<int, double> RowHeights,
  int ImageGapRows = 2,
  int TailRows = 4,
  double DefaultRowHeightPoints = 15.0,
  double? ScaleOverride = null);

public sealed record PlacementPlan(
  PlacementMode Mode,
  int StartRow,
  int EndRow,
  FittedImage Image,
  IReadOnlyList<RowInsertion> Insertions,
  CellReference FocusCell,
  string Reason);

public sealed record RowSafetyState(
  int Row,
  bool HasValueOrFormula,
  bool HasCommentOrNote,
  bool HasHyperlink,
  bool HasShape,
  bool IntersectsMerge)
{
  public bool CanDelete =>
    !HasValueOrFormula &&
    !HasCommentOrNote &&
    !HasHyperlink &&
    !HasShape &&
    !IntersectsMerge;
}
