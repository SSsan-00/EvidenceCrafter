namespace EvidenceCrafter.Core.Models;

/// <summary>One occupied worksheet cell captured without retaining an Excel COM object.</summary>
public sealed record SnapshotCell(
  int Row,
  int Column,
  bool HasValueOrFormula,
  bool HasCommentOrNote,
  bool HasHyperlink,
  bool IntersectsMerge)
{
  public bool IsOccupied =>
    HasValueOrFormula || HasCommentOrNote || HasHyperlink || IntersectsMerge;
}

/// <summary>One worksheet Shape and the cells intersected by its bounding box.</summary>
public sealed record SnapshotShape(
  string Name,
  int StartRow,
  int EndRow,
  int StartColumn,
  int EndColumn,
  bool IsManagedImage)
{
  public double TopPoints { get; init; }
  public double WidthPoints { get; init; }
  public double HeightPoints { get; init; }
  public double HorizontalOffsetPoints { get; init; }
  public ImageDimensions? SourceDimensions { get; init; }
}

/// <summary>
/// Immutable, COM-free worksheet state used by layout analysis and placement planning.
/// </summary>
public sealed record SheetSnapshot(
  string WorksheetName,
  CellReference ActiveCell,
  int RawUsedFirstRow,
  int RawUsedLastRow,
  int RawUsedFirstColumn,
  int RawUsedLastColumn,
  bool IsReadOnly,
  bool IsProtected,
  SheetLayoutSignals LayoutSignals,
  IReadOnlyList<SnapshotCell> Cells,
  IReadOnlyList<SnapshotShape> Shapes,
  IReadOnlyDictionary<int, double> RowHeights,
  IReadOnlyDictionary<int, double> ColumnWidths)
{
  public IReadOnlyList<string> WorksheetNames { get; init; } = [];
}
