using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Excel;

/// <summary>Creates verified row space before a managed image is allowed to grow.</summary>
public sealed class ExcelManagedReplacementLayoutService
{
  private readonly ExcelSheetSnapshotService snapshotService = new();
  private readonly CaseLayoutAnalyzer layoutAnalyzer = new();
  private readonly ImageSizingService sizingService = new();
  private readonly ExcelRowMutationService rowMutationService;

  public ExcelManagedReplacementLayoutService(ExcelRowMutationService? rowMutationService = null)
  {
    this.rowMutationService = rowMutationService ?? new ExcelRowMutationService();
  }

  public ReplacementLayoutResult EnsureSpace(
    WorkbookIdentity workbook,
    ManagedShapeTarget target,
    ImageDimensions image,
    double horizontalMarginPoints)
  {
    var captured = snapshotService.Capture(workbook, target.WorksheetName, target.TopLeftCell.Row);
    if (!captured.Succeeded || captured.Snapshot is null)
    {
      return ReplacementLayoutResult.Failed(captured.Message);
    }

    var snapshot = captured.Snapshot;
    var analyzed = layoutAnalyzer.Analyze(snapshot.LayoutSignals with { ActiveRow = target.TopLeftCell.Row });
    if (!analyzed.IsSafe || analyzed.Layout is null)
    {
      return ReplacementLayoutResult.Failed($"対象Caseを安全に解析できません: {string.Join(" ", analyzed.Reasons)}");
    }

    var layout = analyzed.Layout;
    if (!layout.SupportsSide(target.Metadata.Side))
    {
      return ReplacementLayoutResult.Failed("対象画像のSideは現在のシート構成では利用できません。");
    }
    var columns = layout.RegionFor(target.Metadata.Side);
    var width = Enumerable.Range(columns.FirstColumn, columns.Count)
      .Sum(column => snapshot.ColumnWidths.GetValueOrDefault(column)) - (horizontalMarginPoints * 2);
    var left = Enumerable.Range(1, columns.FirstColumn - 1)
      .Sum(column => snapshot.ColumnWidths.GetValueOrDefault(column)) + horizontalMarginPoints;
    if (!double.IsFinite(width) || width <= 0)
    {
      return ReplacementLayoutResult.Failed("差し替え先の列幅を取得できません。");
    }

    var fitted = sizingService.FitToWidth(image, width);
    var extraHeight = fitted.HeightPoints - target.HeightPoints;
    if (!PlacementGeometryComparison.Exceeds(extraHeight, 0))
    {
      return new ReplacementLayoutResult(true, fitted, left, null, "追加行は不要です。");
    }

    var averageHeight = snapshot.RowHeights
      .Where(pair => pair.Key >= target.TopLeftCell.Row && pair.Key <= layout.EndRow)
      .Select(pair => pair.Value)
      .Where(value => double.IsFinite(value) && value > 0)
      .DefaultIfEmpty(15)
      .Average();
    var count = Math.Max(1, checked((int)Math.Ceiling(extraHeight / averageHeight)));
    // Insert below the old shape. Inserting through the shape would move/resize it
    // before strict identity validation in Replace and make compensation unsafe.
    var oldShapeEndRow = snapshot.Shapes
      .Where(shape => string.Equals(shape.Name, target.ShapeName, StringComparison.Ordinal))
      .Select(shape => shape.EndRow)
      .DefaultIfEmpty(target.TopLeftCell.Row)
      .Single();
    var insertion = new RowInsertion(
      checked(oldShapeEndRow + 1),
      count,
      "Create room for a taller managed-image replacement.");
    var mutated = rowMutationService.InsertRows(workbook, target.WorksheetName, insertion);
    if (!mutated.Succeeded || !mutated.Changed)
    {
      return ReplacementLayoutResult.Failed(mutated.Message) with { RowFailure = mutated };
    }

    var insertedHeight = snapshotService.Capture(
      workbook,
      target.WorksheetName,
      target.TopLeftCell.Row).Snapshot?.RowHeights
      .Where(pair => pair.Key >= mutated.StartRow && pair.Key < mutated.StartRow + mutated.Count)
      .Sum(pair => pair.Value) ?? 0;
    if (PlacementGeometryComparison.Exceeds(extraHeight, insertedHeight))
    {
      var applied = new AppliedRowInsertion(mutated.WorksheetName, mutated.StartRow, mutated.Count, insertion.Reason);
      var reverted = rowMutationService.DeleteRowsIfSafe(workbook, mutated.WorksheetName, mutated.StartRow, mutated.Count);
      return reverted.Succeeded && reverted.Changed
        ? ReplacementLayoutResult.Failed("追加行の実高が画像に不足したため、差し替えを中止して行を戻しました。")
        : new ReplacementLayoutResult(
          false,
          fitted,
          left,
          applied,
          "追加行の実高が画像に不足し、追加行の復旧にも失敗しました。Workbookを保存せず状態を確認してください。") { RowFailure = reverted };
    }

    return new ReplacementLayoutResult(
      true,
      fitted,
      left,
      new AppliedRowInsertion(mutated.WorksheetName, mutated.StartRow, mutated.Count, insertion.Reason),
      mutated.Message);
  }
}

public sealed record ReplacementLayoutResult(
  bool Succeeded,
  FittedImage? FittedImage,
  double? LeftPoints,
  AppliedRowInsertion? Insertion,
  string Message)
{
  public RowMutationResult? RowFailure { get; init; }
  public static ReplacementLayoutResult Failed(string message) => new(false, null, null, null, message);
}
