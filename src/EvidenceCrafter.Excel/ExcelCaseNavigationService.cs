using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Excel;

/// <summary>Moves Excel to a structurally confirmed neighbouring Case anchor.</summary>
public sealed class ExcelCaseNavigationService
{
  private readonly ExcelSheetSnapshotService snapshotService = new();
  private readonly IPlacementFocusService focusService = new ExcelPlacementFocusService();
  private readonly CaseLayoutAnalyzer layoutAnalyzer = new();

  public CaseNavigationResult Navigate(
    WorkbookIdentity workbook,
    string worksheetName,
    CaseNavigationDirection direction,
    string? currentCaseLabel = null,
    EvidenceSide currentSide = EvidenceSide.New,
    bool sameCaseThenNext = true)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return CaseNavigationResult.Failed("Case移動はSTAスレッドで実行する必要があります。");
    }
    if (!string.IsNullOrWhiteSpace(currentCaseLabel) &&
      ExcelAutomaticPlacementService.NormalizeCaseLabel(currentCaseLabel) is null)
    {
      return CaseNavigationResult.Failed("CaseはX-X形式で入力してください（例: 1-2）。");
    }

    var captured = snapshotService.CaptureForNavigation(workbook, worksheetName, includeWorksheetNames: true);
    if (!captured.Succeeded || captured.Snapshot is null)
    {
      return CaseNavigationResult.Failed(captured.Message);
    }

    var snapshot = captured.Snapshot;
    var blocks = AvailableBlocks(snapshot, currentSide, sameCaseThenNext);
    if (blocks.Length == 0)
    {
      return CaseNavigationResult.Failed("Case番号・新／旧ヘッダーから配置先を解析できません。");
    }
    if (blocks[0].Layout.Kind == SideLayoutKind.NewOnly) currentSide = EvidenceSide.New;
    if (!string.IsNullOrWhiteSpace(currentCaseLabel) && !blocks.Any(block =>
      ExcelAutomaticPlacementService.FormatCaseLabel(block.Anchor) == ExcelAutomaticPlacementService.NormalizeCaseLabel(currentCaseLabel)))
    {
      return CaseNavigationResult.Failed("指定Caseが見つかりません。配置先を確認してください。");
    }
    var currentIndex = FindCurrentIndex(blocks, snapshot, currentCaseLabel, currentSide);
    var step = direction is CaseNavigationDirection.Previous ? -1 : 1;
    var targetIndex = currentIndex + step;
    Block? block = targetIndex >= 0 && targetIndex < blocks.Length ? blocks[targetIndex] : null;
    if (block is null)
    {
      return NavigateAdjacentWorksheet(workbook, snapshot, currentSide, sameCaseThenNext, direction);
    }

    var layout = block.Layout;
    var firstColumn = layout.RegionFor(block.Side).FirstColumn;
    var target = FocusCell(layout, firstColumn, direction);
    var focused = focusService.FocusPlacedImage(workbook, snapshot.WorksheetName, target);
    var caseLabel = ExcelAutomaticPlacementService.FormatCaseLabel(block.Anchor);
    return focused.Succeeded
      ? new CaseNavigationResult(true, snapshot.WorksheetName, caseLabel, block.Side, target,
        $"{snapshot.WorksheetName} / Case {caseLabel} / {block.Side} へ移動しました。")
        { LayoutSignals = snapshot.LayoutSignals, LayoutKind = layout.Kind }
      : CaseNavigationResult.Failed(focused.Message);
  }

  private CaseNavigationResult NavigateAdjacentWorksheet(
    WorkbookIdentity workbook,
    SheetSnapshot current,
    EvidenceSide currentSide,
    bool sameCaseThenNext,
    CaseNavigationDirection direction)
  {
    if (!current.WorksheetNames.Contains(current.WorksheetName, StringComparer.OrdinalIgnoreCase))
      return CaseNavigationResult.Failed("対象ブックのシート一覧を確認できません。更新して再度移動してください。");
    var candidates = AdjacentWorksheetNames(current.WorksheetNames, current.WorksheetName, direction);
    var failures = new List<string>();
    foreach (var candidate in candidates)
    {
      var activated = focusService.FocusPlacedImage(workbook, candidate, new CellReference(1, 1));
      if (!activated.Succeeded)
      {
        failures.Add($"{candidate}: {activated.Message}");
        continue;
      }

      var captured = snapshotService.CaptureForNavigation(workbook, candidate);
      if (!captured.Succeeded || captured.Snapshot is null)
      {
        failures.Add($"{candidate}: {captured.Message}");
        continue;
      }

      var snapshot = captured.Snapshot;
      var blocks = AvailableBlocks(snapshot, currentSide, sameCaseThenNext);
      if (blocks.Length == 0)
        failures.Add($"{candidate}: CASE／配置先を解析できません。");
      else
      {
        var block = direction == CaseNavigationDirection.Previous ? blocks[^1] : blocks[0];
        var layout = block.Layout;
        var region = layout.RegionFor(block.Side);
        var target = FocusCell(layout, region.FirstColumn, direction);
        var focused = focusService.FocusPlacedImage(workbook, candidate, target);
        if (focused.Succeeded)
        {
          var caseLabel = ExcelAutomaticPlacementService.FormatCaseLabel(block.Anchor);
          return new CaseNavigationResult(true, candidate, caseLabel, block.Side, target,
            $"{candidate} / Case {caseLabel} / {block.Side} へ移動しました。")
            { LayoutSignals = snapshot.LayoutSignals, LayoutKind = layout.Kind };
        }
        failures.Add($"{candidate}: {focused.Message}");
      }
    }

    var edgeBlocks = AvailableBlocks(current, currentSide, sameCaseThenNext);
    if (edgeBlocks.Length > 0)
    {
      var edgeBlock = direction == CaseNavigationDirection.Next
        ? edgeBlocks[^1]
        : edgeBlocks[0];
      var edgeRegion = edgeBlock.Layout.RegionFor(edgeBlock.Side);
      var edgeCell = FocusCell(edgeBlock.Layout, edgeRegion.FirstColumn, direction);
      var edgeFocused = focusService.FocusPlacedImage(workbook, current.WorksheetName, edgeCell);
      if (edgeFocused.Succeeded)
      {
        var edgeCase = ExcelAutomaticPlacementService.FormatCaseLabel(edgeBlock.Anchor);
        return new CaseNavigationResult(false, current.WorksheetName, edgeCase, edgeBlock.Side, edgeCell,
          failures.Count > 0
            ? $"隣接シートへ移動できません。{string.Join(" / ", failures)}"
            : $"これ以上{(direction == CaseNavigationDirection.Next ? "次" : "前")}のCASEはありません。{edgeCase} の端にフォーカスしました。")
          { LayoutSignals = current.LayoutSignals, LayoutKind = edgeBlock.Layout.Kind };
      }
      failures.Add($"CASEの端へのフォーカス: {edgeFocused.Message}");
    }
    var label = direction == CaseNavigationDirection.Previous ? "前" : "次";
    return CaseNavigationResult.Failed(failures.Count > 0
      ? $"{label}のCASE検索で確認できないシートがあります。{string.Join(" / ", failures)}"
      : $"これ以上{label}のCASEはありません（隣接シートも確認済み）。");
  }

  private static CellReference FocusCell(
    EvidenceCaseLayout layout,
    int firstColumn,
    CaseNavigationDirection direction) =>
    new(direction == CaseNavigationDirection.Next ? layout.StartRow + 1 : layout.EndRow, firstColumn + 1);

  internal static IReadOnlyList<string> AdjacentWorksheetNames(
    IReadOnlyList<string> names, string current, CaseNavigationDirection direction)
  {
    var index = names.ToList().FindIndex(name => string.Equals(name, current, StringComparison.OrdinalIgnoreCase));
    if (index < 0) return [];
    return direction == CaseNavigationDirection.Previous
      ? names.Take(index).Reverse().Take(1).ToArray()
      : names.Skip(index + 1).Take(1).ToArray();
  }

  private static int FindCurrentIndex(
    IReadOnlyList<Block> blocks,
    SheetSnapshot snapshot,
    string? currentCaseLabel,
    EvidenceSide currentSide)
  {
    var normalizedCaseLabel = string.IsNullOrWhiteSpace(currentCaseLabel)
      ? null
      : ExcelAutomaticPlacementService.NormalizeCaseLabel(currentCaseLabel);
    var index = normalizedCaseLabel is not null
      ? Array.FindIndex(blocks.ToArray(), block =>
        string.Equals(
          ExcelAutomaticPlacementService.FormatCaseLabel(block.Anchor),
          normalizedCaseLabel,
          StringComparison.OrdinalIgnoreCase) &&
        block.Side == currentSide)
      : -1;
    if (index >= 0)
    {
      return index;
    }

    index = Array.FindLastIndex(blocks.ToArray(), block => block.Anchor.Row <= snapshot.ActiveCell.Row && block.Side == currentSide);
    return index >= 0 ? index : 0;
  }

  internal static bool IsOccupied(
    SheetSnapshot snapshot,
    EvidenceCaseLayout layout,
    EvidenceSide side)
  {
    var region = layout.RegionFor(side);
    return snapshot.Shapes.Any(shape =>
      shape.StartRow <= layout.EndRow && shape.EndRow >= layout.StartRow + 1 &&
      shape.StartColumn <= region.LastColumn && shape.EndColumn >= region.FirstColumn + 1);
  }

  private Block[] AvailableBlocks(SheetSnapshot snapshot, EvidenceSide side, bool sameCaseThenNext)
  {
    var result = new List<Block>();
    foreach (var anchor in ExcelAutomaticPlacementService.ConfirmedAnchors(snapshot.LayoutSignals))
    {
      var layout = layoutAnalyzer.Analyze(snapshot.LayoutSignals with { ActiveRow = anchor.Row }).Layout;
      if (layout is null) continue;
      if (sameCaseThenNext || layout.Kind == SideLayoutKind.NewOnly)
      {
        result.Add(new Block(anchor, EvidenceSide.New, layout));
        if (sameCaseThenNext && layout.SupportsSide(EvidenceSide.Old))
          result.Add(new Block(anchor, EvidenceSide.Old, layout));
      }
      else result.Add(new Block(anchor, side, layout));
    }
    return result.ToArray();
  }

  private sealed record Block(CaseAnchorSignal Anchor, EvidenceSide Side, EvidenceCaseLayout Layout);

  public static int? ResolveTargetRow(
    SheetLayoutSignals signals,
    CaseNavigationDirection direction)
  {
    ArgumentNullException.ThrowIfNull(signals);
    var confirmed = ExcelAutomaticPlacementService.ConfirmedAnchors(signals)
      .Select(anchor => anchor.Row)
      .Distinct()
      .Order()
      .ToArray();
    var currentIndex = Array.FindLastIndex(confirmed, row => row <= signals.ActiveRow);
    var targetIndex = direction is CaseNavigationDirection.Previous ? currentIndex - 1 : currentIndex + 1;
    return targetIndex >= 0 && targetIndex < confirmed.Length ? confirmed[targetIndex] : null;
  }
}

public enum CaseNavigationDirection
{
  Previous,
  Next,
}

public sealed record CaseNavigationResult(
  bool Succeeded,
  string WorksheetName,
  string CaseLabel,
  EvidenceSide Side,
  CellReference Target,
  string Message)
{
  public SideLayoutKind LayoutKind { get; init; }
  public SheetLayoutSignals? LayoutSignals { get; init; }
  public static CaseNavigationResult Failed(string message) =>
    new(false, string.Empty, string.Empty, EvidenceSide.New, default, message);
}
