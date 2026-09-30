using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Excel;

/// <summary>Runs conservative maintenance for the Case containing a known managed-image row.</summary>
public sealed class ExcelCaseMaintenanceService
{
  private readonly ExcelSheetSnapshotService snapshotService = new();
  private readonly CaseLayoutAnalyzer layoutAnalyzer = new();
  private readonly ExcelRowMutationService rowMutationService;

  public ExcelCaseMaintenanceService(ExcelRowMutationService? rowMutationService = null)
  {
    this.rowMutationService = rowMutationService ?? new ExcelRowMutationService();
  }

  public RowMutationResult TrimCaseTail(
    WorkbookIdentity workbook,
    string worksheetName,
    int caseRow,
    int tailRows = 4) =>
    TrimCaseTail(workbook, worksheetName, caseRow, tailRows, requireBothSides: false);

  public RowMutationResult TrimCompletedCaseTail(
    WorkbookIdentity workbook,
    string worksheetName,
    int caseRow) =>
    TrimCaseTail(workbook, worksheetName, caseRow, tailRows: 2, requireBothSides: true);

  private RowMutationResult TrimCaseTail(
    WorkbookIdentity workbook,
    string worksheetName,
    int caseRow,
    int tailRows,
    bool requireBothSides)
  {
    var captured = snapshotService.Capture(workbook, worksheetName, caseRow, scopeShapes: true);
    if (!captured.Succeeded || captured.Snapshot is null)
    {
      return RowMutationResult.Failed(RowMutationOperation.Delete, worksheetName, captured.Message);
    }

    var signals = captured.Snapshot.LayoutSignals with { ActiveRow = caseRow };
    var analyzed = layoutAnalyzer.Analyze(signals);
    if (!analyzed.IsSafe || analyzed.Layout is null)
    {
      return RowMutationResult.Failed(
        RowMutationOperation.Delete,
        captured.Snapshot.WorksheetName,
        $"Case末尾を安全に解析できないため行整理を行いません: {string.Join(" ", analyzed.Reasons)}");
    }

    if (!analyzed.Layout.CanDeleteTrailingRows || analyzed.Layout.Kind == SideLayoutKind.NewOnly)
    {
      return RowMutationResult.NoChange(RowMutationOperation.Delete, captured.Snapshot.WorksheetName,
        "NewのみのCase、または最終Caseのため自動行整理は行いません。");
    }

    if (requireBothSides && !HasBothSides(captured.Snapshot, analyzed.Layout))
    {
      return RowMutationResult.NoChange(
        RowMutationOperation.Delete,
        captured.Snapshot.WorksheetName,
        "NewとOldの両方に画像がそろっていないため、Case末尾は整理していません。");
    }

    return rowMutationService.DeleteTrailingRowsWithSnapshot(
      workbook,
      captured.Snapshot.WorksheetName,
      analyzed.Layout.StartRow,
      analyzed.Layout.EndRow,
      tailRows);
  }

  internal static bool HasBothSides(SheetSnapshot snapshot, EvidenceCaseLayout layout) =>
    HasManagedImage(snapshot, layout, EvidenceSide.New) &&
    HasManagedImage(snapshot, layout, EvidenceSide.Old);

  internal static bool HasManagedImage(
    SheetSnapshot snapshot,
    EvidenceCaseLayout layout,
    EvidenceSide side)
  {
    if (!layout.SupportsSide(side)) return false;
    var region = layout.RegionFor(side);
    return snapshot.Shapes.Any(shape =>
      shape.IsManagedImage &&
      shape.StartRow <= layout.EndRow &&
      shape.EndRow >= layout.StartRow + 1 &&
      shape.StartColumn >= region.FirstColumn + 1 &&
      shape.StartColumn <= region.LastColumn);
  }
}
