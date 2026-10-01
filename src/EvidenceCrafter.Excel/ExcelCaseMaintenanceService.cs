using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Excel;

/// <summary>Runs conservative maintenance for the Case containing a known managed-image row.</summary>
public sealed class ExcelCaseMaintenanceService
{
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
    return rowMutationService.TrimCaseTail(workbook, worksheetName, caseRow, tailRows, requireBothSides);
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
