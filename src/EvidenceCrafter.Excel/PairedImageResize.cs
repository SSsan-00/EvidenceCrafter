using EvidenceCrafter.Core.Models;

namespace EvidenceCrafter.Excel;

public sealed record PairedImagePlan(
  string ShapeName,
  double Scale,
  double Width,
  double Height,
  int StartRow,
  int EndRow,
  bool Legacy)
{
  public double? ReferenceScale { get; init; }
  public int TargetStartRow { get; init; } = StartRow;
  public double? TargetTopPoints { get; init; }
}

/// <summary>The reference resize and its reserved rows form one reversible part of automatic placement.</summary>
public sealed class PairedImageResize(ManagedShapeTarget before, ManagedShapeTarget after, AppliedRowInsertion? insertion)
{
  public ManagedShapeTarget Before { get; } = before;
  public ManagedShapeTarget After { get; private set; } = after;
  public AppliedRowInsertion? Insertion { get; } = insertion;
  internal bool CanRetryPreparation { get; private set; }
  internal bool CompensationSucceeded { get; private set; } = true;
  private readonly ExcelManagedShapeService shapes = new();
  private readonly ExcelRowMutationService rows = new();

  public bool Matches(WorkbookIdentity workbook, bool applied)
  {
    var expected = applied ? After : Before;
    var current = shapes.Inspect(workbook, expected.WorksheetName, expected.ShapeName);
    return current.Shape is { } target && ExcelManagedShapeService.TargetUnchanged(expected, target);
  }

  public RowMutationResult SetApplied(WorkbookIdentity workbook, bool apply)
  {
    CanRetryPreparation = false;
    CompensationSucceeded = true;
    RowMutationResult Failed(string message) => RowMutationResult.Failed(RowMutationOperation.Insert, Before.WorksheetName, message);
    if (!Matches(workbook, !apply))
    {
      CanRetryPreparation = apply;
      return Failed("参照画像が変更されたためUndo/Redoを停止しました。");
    }
    if (apply && Insertion is { } added)
    {
      var insert = rows.InsertRows(workbook, added.WorksheetName, new(added.StartRow, added.Count, added.Reason));
      if (!insert.Succeeded || !insert.Changed) return Failed(insert.Message);
      var normalize = rows.NormalizeInsertedRows(workbook, added.WorksheetName, added.StartRow, added.Count, 15);
      if (!normalize.Succeeded)
      {
        var rollback = rows.DeleteRowsIfSafe(workbook, added.WorksheetName, added.StartRow, added.Count);
        CompensationSucceeded = rollback.Succeeded && rollback.Changed;
        return Failed(normalize.Message + (CompensationSucceeded ? "" : " 行の復元にも失敗しました。"));
      }
    }
    var resize = shapes.Resize(workbook, apply ? Before : After, apply ? After : Before);
    if (!resize.Succeeded)
    {
      CompensationSucceeded = resize.CompensationSucceeded;
      if (apply && Insertion is { } addedRows)
      {
        var rollback = rows.DeleteRowsIfSafe(workbook, addedRows.WorksheetName, addedRows.StartRow, addedRows.Count);
        CompensationSucceeded &= rollback.Succeeded && rollback.Changed;
        if (!rollback.Succeeded || !rollback.Changed) return Failed(resize.Message + " 追加行を復元できませんでした。");
      }
      CanRetryPreparation = apply && resize.ReferenceChanged && CompensationSucceeded;
      return Failed(resize.Message);
    }
    if (apply) After = resize.After!;
    if (!apply && Insertion is { } removed)
    {
      var deletion = rows.DeleteRowsIfSafe(workbook, removed.WorksheetName, removed.StartRow, removed.Count);
      if (!deletion.Succeeded || !deletion.Changed)
      {
        var restored = shapes.Resize(workbook, resize.After!, After);
        CompensationSucceeded = restored.Succeeded;
        return Failed(deletion.Message + (restored.Succeeded ? "" : " 参照画像の復元にも失敗しました。"));
      }
    }
    return RowMutationResult.NoChange(RowMutationOperation.Insert, Before.WorksheetName, "参照画像の倍率を更新しました。");
  }
}
