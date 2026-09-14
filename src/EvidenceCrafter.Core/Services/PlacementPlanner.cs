using EvidenceCrafter.Core.Models;

namespace EvidenceCrafter.Core.Services;

/// <summary>Creates a side-effect-free plan for image placement and required row insertions.</summary>
public sealed class PlacementPlanner(ImageSizingService imageSizingService)
{
  public const double VerticalInsetPoints = 2.0;
  private const int MinimumImageGapRows = 2;
  private const int MinimumTailRows = 4;
  // Leave two rows below the case header before placing the image.
  private const int PlacementRowInset = 2;
  private const int PlacementColumnInset = 1;
  private const int MaximumWorksheetRow = ExcelWorksheetLimits.MaximumRow;

  public PlacementPlan Plan(PlacementRequest request)
  {
    ArgumentNullException.ThrowIfNull(request);
    ArgumentNullException.ThrowIfNull(request.Layout);
    ArgumentNullException.ThrowIfNull(request.Contents);
    ArgumentNullException.ThrowIfNull(request.RowHeights);

    if (request.ImageGapRows < MinimumImageGapRows ||
      request.ImageGapRows > MaximumWorksheetRow ||
      request.TailRows < MinimumTailRows ||
      request.TailRows > MaximumWorksheetRow)
    {
      throw new ArgumentOutOfRangeException(
        nameof(request),
        $"Placement requires at least {MinimumImageGapRows} image-gap rows and {MinimumTailRows} tail rows.");
    }

    if (!double.IsFinite(request.DefaultRowHeightPoints) || request.DefaultRowHeightPoints <= 0)
    {
      throw new ArgumentOutOfRangeException(nameof(request), "The default row height must be finite and positive.");
    }

    ValidateLayoutAndContents(request);
    var fittedImage = request.ScaleOverride is { } scale
      ? imageSizingService.AtScale(request.Image, request.AvailableWidthPoints, scale)
      : imageSizingService.FitToWidth(request.Image, request.AvailableWidthPoints);
    var caseContents = request.Contents
      .Where(content =>
        content.EndRow >= request.Layout.StartRow &&
        content.StartRow <= request.Layout.EndRow)
      .ToArray();
    var relevantContents = caseContents
      .Where(content =>
        content.Side is null || content.Side == request.Side)
      .OrderBy(content => content.StartRow)
      .ThenBy(content => content.EndRow)
      .ToArray();

    var firstPlacementRow = checked(request.Layout.StartRow + PlacementRowInset);
    var activeRowIsFree = request.ActiveRow >= firstPlacementRow &&
      request.ActiveRow <= request.Layout.EndRow &&
      !relevantContents.Any(content => Covers(content, request.ActiveRow)) &&
      HasRequiredGapBefore(request.ActiveRow, relevantContents, request.ImageGapRows);

    var mode = PlacementMode.Tail;
    var reason = "Placed after the last content in the selected side.";
    int startRow;

    if (request.PreferredStartRow is { } preferredStartRow)
    {
      if (preferredStartRow < firstPlacementRow || preferredStartRow > request.Layout.EndRow)
      {
        throw new InvalidOperationException("対応する反対Side画像の開始行がCASE範囲外です。");
      }
      if (relevantContents.Any(content => Covers(content, preferredStartRow)) ||
        !HasRequiredGapBefore(preferredStartRow, relevantContents, request.ImageGapRows))
      {
        throw new InvalidOperationException("対応する反対Side画像と同じ開始行には既存コンテンツがあるため配置できません。");
      }
      startRow = preferredStartRow;
      mode = preferredStartRow == firstPlacementRow ? PlacementMode.CaseStart : PlacementMode.Gap;
      reason = "Aligned with the corresponding image on the opposite side.";
    }
    else if (request.PreferActiveGap && activeRowIsFree)
    {
      startRow = request.ActiveRow;
      mode = PlacementMode.Gap;
      reason = "ActiveCell identifies an unused gap in the selected side.";
    }
    else if (relevantContents.Length == 0)
    {
      startRow = firstPlacementRow;
      mode = PlacementMode.CaseStart;
      reason = "The selected side has no existing content.";
    }
    else
    {
      startRow = Math.Max(firstPlacementRow, ResolveTailStart(relevantContents, request.ImageGapRows));
    }

    if (startRow < 1 || startRow > MaximumWorksheetRow)
    {
      throw new InvalidOperationException("The placement start would exceed the Excel worksheet row limit.");
    }

    var requiredRows = CountRowsForHeight(
      startRow,
      fittedImage.HeightPoints + VerticalInsetPoints,
      request.RowHeights,
      request.DefaultRowHeightPoints);
    var imageEndValue = (long)startRow + requiredRows - 1;
    if (imageEndValue > MaximumWorksheetRow)
    {
      throw new InvalidOperationException("The placement would exceed the Excel worksheet row limit.");
    }

    var imageEndRow = (int)imageEndValue;
    var insertions = new List<RowInsertion>();

    if (request.PreferredStartRow is not null || mode is PlacementMode.Gap)
    {
      var followingContents = relevantContents
        .Where(content => content.StartRow > startRow)
        .ToArray();
      var insertion = ResolveGapInsertion(imageEndRow, followingContents, request.ImageGapRows);

      if (insertion is not null)
      {
        if (request.PreferredStartRow is not null && followingContents.Any(content =>
          !IsImageLike(content) && content.StartRow <= imageEndRow))
          throw new InvalidOperationException("対応画像の配置領域にセル内容があるため配置できません。");
        insertions.Add(insertion);
      }
    }

    var effectiveCaseEnd = (long)request.Layout.EndRow + insertions.Sum(insertion => (long)insertion.Count);
    if (effectiveCaseEnd > MaximumWorksheetRow)
    {
      throw new InvalidOperationException("The row insertion would exceed the Excel worksheet row limit.");
    }

    var lastExistingContentRow = caseContents
      .Select(content => AdjustedEndRow(content, insertions))
      .DefaultIfEmpty((long)request.Layout.StartRow - 1)
      .Max();
    var requiredCaseEnd = Math.Max(imageEndRow, lastExistingContentRow) + request.TailRows;
    if (requiredCaseEnd > MaximumWorksheetRow)
    {
      throw new InvalidOperationException("The placement would exceed the Excel worksheet row limit.");
    }

    if (requiredCaseEnd > effectiveCaseEnd)
    {
      insertions.Add(new RowInsertion(
        (int)effectiveCaseEnd + 1,
        (int)(requiredCaseEnd - effectiveCaseEnd),
        "Preserve the minimum Case tail after the new image."));
    }

    var sideColumns = request.Layout.RegionFor(request.Side);

    return new PlacementPlan(
      mode,
      startRow,
      imageEndRow,
      fittedImage,
      insertions,
      new CellReference(startRow, checked(sideColumns.FirstColumn + PlacementColumnInset)),
      reason);
  }

  private static bool Covers(ContentSpan content, int row) =>
    content.StartRow <= row && content.EndRow >= row;

  private static bool HasRequiredGapBefore(
    int startRow,
    IReadOnlyList<ContentSpan> contents,
    int imageGapRows)
  {
    var previousImageEnd = contents
      .Where(content => IsImageLike(content) && content.EndRow < startRow)
      .Select(content => content.EndRow)
      .DefaultIfEmpty(int.MinValue)
      .Max();
    return previousImageEnd == int.MinValue || startRow > previousImageEnd + imageGapRows;
  }

  private static RowInsertion? ResolveGapInsertion(
    int imageEndRow,
    IReadOnlyList<ContentSpan> followingContents,
    int imageGapRows)
  {
    var conflicts = followingContents
      .Select(content => new
      {
        Content = content,
        RequiredCount = IsImageLike(content)
          ? (long)imageEndRow + imageGapRows - content.StartRow + 1
          : (long)imageEndRow - content.StartRow + 1,
      })
      .Where(item => item.RequiredCount > 0)
      .ToArray();
    if (conflicts.Length == 0)
    {
      return null;
    }

    return new RowInsertion(
      conflicts.Min(item => item.Content.StartRow),
      checked((int)conflicts.Max(item => item.RequiredCount)),
      "Create enough room at the ActiveCell gap without overwriting content or image spacing.");
  }

  private static long AdjustedEndRow(ContentSpan content, IReadOnlyList<RowInsertion> insertions) =>
    (long)content.EndRow + insertions
      .Where(insertion => insertion.AtRow <= content.EndRow)
      .Sum(insertion => (long)insertion.Count);

  private static bool IsImageLike(ContentSpan content) =>
    content.Kind is ContentKind.ManagedImage or ContentKind.Shape;

  private static int ResolveTailStart(IReadOnlyList<ContentSpan> contents, int imageGapRows)
  {
    var lastImage = contents
      .Where(IsImageLike)
      .OrderBy(content => content.EndRow)
      .LastOrDefault();

    if (lastImage is null)
    {
      return CheckedStartRow((long)contents.Max(content => content.EndRow) + imageGapRows + 1);
    }

    var reservedBandEnd = (long)lastImage.EndRow + imageGapRows;
    var contentBelowReservedBand = contents
      .Where(content => content.EndRow > reservedBandEnd)
      .ToArray();

    if (contentBelowReservedBand.Length == 0)
    {
      return CheckedStartRow(reservedBandEnd + 1);
    }

    return CheckedStartRow(
      (long)contentBelowReservedBand.Max(content => content.EndRow) + imageGapRows + 1);
  }

  private static int CheckedStartRow(long row)
  {
    if (row < 1 || row > MaximumWorksheetRow)
    {
      throw new InvalidOperationException("The placement start would exceed the Excel worksheet row limit.");
    }

    return (int)row;
  }

  private static int CountRowsForHeight(
    int startRow,
    double requiredHeight,
    IReadOnlyDictionary<int, double> rowHeights,
    double defaultRowHeight)
  {
    var total = 0.0;
    var row = startRow;
    while (total < requiredHeight)
    {
      var height = rowHeights.GetValueOrDefault(row, defaultRowHeight);
      if (!double.IsFinite(height) || height < 0)
      {
        throw new InvalidOperationException($"Row {row} has an invalid height.");
      }

      total += height;
      if (row >= MaximumWorksheetRow && total < requiredHeight)
      {
        throw new InvalidOperationException("The image height exceeds the Excel worksheet row limit.");
      }

      row++;
    }

    return row - startRow;
  }

  private static void ValidateLayoutAndContents(PlacementRequest request)
  {
    if (request.Layout.StartRow < 1 ||
      request.Layout.EndRow < request.Layout.StartRow ||
      request.Layout.EndRow > MaximumWorksheetRow ||
      request.ActiveRow is < 1 or > MaximumWorksheetRow ||
      request.Layout.NewRegion.FirstColumn < 1 ||
      request.Layout.NewRegion.Count < 2 ||
      request.Layout.NewRegion.LastColumn > ExcelWorksheetLimits.MaximumColumn ||
      (request.Layout.OldRegion is { } old &&
        (old.FirstColumn < 1 || old.Count < 2 || old.LastColumn > ExcelWorksheetLimits.MaximumColumn ||
          (long)request.Layout.NewRegion.LastColumn + 1 != old.FirstColumn)))
    {
      throw new ArgumentException("The Case layout contains invalid worksheet coordinates.", nameof(request));
    }

    _ = request.Layout.RegionFor(request.Side);

    if (request.Contents.Any(content =>
      content.StartRow < request.Layout.StartRow ||
      content.EndRow > request.Layout.EndRow ||
      content.EndRow < content.StartRow))
    {
      throw new ArgumentException("Content spans must be fully contained by the selected Case.", nameof(request));
    }
  }
}
