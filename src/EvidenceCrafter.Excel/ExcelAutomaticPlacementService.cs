using System.Security.Cryptography;
using System.Text;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Excel;

/// <summary>Connects a live worksheet snapshot to Case analysis, placement planning and verified Excel mutations.</summary>
public sealed class ExcelAutomaticPlacementService
{
  private const double InsertedRowHeightPoints = 15;
  private readonly ExcelSheetSnapshotService snapshotService;
  private readonly CaseLayoutAnalyzer layoutAnalyzer;
  private readonly ContentOccupancyAnalyzer occupancyAnalyzer;
  private readonly PlacementPlanner placementPlanner;
  private readonly ExcelRowMutationService rowMutationService;
  private readonly ExcelImagePlacementService imagePlacementService;

  public ExcelAutomaticPlacementService(
    ExcelSheetSnapshotService? snapshotService = null,
    CaseLayoutAnalyzer? layoutAnalyzer = null,
    ContentOccupancyAnalyzer? occupancyAnalyzer = null,
    PlacementPlanner? placementPlanner = null,
    ExcelRowMutationService? rowMutationService = null,
    ExcelImagePlacementService? imagePlacementService = null)
  {
    this.snapshotService = snapshotService ?? new ExcelSheetSnapshotService();
    this.layoutAnalyzer = layoutAnalyzer ?? new CaseLayoutAnalyzer();
    this.occupancyAnalyzer = occupancyAnalyzer ?? new ContentOccupancyAnalyzer();
    this.placementPlanner = placementPlanner ?? new PlacementPlanner(new ImageSizingService());
    this.rowMutationService = rowMutationService ?? new ExcelRowMutationService();
    this.imagePlacementService = imagePlacementService ?? new ExcelImagePlacementService();
  }

  /// <summary>Builds a COM-free plan for every image without changing Excel.</summary>
  public AutomaticPlacementAnalysisResult Analyze(
    WorkbookIdentity workbook,
    string worksheetName,
    EvidenceSide side,
    IReadOnlyList<AutomaticPlacementImage> images,
    bool preferActiveGap = false,
    double horizontalMarginPoints = 6,
    string? requestedCaseLabel = null,
    bool autoDetectSide = false,
    bool sameCaseThenNext = false)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ArgumentNullException.ThrowIfNull(images);
    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return AutomaticPlacementAnalysisResult.Failed("自動配置の解析はSTAスレッドで実行する必要があります。");
    }

    var validation = ValidateImages(images, requireFiles: false);
    if (validation is not null)
    {
      return AutomaticPlacementAnalysisResult.Failed(validation);
    }

    var captured = snapshotService.Capture(workbook, worksheetName);
    if (!captured.Succeeded || captured.Snapshot is null)
    {
      return AutomaticPlacementAnalysisResult.Failed(captured.Message);
    }

    var snapshot = captured.Snapshot;
    var selectedCase = ResolveRequestedCase(snapshot, requestedCaseLabel, side, sameCaseThenNext);
    if (!selectedCase.Succeeded)
    {
      return AutomaticPlacementAnalysisResult.Failed(selectedCase.Message);
    }

    var activeCaseRow = ConfirmedAnchors(snapshot.LayoutSignals)
      .LastOrDefault(anchor => anchor.Row <= snapshot.ActiveCell.Row)?.Row;
    if (selectedCase.Row != activeCaseRow)
    {
      captured = snapshotService.Capture(workbook, snapshot.WorksheetName, selectedCase.Row);
      if (!captured.Succeeded || captured.Snapshot is null)
      {
        return AutomaticPlacementAnalysisResult.Failed(captured.Message);
      }

      snapshot = captured.Snapshot;
    }

    snapshot = snapshot with
    {
      ActiveCell = new CellReference(selectedCase.Row, snapshot.ActiveCell.Column),
      LayoutSignals = snapshot.LayoutSignals with { ActiveRow = selectedCase.Row },
    };
    if (selectedCase.Side is { } fallbackSide)
    {
      side = fallbackSide;
    }
    else if (autoDetectSide)
    {
      var currentLayout = layoutAnalyzer.Analyze(snapshot.LayoutSignals).Layout;
      if (currentLayout is not null)
      {
        side = SideForColumn(snapshot.ActiveCell.Column, currentLayout, side);
      }
    }
    if (snapshot.IsReadOnly || snapshot.IsProtected)
    {
      return AutomaticPlacementAnalysisResult.Failed(
        snapshot.IsReadOnly ? "対象Workbookは読み取り専用です。" : $"シート {snapshot.WorksheetName} は保護されています。");
    }

    return AnalyzeSnapshot(snapshot, side, images, preferActiveGap, horizontalMarginPoints);
  }

  /// <summary>Builds the same plan from an already captured immutable snapshot.</summary>
  public AutomaticPlacementAnalysisResult AnalyzeSnapshot(
    SheetSnapshot snapshot,
    EvidenceSide side,
    IReadOnlyList<AutomaticPlacementImage> images,
    bool preferActiveGap = false,
    double horizontalMarginPoints = 6,
    string? requestedCaseLabel = null,
    bool sameCaseThenNext = false)
  {
    ArgumentNullException.ThrowIfNull(snapshot);
    ArgumentNullException.ThrowIfNull(images);
    var validation = ValidateImages(images, requireFiles: false);
    if (validation is not null)
    {
      return AutomaticPlacementAnalysisResult.Failed(validation);
    }

    if (snapshot.IsReadOnly || snapshot.IsProtected)
    {
      return AutomaticPlacementAnalysisResult.Failed(
        snapshot.IsReadOnly ? "対象Workbookは読み取り専用です。" : $"シート {snapshot.WorksheetName} は保護されています。");
    }

    var selectedCase = ResolveRequestedCase(snapshot, requestedCaseLabel, side, sameCaseThenNext);
    if (!selectedCase.Succeeded)
    {
      return AutomaticPlacementAnalysisResult.Failed(selectedCase.Message);
    }

    snapshot = snapshot with
    {
      ActiveCell = new CellReference(selectedCase.Row, snapshot.ActiveCell.Column),
      LayoutSignals = snapshot.LayoutSignals with { ActiveRow = selectedCase.Row },
    };
    if (selectedCase.Side is { } fallbackSide)
    {
      side = fallbackSide;
    }

    var analyzed = layoutAnalyzer.Analyze(snapshot.LayoutSignals);
    if (!analyzed.IsSafe || analyzed.Layout is null)
    {
      return AutomaticPlacementAnalysisResult.Failed(string.Join(" ", analyzed.Reasons));
    }

    try
    {
      var layout = analyzed.Layout;
      var contents = occupancyAnalyzer.Analyze(snapshot, layout).ToList();
      // Reserve the other side's earlier and later image rows as shared bands.
      // Only the matching ordinal may occupy the same rows as the new image.
      if (images.Count == 1 && layout.Kind == SideLayoutKind.Both)
      {
        var ordinal = CaseImages(snapshot, layout, side).Length;
        var opposite = side == EvidenceSide.New ? EvidenceSide.Old : EvidenceSide.New;
        contents.AddRange(CaseImages(snapshot, layout, opposite)
          .Where((_, index) => index != ordinal)
          .Select(shape => new ContentSpan(null, shape.StartRow, shape.EndRow, ContentKind.ManagedImage)));
      }
      var rowHeights = snapshot.RowHeights.ToDictionary(pair => pair.Key, pair => pair.Value);
      var plans = new List<AutomaticPlacementStep>(images.Count);
      for (var index = 0; index < images.Count; index++)
      {
        var sideColumns = layout.RegionFor(side);
        var width = AvailableWidth(snapshot, sideColumns, horizontalMarginPoints);
        var pair = images.Count == 1
          ? FindPair(snapshot, layout, side, images[index].Dimensions, width, horizontalMarginPoints,
            images[index].PreserveReferenceSize) : null;
        var plan = placementPlanner.Plan(new PlacementRequest(
          layout,
          side,
          images[index].Dimensions,
          width,
          index == 0 ? snapshot.ActiveCell.Row : layout.EndRow,
          index == 0 && preferActiveGap,
          contents,
          rowHeights,
          ScaleOverride: images[index].ScaleOverride ?? pair?.Scale,
          PreferredStartRow: pair?.StartRow));
        plans.Add(new AutomaticPlacementStep(index, images[index], plan, width) { Pair = pair });
        ApplyPlanToModel(plan, side, ref layout, contents, rowHeights);
      }

      return new AutomaticPlacementAnalysisResult(
        true,
        snapshot.WorksheetName,
        analyzed,
        plans,
        Fingerprint(snapshot),
        ResolveCaseLabel(snapshot, analyzed.Layout!.StartRow),
        snapshot.ActiveCell.Row,
        side,
        $"{snapshot.WorksheetName} の{side}側へ{plans.Count}件を配置する計画を作成しました。")
      {
        LayoutSignals = snapshot.LayoutSignals,
        CompletesCaseAfterPlacement = analyzed.Layout.Kind == SideLayoutKind.Both &&
          analyzed.Layout.CanDeleteTrailingRows && ExcelCaseMaintenanceService.HasManagedImage(
          snapshot,
          analyzed.Layout,
          side is EvidenceSide.New ? EvidenceSide.Old : EvidenceSide.New),
      };
    }
    catch (Exception exception) when (exception is ArgumentException or InvalidOperationException or OverflowException)
    {
      if (images.Count == 1 && !images[0].PreserveReferenceSize && images[0].ScaleOverride is null)
      {
        var recovered = AnalyzeSnapshot(snapshot, side,
          [images[0] with { PreserveReferenceSize = true }], preferActiveGap, horizontalMarginPoints,
          requestedCaseLabel, sameCaseThenNext);
        if (recovered.Succeeded) return recovered;
      }
      return AutomaticPlacementAnalysisResult.Failed($"配置計画を作成できません: {exception.Message}");
    }
  }

  /// <summary>Analyzes and applies all images. A failed step compensates prior Shapes and inserted rows in reverse order.</summary>
  public AutomaticPlacementResult PlaceImages(
    WorkbookIdentity workbook,
    string worksheetName,
    EvidenceSide side,
    IReadOnlyList<AutomaticPlacementImage> images,
    bool preferActiveGap = false,
    double horizontalMarginPoints = 6,
    AutomaticPlacementAnalysisResult? preparedAnalysis = null,
    string? requestedCaseLabel = null,
    bool referencePrepared = false)
  {
    var validation = ValidateImages(images, requireFiles: true);
    if (validation is not null)
    {
      return AutomaticPlacementResult.Failed(validation);
    }

    var initialAnalysis = preparedAnalysis ?? Analyze(
      workbook,
      worksheetName,
      side,
      images,
      preferActiveGap,
      horizontalMarginPoints,
      requestedCaseLabel);
    if (!initialAnalysis.Succeeded)
    {
      return AutomaticPlacementResult.Failed(initialAnalysis.Message, initialAnalysis);
    }

    if (!AnalysisMatchesRequest(initialAnalysis, worksheetName, images) || !SnapshotStillMatches(workbook, initialAnalysis))
    {
      // The user may edit Excel while the preview is open. Refresh the plan against
      // the current worksheet instead of rejecting an otherwise valid placement.
      var refreshed = Analyze(
        workbook,
        worksheetName,
        side,
        images,
        preferActiveGap,
        horizontalMarginPoints,
        requestedCaseLabel ?? initialAnalysis.CaseLabel);
      if (!refreshed.Succeeded)
      {
        return AutomaticPlacementResult.Failed(refreshed.Message, refreshed);
      }

      initialAnalysis = refreshed;
    }

    images = initialAnalysis.Steps.Select(step => step.Image).ToArray();
    var appliedRows = new List<AppliedRowInsertion>();
    if (!referencePrepared && images.Count == 1 && !images[0].PreserveReferenceSize && initialAnalysis.Steps[0].Pair is not null)
    {
      // Excel can be edited between Inspect and SetApplied. Re-analyze once so
      // placement uses the latest reference geometry instead of surfacing a
      // transient mismatch to the user.
      for (var attempt = 0; attempt < 2; attempt++)
      {
        if (initialAnalysis.Steps[0].Pair is not { } pair) break;
        var service = new ExcelManagedShapeService();
        var inspected = service.Inspect(workbook, initialAnalysis.WorksheetName, pair.ShapeName);
        if (!inspected.Succeeded || inspected.Shape is not { } before)
          return AutomaticPlacementResult.Failed(inspected.Message, initialAnalysis);
        var metadata = before.Metadata with { AppliedScale = pair.Legacy ? null : pair.Scale };
        var desired = before with { WidthPoints = pair.Width, HeightPoints = pair.Height,
          Metadata = metadata, AlternativeText = metadata.Serialize() };
        var extra = pair.Height - before.HeightPoints;
        var count = extra > 0.05 ? checked((int)Math.Ceiling(extra / 15) + 1) : 0;
        if ((long)pair.EndRow + count > ExcelWorksheetLimits.MaximumRow)
          return AutomaticPlacementResult.Failed("必要な行数がシート上限を超えます。", initialAnalysis);
        var reference = new PairedImageResize(before, desired, count > 0
          ? new AppliedRowInsertion(before.WorksheetName, pair.EndRow + 1, count, "参照画像の共通倍率用の領域") : null);
        var prepared = reference.SetApplied(workbook, true);
        if (!prepared.Succeeded && attempt == 0 && reference.CanRetryPreparation)
        {
          var refreshed = Analyze(workbook, worksheetName, side, images, preferActiveGap,
            horizontalMarginPoints, requestedCaseLabel ?? initialAnalysis.CaseLabel);
          if (!refreshed.Succeeded)
            return AutomaticPlacementResult.Failed(refreshed.Message, refreshed);
          initialAnalysis = refreshed;
          images = refreshed.Steps.Select(step => step.Image).ToArray();
          if (images[0].PreserveReferenceSize) break;
          continue;
        }
        if (!prepared.Succeeded && reference.CompensationSucceeded)
        {
          // Recovery starts with a fresh snapshot and keeps the existing reference unchanged.
          // The regular planner still validates band spacing, contents and CASE boundaries.
          var recovered = PlaceImages(workbook, initialAnalysis.WorksheetName, initialAnalysis.ResolvedSide,
            [images[0] with { PreserveReferenceSize = true, ScaleOverride = null }],
            preferActiveGap, horizontalMarginPoints, requestedCaseLabel: initialAnalysis.CaseLabel);
          return recovered.Succeeded
            ? recovered with { Message = recovered.Message + " 参照画像の大きさを保持して配置しました。" }
            : recovered;
        }
        if (!prepared.Succeeded) return AutomaticPlacementResult.Failed(prepared.Message, initialAnalysis) with
        {
          CompensationSucceeded = reference.CompensationSucceeded,
          CompensationErrors = reference.CompensationSucceeded ? [] : [prepared.Message],
        };
        var result = PlaceImages(workbook, worksheetName, side,
          [images[0] with { ScaleOverride = pair.Scale }], preferActiveGap, horizontalMarginPoints,
          requestedCaseLabel: initialAnalysis.CaseLabel, referencePrepared: true);
        if (result.Succeeded) return result with { ReferenceResize = reference };
        if (!result.CompensationSucceeded) return result with { ReferenceResize = reference };
        var undone = reference.SetApplied(workbook, false);
        return result with { CompensationSucceeded = undone.Succeeded,
          Message = result.Message + (undone.Succeeded ? "" : " " + undone.Message) };
      }
    }
    var placed = new List<AutomaticPlacedImage>();
    var executedSteps = new List<AutomaticPlacementStep>();
    for (var index = 0; index < images.Count; index++)
    {
      var step = initialAnalysis.Steps[index];
      var verifiedStep = step;
      var currentCaseEnd = initialAnalysis.LayoutAnalysis!.Layout!.EndRow + appliedRows.Sum(row => row.Count);
      var pendingInsertions = step.Plan.Insertions;
      for (var expansionPass = 0; pendingInsertions.Count > 0; expansionPass++)
      {
        if (expansionPass >= 3)
        {
          return Compensate(workbook, initialAnalysis, placed, appliedRows,
            "追加行を3回補正しても必要な配置領域を確保できなかったため、画像を配置せず行挿入を戻します。");
        }

        foreach (var insertion in pendingInsertions)
        {
          var appliedInsertion = ResolveAppliedInsertion(currentCaseEnd, insertion,
            initialAnalysis.LayoutAnalysis!.Layout!.CanDeleteTrailingRows);
          var mutation = rowMutationService.InsertRows(
            workbook,
            initialAnalysis.WorksheetName,
            appliedInsertion);
          if (!mutation.Succeeded || !mutation.Changed)
          {
            return Compensate(workbook, initialAnalysis, placed, appliedRows, mutation.Message);
          }

          appliedRows.Add(new AppliedRowInsertion(mutation.WorksheetName, mutation.StartRow, mutation.Count, insertion.Reason));
          var normalized = rowMutationService.NormalizeInsertedRows(
            workbook,
            mutation.WorksheetName,
            mutation.StartRow,
            mutation.Count,
            InsertedRowHeightPoints);
          if (!normalized.Succeeded || !normalized.Changed)
          {
            return Compensate(workbook, initialAnalysis, placed, appliedRows, normalized.Message);
          }
          currentCaseEnd = checked(currentCaseEnd + mutation.Count);
        }

        // Inserted rows can inherit a different height or Hidden state in Excel.
        // Verify the real space before adding the picture, not the assumed model alone.
        var expanded = Analyze(workbook, initialAnalysis.WorksheetName, side,
          [step.Image], preferActiveGap, horizontalMarginPoints, initialAnalysis.CaseLabel);
        if (!expanded.Succeeded)
        {
          return Compensate(workbook, initialAnalysis, placed, appliedRows,
            $"追加行の実際の高さ・配置位置を確認できないため、画像を配置せず行挿入を戻します。{expanded.Message}");
        }
        verifiedStep = expanded.Steps[0] with { Index = step.Index, Image = step.Image };
        pendingInsertions = verifiedStep.Plan.Insertions;
      }

      var placement = imagePlacementService.PlaceImage(
        workbook,
        initialAnalysis.WorksheetName,
        verifiedStep.Plan.FocusCell,
        side,
        verifiedStep.Image.ImagePath,
        verifiedStep.Image.Dimensions,
        verifiedStep.AvailableWidthPoints,
        horizontalMarginPoints,
        verifiedStep.Plan.Image.Scale);
      if (!placement.Succeeded)
      {
        return Compensate(workbook, initialAnalysis, placed, appliedRows, placement.Message);
      }

      placed.Add(new AutomaticPlacedImage(
        step.Index,
        placement.ShapeName,
        placement.WorksheetName,
        placement.FocusCell,
        placement.FocusSucceeded,
        placement.Target!,
        verifiedStep.Plan,
        verifiedStep.AvailableWidthPoints));
      executedSteps.Add(verifiedStep);
    }

    var executedAnalysis = initialAnalysis with { Steps = executedSteps };

    return new AutomaticPlacementResult(
      true,
      true,
      executedAnalysis,
      placed,
      appliedRows,
      [],
      $"{placed.Count}件の画像を配置しました（未保存）。");
  }

  private bool SnapshotStillMatches(
    WorkbookIdentity workbook,
    AutomaticPlacementAnalysisResult expected)
  {
    var current = snapshotService.Capture(workbook, expected.WorksheetName, expected.AnalysisRow);
    if (!current.Succeeded || current.Snapshot is null)
    {
      return false;
    }

    var normalized = current.Snapshot with
    {
      ActiveCell = new CellReference(expected.AnalysisRow, current.Snapshot.ActiveCell.Column),
      LayoutSignals = current.Snapshot.LayoutSignals with { ActiveRow = expected.AnalysisRow },
    };
    return string.Equals(Fingerprint(normalized), expected.SnapshotFingerprint, StringComparison.Ordinal);
  }

  private static bool AnalysisMatchesRequest(
    AutomaticPlacementAnalysisResult analysis,
    string worksheetName,
    IReadOnlyList<AutomaticPlacementImage> images) =>
    (string.Equals(worksheetName, "ActiveSheet", StringComparison.OrdinalIgnoreCase) ||
      string.Equals(worksheetName, analysis.WorksheetName, StringComparison.OrdinalIgnoreCase)) &&
    analysis.Steps.Count == images.Count &&
    analysis.Steps.Select(step => step.Image).SequenceEqual(images);

  private static string ResolveCaseLabel(SheetSnapshot snapshot, int startRow)
  {
    var anchor = ConfirmedAnchors(snapshot.LayoutSignals).Last(anchor => anchor.Row <= startRow);
    return FormatCaseLabel(anchor);
  }

  private RequestedCaseResolution ResolveRequestedCase(
    SheetSnapshot snapshot,
    string? requestedCaseLabel,
    EvidenceSide side,
    bool sameCaseThenNext)
  {
    if (string.IsNullOrWhiteSpace(requestedCaseLabel))
    {
      var anchors = ConfirmedAnchors(snapshot.LayoutSignals);
      var activeAnchor = anchors.LastOrDefault(anchor => anchor.Row <= snapshot.ActiveCell.Row);
      var activeCaseEnd = activeAnchor is null
        ? 0
        : anchors.FirstOrDefault(anchor => anchor.Row > activeAnchor.Row)?.Row - 1 ??
          snapshot.LayoutSignals.LogicalEvidenceLastRow;
      if (activeAnchor is not null && snapshot.ActiveCell.Row <= activeCaseEnd)
      {
        return new RequestedCaseResolution(true, activeAnchor.Row, null, string.Empty);
      }

      foreach (var anchor in anchors)
      {
        var layout = layoutAnalyzer.Analyze(snapshot.LayoutSignals with { ActiveRow = anchor.Row }).Layout;
        if (layout is null)
        {
          continue;
        }

        var availableSides = sameCaseThenNext && layout.Kind == SideLayoutKind.Both
          ? new[] { side, side is EvidenceSide.New ? EvidenceSide.Old : EvidenceSide.New }
          : [layout.SupportsSide(side) ? side : EvidenceSide.New];
        foreach (var candidateSide in availableSides.Distinct())
        {
          if (!ExcelCaseNavigationService.IsOccupied(snapshot, layout, candidateSide))
          {
            return new RequestedCaseResolution(true, anchor.Row, candidateSide, string.Empty);
          }
        }
      }

      return new RequestedCaseResolution(false, 0, null, anchors.Length == 0
        ? "Case番号を検出できません。"
        : "未配置のCaseがありません。");
    }

    var normalizedLabel = NormalizeCaseLabel(requestedCaseLabel);
    if (normalizedLabel is null)
    {
      return new RequestedCaseResolution(false, 0, null, "CaseはX-X形式で入力してください（例: 1-2）。");
    }

    var matches = ConfirmedAnchors(snapshot.LayoutSignals)
      .Where(anchor => string.Equals(FormatCaseLabel(anchor), normalizedLabel, StringComparison.OrdinalIgnoreCase))
      .ToArray();
    return matches.Length switch
    {
      1 => new RequestedCaseResolution(true, matches[0].Row, null, string.Empty),
      0 => new RequestedCaseResolution(false, 0, null, $"Case '{normalizedLabel}' が見つかりません。"),
      _ => new RequestedCaseResolution(false, 0, null,
        $"Case '{normalizedLabel}' が複数あります（行: {string.Join(", ", matches.Select(anchor => anchor.Row))}）。"),
    };
  }

  public static string FormatCaseLabel(CaseAnchorSignal anchor)
  {
    var major = anchor.ColumnAValue?.Trim();
    var minor = anchor.ColumnBValue?.Trim();
    return string.IsNullOrWhiteSpace(major) || string.IsNullOrWhiteSpace(minor)
      ? $"開始行 {anchor.Row}"
      : $"{major}-{minor}";
  }

  public static CaseAnchorSignal[] ConfirmedAnchors(SheetLayoutSignals signals) =>
    CaseAnchorNormalizer.Normalize(signals.Anchors);

  internal static string? NormalizeCaseLabel(string label) => CaseAnchorNormalizer.NormalizeCaseLabel(label);

  private static EvidenceSide SideForColumn(
    int column,
    EvidenceCaseLayout layout,
    EvidenceSide fallback) =>
    layout.OldRegion is { } old && column >= old.FirstColumn && column <= old.LastColumn
      ? EvidenceSide.Old
      : column >= layout.NewRegion.FirstColumn && column <= layout.NewRegion.LastColumn
        ? EvidenceSide.New
        : layout.SupportsSide(fallback) ? fallback : EvidenceSide.New;

  private static RowInsertion ResolveAppliedInsertion(int currentCaseEnd, RowInsertion insertion, bool hasNextCase)
  {
    // A following Case anchor moves down with the insertion and defines the new end.
    // Inserting above the current last row instead moves its content on every retry,
    // so a required tail after that content can never be created.
    return !hasNextCase && currentCaseEnd < ExcelWorksheetLimits.MaximumRow && insertion.AtRow == currentCaseEnd + 1
      ? insertion with
      {
        AtRow = currentCaseEnd,
        Reason = $"{insertion.Reason} Insert above the Case bottom boundary.",
      }
      : insertion;
  }

  private AutomaticPlacementResult Compensate(
    WorkbookIdentity workbook,
    AutomaticPlacementAnalysisResult analysis,
    IReadOnlyList<AutomaticPlacedImage> placed,
    IReadOnlyList<AppliedRowInsertion> insertedRows,
    string failure)
  {
    var errors = new List<string>();
    foreach (var image in placed.Reverse())
    {
      var deletion = imagePlacementService.DeletePlacedImage(workbook, image.WorksheetName, image.ShapeName);
      if (!deletion.Succeeded)
      {
        errors.Add($"Shape {image.ShapeName}: {deletion.Message}");
      }
    }

    foreach (var insertion in insertedRows.Reverse())
    {
      var deletion = rowMutationService.DeleteRowsIfSafe(
        workbook,
        insertion.WorksheetName,
        insertion.StartRow,
        insertion.Count);
      if (!deletion.Succeeded || !deletion.Changed)
      {
        errors.Add($"Rows {insertion.StartRow}-{insertion.StartRow + insertion.Count - 1}: {deletion.Message}");
      }
    }

    var compensated = errors.Count == 0;
    var message = compensated
      ? $"自動配置に失敗したため変更を取り消しました: {failure}"
      : $"自動配置に失敗し、一部を自動復旧できませんでした: {failure} {string.Join(" ", errors)}";
    return new AutomaticPlacementResult(
      false,
      compensated,
      analysis,
      placed,
      insertedRows,
      errors,
      message);
  }

  private static string? ValidateImages(IReadOnlyList<AutomaticPlacementImage>? images, bool requireFiles)
  {
    if (images is null || images.Count == 0)
    {
      return "配置する画像がありません。";
    }

    foreach (var image in images)
    {
      if (image is null || string.IsNullOrWhiteSpace(image.ImagePath))
      {
        return "画像パスが指定されていません。";
      }

      if (!double.IsFinite(image.Dimensions.WidthPoints) || image.Dimensions.WidthPoints <= 0 ||
        !double.IsFinite(image.Dimensions.HeightPoints) || image.Dimensions.HeightPoints <= 0)
      {
        return $"画像サイズが不正です: {image.ImagePath}";
      }

      if (requireFiles && !File.Exists(image.ImagePath))
      {
        return $"画像ファイルが見つかりません: {image.ImagePath}";
      }
    }

    return null;
  }

  private static double AvailableWidth(
    SheetSnapshot snapshot,
    ColumnRange columns,
    double horizontalMarginPoints)
  {
    if (!double.IsFinite(horizontalMarginPoints) || horizontalMarginPoints < 0)
    {
      throw new ArgumentOutOfRangeException(nameof(horizontalMarginPoints));
    }

    if (columns.Count < 2)
    {
      throw new InvalidOperationException("配置先Sideには余白列と画像列が必要です。");
    }

    var width = Enumerable.Range(columns.FirstColumn + 1, columns.Count - 1)
      .Sum(column => snapshot.ColumnWidths.GetValueOrDefault(column));
    var availableWidth = width - (horizontalMarginPoints * 2);
    if (!double.IsFinite(availableWidth) || availableWidth <= 0)
    {
      throw new InvalidOperationException("配置先Sideの列幅を取得できません。");
    }

    return availableWidth;
  }

  private static PairedImagePlan? FindPair(SheetSnapshot snapshot, EvidenceCaseLayout layout,
    EvidenceSide side, ImageDimensions image, double width, double margin, bool preserveReferenceSize = false)
  {
    if (layout.Kind != SideLayoutKind.Both) return null;
    var opposite = side == EvidenceSide.New ? EvidenceSide.Old : EvidenceSide.New;
    var reference = CaseImages(snapshot, layout, opposite).ElementAtOrDefault(CaseImages(snapshot, layout, side).Length);
    if (reference is null) return null;
    // Old snapshots without picture geometry cannot establish a safe reference scale.
    if (reference.WidthPoints <= 0 || reference.HeightPoints <= 0) return null;
    var region = layout.RegionFor(opposite);
    var remainingWidth = Enumerable.Range(reference.StartColumn, region.LastColumn - reference.StartColumn + 1)
      .Sum(column => snapshot.ColumnWidths.GetValueOrDefault(column)) - reference.HorizontalOffsetPoints - margin;
    var referenceWidth = Math.Min(AvailableWidth(snapshot, region, margin), remainingWidth);
    if (referenceWidth <= 0) throw new InvalidOperationException("参照画像の配置幅がありません。");
    if (preserveReferenceSize)
    {
      if (reference.WidthPoints > referenceWidth + 0.05)
        throw new InvalidOperationException("参照画像が配置範囲を超えているため整合性を確認できません。");
      var currentScale = reference.SourceDimensions is { } original
        ? reference.WidthPoints / original.WidthPoints
        : reference.WidthPoints / image.WidthPoints;
      var scale = Math.Min(currentScale, width / image.WidthPoints);
      return new(reference.Name, scale, reference.WidthPoints, reference.HeightPoints,
        reference.StartRow, reference.EndRow, reference.SourceDimensions is null);
    }
    if (reference.SourceDimensions is { } source)
    {
      var scale = Math.Min(width / image.WidthPoints, referenceWidth / source.WidthPoints);
      return new(reference.Name, scale, source.WidthPoints * scale, source.HeightPoints * scale,
        reference.StartRow, reference.EndRow, false);
    }
    // Legacy pictures lack their original dimensions: match displayed widths, preserving aspect ratios.
    var commonWidth = Math.Min(width, referenceWidth);
    return new(reference.Name, commonWidth / image.WidthPoints, commonWidth,
      reference.HeightPoints * commonWidth / reference.WidthPoints,
      reference.StartRow, reference.EndRow, true);
  }

  private static SnapshotShape[] CaseImages(SheetSnapshot snapshot, EvidenceCaseLayout layout, EvidenceSide side)
  {
    var columns = layout.RegionFor(side);
    return snapshot.Shapes.Where(shape => shape.IsManagedImage &&
      shape.StartRow >= layout.StartRow && shape.EndRow <= layout.EndRow &&
      shape.StartColumn >= columns.FirstColumn && shape.EndColumn <= columns.LastColumn)
      .OrderBy(shape => shape.StartRow).ThenBy(shape => shape.TopPoints)
      .ThenBy(shape => shape.Name, StringComparer.Ordinal).ToArray();
  }

  private static string Fingerprint(SheetSnapshot snapshot)
  {
    var text = new StringBuilder()
      .Append(snapshot.WorksheetName).Append('|')
      .Append(snapshot.ActiveCell.Row).Append(',').Append(snapshot.ActiveCell.Column).Append('|')
      .Append(snapshot.RawUsedFirstRow).Append(',').Append(snapshot.RawUsedLastRow).Append(',')
      .Append(snapshot.RawUsedFirstColumn).Append(',').Append(snapshot.RawUsedLastColumn).Append('|')
      .Append(snapshot.IsReadOnly).Append(',').Append(snapshot.IsProtected).Append('|');
    var signals = snapshot.LayoutSignals;
    text.Append('L').Append(signals.ActiveRow).Append(',').Append(signals.RawUsedLastRow).Append(',')
      .Append(signals.LogicalEvidenceLastRow).Append(',').Append(signals.NewFirstColumn).Append(',')
      .Append(signals.ObservedLastEvidenceColumn).Append('|');
    foreach (var anchor in signals.Anchors.OrderBy(anchor => anchor.Row))
    {
      text.Append('A').Append(anchor.Row).Append(',').Append(anchor.HasValueInColumnA).Append(',')
        .Append(anchor.HasValueInColumnB).Append(',').Append(anchor.HasTopBorder).Append(',')
        .Append(anchor.ColumnAValue).Append(',').Append(anchor.ColumnBValue).Append('|');
    }

    foreach (var boundary in signals.VerticalBoundaries
      .OrderBy(boundary => boundary.RightColumn)
      .ThenBy(boundary => boundary.StartRow))
    {
      text.Append('V').Append(boundary.RightColumn).Append(',').Append(boundary.StartRow).Append(',')
        .Append(boundary.EndRow).Append('|');
    }

    foreach (var boundary in signals.HorizontalBoundaries
      .OrderBy(boundary => boundary.Row)
      .ThenBy(boundary => boundary.FirstColumn))
    {
      text.Append('H').Append(boundary.Row).Append(',').Append(boundary.FirstColumn).Append(',')
        .Append(boundary.LastColumn).Append('|');
    }

    foreach (var column in signals.NewHeaderColumns.Order())
    {
      text.Append('N').Append(column).Append('|');
    }

    foreach (var column in signals.OldHeaderColumns.Order())
    {
      text.Append('O').Append(column).Append('|');
    }

    foreach (var cell in snapshot.Cells.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column))
    {
      text.Append('C').Append(cell.Row).Append(',').Append(cell.Column).Append(',')
        .Append(cell.HasValueOrFormula).Append(',').Append(cell.HasCommentOrNote).Append(',')
        .Append(cell.HasHyperlink).Append(',').Append(cell.IntersectsMerge).Append('|');
    }

    foreach (var shape in snapshot.Shapes.OrderBy(shape => shape.Name, StringComparer.Ordinal))
    {
      text.Append('S').Append(shape.Name).Append(',').Append(shape.StartRow).Append(',')
        .Append(shape.EndRow).Append(',').Append(shape.StartColumn).Append(',')
        .Append(shape.EndColumn).Append(',').Append(shape.IsManagedImage).Append('|')
        .Append(FormattableString.Invariant($"{shape.TopPoints:R},{shape.WidthPoints:R},{shape.HeightPoints:R},{shape.HorizontalOffsetPoints:R},{shape.SourceDimensions?.WidthPoints:R},{shape.SourceDimensions?.HeightPoints:R}|"));
    }

    foreach (var height in snapshot.RowHeights.OrderBy(pair => pair.Key))
    {
      text.Append('R').Append(height.Key).Append('=').Append(height.Value.ToString("R", System.Globalization.CultureInfo.InvariantCulture)).Append('|');
    }

    foreach (var width in snapshot.ColumnWidths.OrderBy(pair => pair.Key))
    {
      text.Append('W').Append(width.Key).Append('=').Append(width.Value.ToString("R", System.Globalization.CultureInfo.InvariantCulture)).Append('|');
    }

    return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text.ToString())));
  }

  private static void ApplyPlanToModel(
    PlacementPlan plan,
    EvidenceSide side,
    ref EvidenceCaseLayout layout,
    List<ContentSpan> contents,
    Dictionary<int, double> rowHeights)
  {
    foreach (var insertion in plan.Insertions)
    {
      for (var index = 0; index < contents.Count; index++)
      {
        var content = contents[index];
        contents[index] = content with
        {
          StartRow = content.StartRow >= insertion.AtRow ? checked(content.StartRow + insertion.Count) : content.StartRow,
          EndRow = content.EndRow >= insertion.AtRow ? checked(content.EndRow + insertion.Count) : content.EndRow,
        };
      }

      var shiftedHeights = rowHeights
        .OrderByDescending(pair => pair.Key)
        .Where(pair => pair.Key >= insertion.AtRow)
        .ToArray();
      foreach (var pair in shiftedHeights)
      {
        rowHeights.Remove(pair.Key);
        rowHeights[checked(pair.Key + insertion.Count)] = pair.Value;
      }

      layout = layout with { EndRow = checked(layout.EndRow + insertion.Count) };
    }

    contents.Add(new ContentSpan(side, plan.StartRow, plan.EndRow, ContentKind.ManagedImage));
  }

  private sealed record RequestedCaseResolution(
    bool Succeeded,
    int Row,
    EvidenceSide? Side,
    string Message);
}

public sealed record AutomaticPlacementImage(string ImagePath, ImageDimensions Dimensions)
{
  internal bool PreserveReferenceSize { get; init; }
  public double? ScaleOverride { get; init; }
}

public sealed record AutomaticPlacementStep(
  int Index,
  AutomaticPlacementImage Image,
  PlacementPlan Plan,
  double AvailableWidthPoints)
{
  public PairedImagePlan? Pair { get; init; }
}

public sealed record AutomaticPlacementAnalysisResult(
  bool Succeeded,
  string WorksheetName,
  LayoutAnalysisResult? LayoutAnalysis,
  IReadOnlyList<AutomaticPlacementStep> Steps,
  string SnapshotFingerprint,
  string CaseLabel,
  int AnalysisRow,
  EvidenceSide ResolvedSide,
  string Message)
{
  public SheetLayoutSignals? LayoutSignals { get; init; }
  public bool CompletesCaseAfterPlacement { get; init; }

  public static AutomaticPlacementAnalysisResult Failed(string message) =>
    new(false, string.Empty, null, [], string.Empty, string.Empty, 0, EvidenceSide.New, message);
}

public sealed record AppliedRowInsertion(string WorksheetName, int StartRow, int Count, string Reason);

public sealed record AutomaticPlacedImage(
  int Index,
  string ShapeName,
  string WorksheetName,
  CellReference FocusCell,
  bool FocusSucceeded,
  ManagedShapeTarget Target,
  PlacementPlan Plan,
  double AvailableWidthPoints = 0);

public sealed record AutomaticPlacementResult(
  bool Succeeded,
  bool CompensationSucceeded,
  AutomaticPlacementAnalysisResult? Analysis,
  IReadOnlyList<AutomaticPlacedImage> PlacedImages,
  IReadOnlyList<AppliedRowInsertion> AppliedInsertions,
  IReadOnlyList<string> CompensationErrors,
  string Message)
{
  public PairedImageResize? ReferenceResize { get; init; }
  public static AutomaticPlacementResult Failed(
    string message,
    AutomaticPlacementAnalysisResult? analysis = null) =>
    new(false, true, analysis, [], [], [], message);
}
