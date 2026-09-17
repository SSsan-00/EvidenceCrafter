using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class AutomaticPlacementServiceTests
{
  [TestMethod]
  public void AnalyzeSnapshot_ExposesCurrentCaseLabelForPreview()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json");
    var snapshot = Snapshot(signals with
    {
      Anchors = [new CaseAnchorSignal(3, true, true, false, "1", "1")],
    });

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))]);

    Assert.IsTrue(result.Succeeded, result.Message);
    Assert.AreEqual("1-1", result.CaseLabel);
    Assert.AreEqual(snapshot.LayoutSignals with { ActiveRow = 3 }, result.LayoutSignals);
  }

  [TestMethod]
  public void FormatCaseLabel_CombinesBothCaseColumns()
  {
    var anchor = new CaseAnchorSignal(3, true, true, false, "1", "2");

    Assert.AreEqual("1-2", ExcelAutomaticPlacementService.FormatCaseLabel(anchor));
  }

  [TestMethod]
  public void ConfirmedAnchors_KeepCanonicalLabelsForInheritedMajorNumber()
  {
    var source = FixtureLoader.LoadLayout("default-final-case.json");
    var signals = source with
    {
      Anchors =
      [
        new CaseAnchorSignal(3, true, true, false, "1", "1"),
        new CaseAnchorSignal(33, false, true, true, null, "2"),
        new CaseAnchorSignal(63, true, true, true, "2", "1"),
        new CaseAnchorSignal(93, false, true, true, null, "2"),
      ],
    };

    var labels = ExcelAutomaticPlacementService.ConfirmedAnchors(signals)
      .Select(ExcelAutomaticPlacementService.FormatCaseLabel)
      .ToArray();

    CollectionAssert.AreEqual(new[] { "1-1", "1-2", "2-1", "2-2" }, labels);
  }

  [TestMethod]
  public void AnalyzeSnapshot_PartialCaseLabel_IsRejected()
  {
    var snapshot = Snapshot(FixtureLoader.LoadLayout("default-final-case.json"));

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))],
      requestedCaseLabel: "2");

    Assert.IsFalse(result.Succeeded);
    StringAssert.Contains(result.Message, "X-X形式");
  }

  [TestMethod]
  public void NormalizeCaseLabel_AcceptsCommonFullWidthDash()
  {
    Assert.AreEqual("2-3", ExcelAutomaticPlacementService.NormalizeCaseLabel(" 2－3 "));
  }

  [TestMethod]
  public void AnalyzeSnapshot_CanonicalDuplicate_FailsClosedWithRows()
  {
    var source = FixtureLoader.LoadLayout("default-final-case.json");
    var snapshot = Snapshot(source with
    {
      Anchors =
      [
        new CaseAnchorSignal(3, true, true, false, "1", "1"),
        new CaseAnchorSignal(53, true, true, true, "1", "1"),
      ],
    });

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))],
      requestedCaseLabel: "1-1");

    Assert.IsFalse(result.Succeeded);
    StringAssert.Contains(result.Message, "行: 3, 53");
  }

  [TestMethod]
  public void AnalyzeSnapshot_RequestedCaseLabel_OverridesActiveCase()
  {
    var source = FixtureLoader.LoadLayout("default-final-case.json");
    var snapshot = Snapshot(source with
    {
      Anchors =
      [
        new CaseAnchorSignal(3, true, true, false, "1", "1"),
        new CaseAnchorSignal(53, true, true, true, "1", "2"),
      ],
    });

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))],
      requestedCaseLabel: "1-1");

    Assert.IsTrue(result.Succeeded, result.Message);
    Assert.AreEqual("1-1", result.CaseLabel);
    Assert.AreEqual(3, result.AnalysisRow);
    Assert.AreEqual(3, result.LayoutAnalysis!.Layout!.StartRow);
  }

  [TestMethod]
  public void AnalyzeSnapshot_ActiveCellAboveFirstCase_UsesFirstUnoccupiedCase()
  {
    var source = FixtureLoader.LoadLayout("default-final-case.json");
    var signals = source with
    {
      ActiveRow = 1,
      Anchors =
      [
        new CaseAnchorSignal(3, true, true, false, "1", "1"),
        new CaseAnchorSignal(53, true, true, true, "1", "2"),
      ],
    };
    var snapshot = Snapshot(signals) with
    {
      ActiveCell = new CellReference(1, 1),
      Shapes = [new SnapshotShape("existing", 4, 20, 4, 10, true)],
    };

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))]);

    Assert.IsTrue(result.Succeeded, result.Message);
    Assert.AreEqual("1-2", result.CaseLabel);
    Assert.AreEqual(53, result.AnalysisRow);
    Assert.AreEqual(new CellReference(55, 4), result.Steps[0].Plan.FocusCell);
  }

  [TestMethod]
  public void AnalyzeSnapshot_OutsideCases_SameCaseModeUsesEmptyOppositeSideFirst()
  {
    var source = FixtureLoader.LoadLayout("default-final-case.json");
    var signals = source with { ActiveRow = 1 };
    var snapshot = Snapshot(signals) with
    {
      ActiveCell = new CellReference(1, 1),
      Shapes = [new SnapshotShape("existing-new", 4, 20, 4, 10, true)],
    };

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))],
      sameCaseThenNext: true);

    Assert.IsTrue(result.Succeeded, result.Message);
    Assert.AreEqual("1-1", result.CaseLabel);
    Assert.AreEqual(EvidenceSide.Old, result.ResolvedSide);
    Assert.AreEqual(new CellReference(5, 19), result.Steps[0].Plan.FocusCell);
  }

  [TestMethod]
  public void AnalyzeSnapshot_ActiveCellBelowEvidence_UsesFirstUnoccupiedCase()
  {
    var source = FixtureLoader.LoadLayout("default-final-case.json");
    var signals = source with { ActiveRow = 200 };
    var snapshot = Snapshot(signals) with { ActiveCell = new CellReference(200, 1) };

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))]);

    Assert.IsTrue(result.Succeeded, result.Message);
    Assert.AreEqual("1-1", result.CaseLabel);
    Assert.AreEqual(3, result.AnalysisRow);
  }

  [TestMethod]
  public void AnalyzeSnapshot_ActiveCellOutsideCases_FailsWhenEveryCaseIsOccupied()
  {
    var source = FixtureLoader.LoadLayout("default-final-case.json");
    var signals = source with { ActiveRow = 1 };
    var snapshot = Snapshot(signals) with
    {
      ActiveCell = new CellReference(1, 1),
      Shapes =
      [
        new SnapshotShape("first", 4, 20, 4, 10, true),
        new SnapshotShape("second", 54, 70, 4, 10, true),
      ],
    };

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))]);

    Assert.IsFalse(result.Succeeded);
    StringAssert.Contains(result.Message, "未配置のCaseがありません");
  }

  [TestMethod]
  public void AnalyzeSnapshot_UnknownRequestedCaseLabel_FailsClearly()
  {
    var source = FixtureLoader.LoadLayout("default-final-case.json");
    var snapshot = Snapshot(source with
    {
      Anchors =
      [
        new CaseAnchorSignal(3, true, true, false, "1", "1"),
        new CaseAnchorSignal(53, true, true, true, "1", "2"),
      ],
    });

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))],
      requestedCaseLabel: "9-9");

    Assert.IsFalse(result.Succeeded);
    StringAssert.Contains(result.Message, "Case '9-9' が見つかりません");
  }

  [TestMethod]
  public void AnalyzeSnapshot_DefaultPlacement_IgnoresActiveCellGap()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json");
    var snapshot = Snapshot(signals);

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))]);

    Assert.IsTrue(result.Succeeded, result.Message);
    Assert.HasCount(1, result.Steps);
    Assert.AreEqual(PlacementMode.CaseStart, result.Steps[0].Plan.Mode);
    Assert.AreEqual(55, result.Steps[0].Plan.StartRow);
    Assert.AreEqual(new CellReference(55, 4), result.Steps[0].Plan.FocusCell);
    Assert.IsFalse(string.IsNullOrWhiteSpace(result.SnapshotFingerprint));
  }

  [TestMethod]
  public void AnalyzeSnapshot_ProtectedSheet_FailsWithoutPlan()
  {
    var snapshot = Snapshot(FixtureLoader.LoadLayout("default-final-case.json")) with
    {
      IsProtected = true,
    };

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))]);

    Assert.IsFalse(result.Succeeded);
    Assert.IsEmpty(result.Steps);
    StringAssert.Contains(result.Message, "保護");
  }

  [TestMethod]
  public void AnalyzeSnapshot_CustomMargin_ReducesAvailableWidthOnBothSides()
  {
    var snapshot = Snapshot(FixtureLoader.LoadLayout("default-final-case.json"));
    var image = new[] { new AutomaticPlacementImage("image.png", new ImageDimensions(1_000, 500)) };
    var service = new ExcelAutomaticPlacementService();

    var defaultMargin = service.AnalyzeSnapshot(snapshot, EvidenceSide.New, image, horizontalMarginPoints: 6);
    var wideMargin = service.AnalyzeSnapshot(snapshot, EvidenceSide.New, image, horizontalMarginPoints: 20);

    Assert.IsTrue(defaultMargin.Succeeded, defaultMargin.Message);
    Assert.IsTrue(wideMargin.Succeeded, wideMargin.Message);
    Assert.AreEqual(28, defaultMargin.Steps[0].AvailableWidthPoints - wideMargin.Steps[0].AvailableWidthPoints, 0.001);
  }

  [TestMethod]
  public void AnalyzeSnapshot_PairsOppositeImageByOrdinal_AtTheLargerDisplayedWidth()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json") with { ActiveRow = 3 };
    var snapshot = Snapshot(signals) with
    {
      Shapes = [new SnapshotShape("old-1", 5, 10, 19, 25, true)
      {
        WidthPoints = 100,
        HeightPoints = 50,
        SourceDimensions = new ImageDimensions(200, 100),
      }],
    };

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("new.png", new ImageDimensions(100, 50))]);

    Assert.IsTrue(result.Succeeded, result.Message);
    var pair = result.Steps[0].Pair ?? throw new AssertFailedException("Expected an opposite-side pair.");
    Assert.AreEqual("old-1", pair.ShapeName);
    Assert.AreEqual(1, pair.Scale, 0.001);
    Assert.AreEqual(0.5, pair.ReferenceScale!.Value, 0.001);
    Assert.AreEqual(100, pair.Width, 0.001);
    Assert.AreEqual(100, result.Steps[0].Plan.Image.WidthPoints, 0.001);
    Assert.AreEqual(5, result.Steps[0].Plan.StartRow);
  }

  [TestMethod]
  public void AnalyzeSnapshot_PairedWidthUsesTheLargerImageButRespectsBothSides()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json") with { ActiveRow = 3 };
    var snapshot = Snapshot(signals) with
    {
      Shapes = [new SnapshotShape("old-1", 5, 10, 19, 25, true)
      {
        WidthPoints = 120,
        HeightPoints = 60,
        SourceDimensions = new ImageDimensions(120, 60),
      }],
    };

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("new.png", new ImageDimensions(500, 250))]);

    Assert.IsTrue(result.Succeeded, result.Message);
    var pair = result.Steps[0].Pair ?? throw new AssertFailedException("Expected an opposite-side pair.");
    Assert.AreEqual(268, pair.Width, 0.001);
    Assert.AreEqual(268d / 500d, pair.Scale, 0.001);
    Assert.AreEqual(268d / 120d, pair.ReferenceScale!.Value, 0.001);
    Assert.AreEqual(268, result.Steps[0].Plan.Image.WidthPoints, 0.001);
  }

  [TestMethod]
  public void AnalyzeSnapshot_PairsSecondImageByOrdinal_NotNearestImage()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json") with { ActiveRow = 3 };
    var snapshot = Snapshot(signals) with
    {
      Shapes =
      [
        new SnapshotShape("new-1", 5, 10, 4, 10, true) { WidthPoints = 80, HeightPoints = 40, SourceDimensions = new ImageDimensions(100, 50) },
        new SnapshotShape("new-2", 20, 25, 4, 10, true) { WidthPoints = 120, HeightPoints = 60, SourceDimensions = new ImageDimensions(120, 60) },
        new SnapshotShape("old-1", 5, 10, 19, 25, true) { WidthPoints = 90, HeightPoints = 45, SourceDimensions = new ImageDimensions(90, 45) },
      ],
    };

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.Old,
      [new AutomaticPlacementImage("old.png", new ImageDimensions(60, 30))]);

    Assert.IsTrue(result.Succeeded, result.Message);
    var pair = result.Steps[0].Pair ?? throw new AssertFailedException("Expected the second opposite-side image.");
    Assert.AreEqual("new-2", pair.ShapeName);
    Assert.AreEqual(2, pair.Scale, 0.001);
    Assert.AreEqual(20, result.Steps[0].Plan.StartRow);
  }

  [TestMethod]
  public void AnalyzeSnapshot_LegacyPairMatchesDisplayedWidth()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json") with { ActiveRow = 3 };
    var snapshot = Snapshot(signals) with
    {
      Shapes = [new SnapshotShape("legacy-old", 5, 10, 19, 25, true)
      {
        WidthPoints = 100,
        HeightPoints = 50,
      }],
    };

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("new.png", new ImageDimensions(50, 25))]);

    Assert.IsTrue(result.Succeeded, result.Message);
    var pair = result.Steps[0].Pair ?? throw new AssertFailedException("Expected a legacy opposite-side pair.");
    Assert.IsTrue(pair.Legacy);
    Assert.AreEqual(2, pair.Scale, 0.001);
    Assert.AreEqual(100, pair.Width, 0.001);
  }

  [TestMethod]
  public void AnalyzeSnapshot_NextBandStartsBelowBothSides()
  {
    var snapshot = Snapshot(FixtureLoader.LoadLayout("default-final-case.json") with { ActiveRow = 3 }) with
    {
      Shapes = [new SnapshotShape("new-1", 5, 10, 4, 10, true),
        new SnapshotShape("old-1", 5, 20, 19, 25, true)],
    };
    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(snapshot, EvidenceSide.New,
      [new AutomaticPlacementImage("new.png", new ImageDimensions(60, 30))]);
    Assert.IsTrue(result.Succeeded, result.Message);
    Assert.AreEqual(23, result.Steps[0].Plan.StartRow);
  }

  [TestMethod]
  public void AnalyzeSnapshot_RecoveryKeepsReferenceGeometryAndBandSpacing()
  {
    var snapshot = Snapshot(FixtureLoader.LoadLayout("default-final-case.json") with { ActiveRow = 3 }) with
    {
      Shapes = [new SnapshotShape("new-1", 5, 10, 4, 10, true)
        { WidthPoints = 60, HeightPoints = 30, SourceDimensions = new ImageDimensions(120, 60) },
        new SnapshotShape("new-2", 15, 20, 4, 10, true)],
    };
    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(snapshot, EvidenceSide.Old,
      [new AutomaticPlacementImage("old.png", new ImageDimensions(120, 600)) { PreserveReferenceSize = true }]);
    Assert.IsTrue(result.Succeeded, result.Message);
    Assert.AreEqual(0.5, result.Steps[0].Plan.Image.Scale, 0.001);
    Assert.AreEqual(60d, result.Steps[0].Pair!.Width);
    Assert.AreEqual(30d, result.Steps[0].Pair!.Height);
    Assert.AreEqual(5, result.Steps[0].Plan.StartRow);
    var insertion = result.Steps[0].Plan.Insertions.Single(row => row.AtRow == 15);
    Assert.IsGreaterThan(result.Steps[0].Plan.EndRow + 2, 15 + insertion.Count);
    Assert.IsFalse(new ExcelAutomaticPlacementService().AnalyzeSnapshot(snapshot with { IsProtected = true },
      EvidenceSide.Old, [new AutomaticPlacementImage("old.png", new ImageDimensions(120, 600))
        { PreserveReferenceSize = true }]).Succeeded);
  }

  [TestMethod]
  public void AnalyzeSnapshot_ReanalysisKeepsTheOriginalPairedShapeAfterItsOrderChanges()
  {
    var snapshot = Snapshot(FixtureLoader.LoadLayout("default-final-case.json") with { ActiveRow = 3 }) with
    {
      // new-1 was relocated below new-2 while making room for the matching OLD image.
      Shapes =
      [
        new SnapshotShape("new-2", 5, 10, 4, 10, true),
        new SnapshotShape("new-1", 30, 35, 4, 10, true)
        {
          WidthPoints = 60, HeightPoints = 30, SourceDimensions = new ImageDimensions(120, 60),
        },
      ],
    };

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(snapshot, EvidenceSide.Old,
      [new AutomaticPlacementImage("old.png", new ImageDimensions(120, 60))
      {
        ReferenceShapeName = "new-1",
      }]);

    Assert.IsTrue(result.Succeeded, result.Message);
    Assert.AreEqual("new-1", result.Steps[0].Pair!.ShapeName);
    Assert.AreEqual(30, result.Steps[0].Plan.StartRow);
  }

  [TestMethod]
  public void AnalyzeSnapshot_PairedPlacementMakesRoomForExistingCells()
  {
    var snapshot = Snapshot(FixtureLoader.LoadLayout("default-final-case.json") with { ActiveRow = 3 }) with
    {
      Shapes = [new SnapshotShape("new-1", 5, 7, 4, 10, true)
        { WidthPoints = 60, HeightPoints = 30, SourceDimensions = new ImageDimensions(120, 60) }],
      Cells = [new SnapshotCell(12, 19, true, false, false, false)],
    };
    var service = new ExcelAutomaticPlacementService();
    var images = new[] { new AutomaticPlacementImage("old.png", new ImageDimensions(120, 240)) };
    var recovered = service.AnalyzeSnapshot(snapshot, EvidenceSide.Old, images);
    Assert.IsTrue(recovered.Succeeded, recovered.Message);
    var plan = recovered.Steps[0].Plan;
    Assert.AreEqual(5, plan.StartRow);
    var insertion = plan.Insertions.Single(row => row.AtRow == 12);
    Assert.IsGreaterThan(plan.EndRow, 12 + insertion.Count);
    var blocked = service.AnalyzeSnapshot(snapshot with
    {
      Cells = [new SnapshotCell(5, 19, true, false, false, false)],
    }, EvidenceSide.Old, images);
    Assert.IsTrue(blocked.Succeeded, blocked.Message);
    Assert.IsGreaterThan(5, blocked.Steps[0].Plan.StartRow,
      "An occupied starting cell must be preserved by placing below it.");
    Assert.AreEqual(blocked.Steps[0].Plan.StartRow, blocked.Steps[0].Pair!.TargetStartRow,
      "The existing reference and the new image must move to the same start row.");
  }

  [TestMethod]
  public void AnalyzeSnapshot_TallBackfillReservesRowsBeforeOppositeSecondImage()
  {
    var snapshot = Snapshot(FixtureLoader.LoadLayout("default-final-case.json") with { ActiveRow = 3 }) with
    {
      Shapes = [new SnapshotShape("new-1", 5, 10, 4, 10, true)
        { WidthPoints = 120, HeightPoints = 60, SourceDimensions = new ImageDimensions(120, 60) },
        new SnapshotShape("new-2", 15, 20, 4, 10, true)],
    };
    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(snapshot, EvidenceSide.Old,
      [new AutomaticPlacementImage("old.png", new ImageDimensions(120, 240))]);
    Assert.IsTrue(result.Succeeded, result.Message);
    var plan = result.Steps[0].Plan;
    Assert.AreEqual(5, plan.StartRow);
    var insertion = plan.Insertions.Single(row => row.AtRow == 15);
    Assert.IsGreaterThan(plan.EndRow + 2, 15 + insertion.Count);
  }

  [TestMethod]
  public void ManagedShapeMetadata_RoundTripsScaleAndKeepsLegacyFormatReadable()
  {
    var modern = new ManagedShapeMetadata(1, EvidenceSide.New, new CellReference(5, 4))
    {
      SourceDimensions = new ImageDimensions(800, 400),
      AppliedScale = 0.5,
    };
    Assert.IsTrue(ManagedShapeMetadata.TryParse(modern.Serialize(), out var parsed));
    Assert.AreEqual(new ImageDimensions(800, 400), parsed.SourceDimensions);
    Assert.AreEqual(0.5, parsed.AppliedScale!.Value, 0.001);

    var legacy = new ManagedShapeMetadata(1, EvidenceSide.Old, new CellReference(5, 19)).Serialize();
    Assert.IsTrue(ManagedShapeMetadata.TryParse(legacy, out var legacyParsed));
    Assert.IsNull(legacyParsed.SourceDimensions);
    Assert.IsNull(legacyParsed.AppliedScale);
  }

  [TestMethod]
  public void ContentOccupancy_IgnoresShapeOutsideEvidenceColumns_ButKeepsCrossSideShape()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json");
    var snapshot = Snapshot(signals) with
    {
      Shapes =
      [
        new SnapshotShape("outside", 60, 61, 1, 2, false),
        new SnapshotShape("cross-side", 62, 63, 17, 18, false),
      ],
    };
    var layout = new CaseLayoutAnalyzer().Analyze(signals).Layout!;

    var spans = new ContentOccupancyAnalyzer().Analyze(snapshot, layout);

    Assert.HasCount(1, spans);
    Assert.AreEqual("cross-side", snapshot.Shapes[1].Name);
    Assert.IsNull(spans[0].Side);
    Assert.AreEqual(62, spans[0].StartRow);
  }

  [TestMethod]
  public void CompletedCase_RequiresManagedImageOnEachSide()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json") with { ActiveRow = 3 };
    var layout = new CaseLayoutAnalyzer().Analyze(signals).Layout!;
    var newOnly = Snapshot(signals) with
    {
      Shapes =
      [
        new SnapshotShape("new", 5, 10, 4, 10, true),
        new SnapshotShape("unmanaged-old", 5, 10, 19, 25, false),
      ],
    };
    var bothSides = newOnly with
    {
      Shapes = newOnly.Shapes.Append(new SnapshotShape("old", 5, 10, 19, 25, true)).ToArray(),
    };

    Assert.IsFalse(ExcelCaseMaintenanceService.HasBothSides(newOnly, layout));
    Assert.IsTrue(ExcelCaseMaintenanceService.HasBothSides(bothSides, layout));
  }

  [TestMethod]
  public void AnalyzeSnapshot_ReportsWhetherPlacementCanCompleteCase()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json") with { ActiveRow = 3 };
    var image = new[] { new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60)) };
    var withoutOpposite = Snapshot(signals);
    var withOld = withoutOpposite with
    {
      Shapes = [new SnapshotShape("old", 5, 10, 19, 25, true)],
    };

    var incomplete = new ExcelAutomaticPlacementService().AnalyzeSnapshot(withoutOpposite, EvidenceSide.New, image);
    var complete = new ExcelAutomaticPlacementService().AnalyzeSnapshot(withOld, EvidenceSide.New, image);

    Assert.IsFalse(incomplete.CompletesCaseAfterPlacement);
    Assert.IsTrue(complete.CompletesCaseAfterPlacement);
  }

  [TestMethod]
  public void CaseNavigation_ResolvesConfirmedPreviousAndNextAnchors()
  {
    var source = FixtureLoader.LoadLayout("default-final-case.json");
    var signals = source with
    {
      ActiveRow = 60,
      Anchors = source.Anchors.Append(new CaseAnchorSignal(103, false, true, true, null, "3")).ToArray(),
    };

    var previous = ExcelCaseNavigationService.ResolveTargetRow(signals, CaseNavigationDirection.Previous);
    var next = ExcelCaseNavigationService.ResolveTargetRow(signals, CaseNavigationDirection.Next);

    Assert.AreEqual(3, previous);
    Assert.AreEqual(103, next);
  }

  [TestMethod]
  public void CaseNavigation_CrossesSheetsInTabOrderWithoutWrapping()
  {
    string[] names = ["表紙", "B10", "帳票", "B2"];
    CollectionAssert.AreEqual(new[] { "帳票" },
      ExcelCaseNavigationService.AdjacentWorksheetNames(names, "B10", CaseNavigationDirection.Next).ToArray());
    CollectionAssert.AreEqual(new[] { "B10" },
      ExcelCaseNavigationService.AdjacentWorksheetNames(names, "帳票", CaseNavigationDirection.Previous).ToArray());
    Assert.HasCount(0, ExcelCaseNavigationService.AdjacentWorksheetNames(names, "B2", CaseNavigationDirection.Next));
    Assert.HasCount(0, ExcelCaseNavigationService.AdjacentWorksheetNames(names, "表紙", CaseNavigationDirection.Previous));
  }

  private static SheetSnapshot Snapshot(SheetLayoutSignals signals) =>
    new(
      "Evidence",
      new CellReference(signals.ActiveRow, 3),
      1,
      signals.RawUsedLastRow,
      1,
      signals.ObservedLastEvidenceColumn ?? 32,
      false,
      false,
      signals,
      [],
      [],
      Enumerable.Range(1, signals.RawUsedLastRow).ToDictionary(row => row, _ => 15.0),
      Enumerable.Range(3, 30).ToDictionary(column => column, _ => 20.0));
}
