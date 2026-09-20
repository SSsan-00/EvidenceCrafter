using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class ExcelSheetSnapshotServiceTests
{
  [TestMethod]
  public void ResolveScopeRow_ReturnsOnlyUniqueNormalizedCase()
  {
    CaseAnchorSignal[] anchors =
    [
      new(3, true, true, false, "１", "１"),
      new(33, false, true, false, null, "2"),
      new(63, true, true, false, "2", "1"),
    ];

    Assert.AreEqual(33, ExcelSheetSnapshotService.ResolveScopeRow(anchors, "1-2"));
    Assert.IsNull(ExcelSheetSnapshotService.ResolveScopeRow(anchors, "invalid"));

    anchors = [.. anchors, new(93, true, true, false, "1", "2")];
    Assert.IsNull(ExcelSheetSnapshotService.ResolveScopeRow(anchors, "1-2"));
  }
}
