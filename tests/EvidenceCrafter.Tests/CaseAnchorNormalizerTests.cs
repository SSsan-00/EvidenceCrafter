using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class CaseAnchorNormalizerTests
{
  [TestMethod]
  public void Normalize_InheritsNumericMajorWithoutMinorAndIgnoresText()
  {
    CaseAnchorSignal[] anchors =
    [
      new(3, true, true, false, "１", "１"),
      new(10, true, true, true, "Description", "２"),
      new(20, false, true, false, null, "2"),
      new(25, true, false, false, "9", null),
      new(30, false, true, false, null, "3"),
      new(40, false, true, true, null, "Notes"),
    ];
    var result = CaseAnchorNormalizer.Normalize(anchors);
    CollectionAssert.AreEqual(new[] { 3, 20, 30 }, result.Select(anchor => anchor.Row).ToArray());
    CollectionAssert.AreEqual(new[] { "1", "1", "9" }, result.Select(anchor => anchor.ColumnAValue).ToArray());
  }

  [TestMethod]
  public void Normalize_UsesMinorRowForSplitNumberingAndUpdatesMajor()
  {
    CaseAnchorSignal[] anchors =
    [
      new(2, false, true, false, null, "1"),
      new(5, true, false, false, "１", null),
      new(6, false, true, false, null, "１"),
      new(15, false, true, false, null, "2"),
      new(25, true, false, false, "2", null),
      new(26, false, true, false, null, "1"),
      new(40, true, true, false, "3", "1"),
    ];
    var result = CaseAnchorNormalizer.Normalize(anchors.Reverse());
    CollectionAssert.AreEqual(new[] { 6, 15, 26, 40 }, result.Select(anchor => anchor.Row).ToArray());
    CollectionAssert.AreEqual(new[] { "1-1", "1-2", "2-1", "3-1" },
      result.Select(anchor => $"{anchor.ColumnAValue}-{anchor.ColumnBValue}").ToArray());
    CollectionAssert.AreEqual(result, CaseAnchorNormalizer.Normalize(result));
  }

  [TestMethod]
  public void NormalizeCaseLabel_RequiresNumericPairAndAcceptsUnicodeHyphens()
  {
    Assert.AreEqual("1-2", CaseAnchorNormalizer.NormalizeCaseLabel(" １－２ "));
    Assert.AreEqual("1-2", CaseAnchorNormalizer.NormalizeCaseLabel("1—2"));
    Assert.IsNull(CaseAnchorNormalizer.NormalizeCaseLabel("2"));
    Assert.IsNull(CaseAnchorNormalizer.NormalizeCaseLabel("Case-2"));
    Assert.IsNull(CaseAnchorNormalizer.NormalizeCaseLabel("1-2-3"));
  }
}
