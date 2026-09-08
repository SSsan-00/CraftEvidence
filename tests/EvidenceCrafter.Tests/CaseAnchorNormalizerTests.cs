using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class CaseAnchorNormalizerTests
{
  [TestMethod]
  public void Normalize_InheritsOnlyConfirmedMajorAndIgnoresText()
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
    CollectionAssert.AreEqual(new[] { "1", "1", "1" }, result.Select(anchor => anchor.ColumnAValue).ToArray());
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
