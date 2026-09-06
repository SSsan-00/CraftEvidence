using CraftEvidence.Core.Services;

namespace CraftEvidence.Tests;

[TestClass]
public sealed class ScreenImageSizingTests
{
  [TestMethod]
  public void FromPixels_UsesWindowsLogicalSize()
  {
    var result = ScreenImageSizing.FromPixels(800, 400);

    Assert.AreEqual(600, result.WidthPoints);
    Assert.AreEqual(300, result.HeightPoints);
  }
}
