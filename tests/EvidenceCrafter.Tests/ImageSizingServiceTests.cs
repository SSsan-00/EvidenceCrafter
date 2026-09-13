using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class ImageSizingServiceTests
{
  private readonly ImageSizingService service = new();

  [TestMethod]
  public void FitToWidth_WideImage_ShrinksAndPreservesAspectRatio()
  {
    var result = service.FitToWidth(new ImageDimensions(800, 400), 400);

    Assert.AreEqual(400, result.WidthPoints, 0.001);
    Assert.AreEqual(200, result.HeightPoints, 0.001);
    Assert.AreEqual(0.5, result.Scale, 0.001);
  }

  [TestMethod]
  public void FitToWidth_SmallImage_DoesNotEnlarge()
  {
    var result = service.FitToWidth(new ImageDimensions(200, 100), 400);

    Assert.AreEqual(200, result.WidthPoints, 0.001);
    Assert.AreEqual(100, result.HeightPoints, 0.001);
    Assert.AreEqual(1.0, result.Scale, 0.001);
  }

  [TestMethod]
  public void FitToWidth_InvalidImage_Throws()
  {
    Assert.ThrowsExactly<ArgumentOutOfRangeException>(() =>
      service.FitToWidth(new ImageDimensions(0, 100), 400));
  }

  [TestMethod]
  public void FitToWidth_NonFiniteImage_Throws()
  {
    Assert.ThrowsExactly<ArgumentOutOfRangeException>(() =>
      service.FitToWidth(new ImageDimensions(100, double.PositiveInfinity), 200));
    Assert.ThrowsExactly<ArgumentOutOfRangeException>(() =>
      service.FitToWidth(new ImageDimensions(double.NaN, 100), 200));
  }

  [TestMethod]
  public void AtScale_AllowsCommonPairEnlargementWithinWidth()
  {
    var result = service.AtScale(new ImageDimensions(100, 50), 268, 2.0);

    Assert.AreEqual(200, result.WidthPoints, 0.001);
    Assert.AreEqual(100, result.HeightPoints, 0.001);
    Assert.AreEqual(2.0, result.Scale, 0.001);
  }

  [TestMethod]
  public void AtScale_RejectsScaleThatExceedsAvailableWidth()
  {
    Assert.ThrowsExactly<ArgumentOutOfRangeException>(() =>
      service.AtScale(new ImageDimensions(100, 50), 268, 2.69));
  }
}
