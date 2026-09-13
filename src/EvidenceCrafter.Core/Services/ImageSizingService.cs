using EvidenceCrafter.Core.Models;

namespace EvidenceCrafter.Core.Services;

public sealed class ImageSizingService
{
  public FittedImage AtScale(ImageDimensions image, double width, double scale)
  {
    _ = FitToWidth(image, width);
    if (!double.IsFinite(scale) || scale <= 0 || image.WidthPoints * scale > width + 0.01)
      throw new ArgumentOutOfRangeException(nameof(scale), "The selected scale exceeds the available width.");
    return new(image.WidthPoints * scale, image.HeightPoints * scale, scale);
  }

  /// <summary>Fits an image to the available width while preventing enlargement.</summary>
  public FittedImage FitToWidth(ImageDimensions image, double availableWidthPoints)
  {
    if (!double.IsFinite(image.WidthPoints) ||
      !double.IsFinite(image.HeightPoints) ||
      image.WidthPoints <= 0 ||
      image.HeightPoints <= 0)
    {
      throw new ArgumentOutOfRangeException(nameof(image), "Image dimensions must be positive.");
    }

    if (!double.IsFinite(availableWidthPoints) || availableWidthPoints <= 0)
    {
      throw new ArgumentOutOfRangeException(
        nameof(availableWidthPoints),
        "Available width must be positive.");
    }

    var scale = Math.Min(1.0, availableWidthPoints / image.WidthPoints);
    return new FittedImage(image.WidthPoints * scale, image.HeightPoints * scale, scale);
  }
}
