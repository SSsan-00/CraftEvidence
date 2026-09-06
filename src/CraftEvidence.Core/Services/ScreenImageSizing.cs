using CraftEvidence.Core.Models;

namespace CraftEvidence.Core.Services;

public static class ScreenImageSizing
{
  private const double WindowsDpi = 96.0;

  public static ImageDimensions FromPixels(int width, int height)
  {
    if (width <= 0 || height <= 0)
    {
      throw new ArgumentOutOfRangeException(nameof(width), "Image dimensions must be positive.");
    }

    return new ImageDimensions(width * 72.0 / WindowsDpi, height * 72.0 / WindowsDpi);
  }
}
