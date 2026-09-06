using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace CraftEvidence.App;

internal sealed class ImageEditDocument : IDisposable
{
  private const int HistoryLimit = 20;
  private readonly Bitmap original;
  private readonly List<ImageState> undoStates = [];
  private readonly List<ImageState> redoStates = [];
  private Bitmap current;
  private int currentStateId;
  private int nextStateId = 1;
  private bool disposed;

  public ImageEditDocument(Image image)
  {
    ArgumentNullException.ThrowIfNull(image);
    if (image.Width <= 0 || image.Height <= 0)
    {
      throw new ArgumentException("画像のサイズが不正です。", nameof(image));
    }

    original = CopyBitmap(image);
    current = CopyBitmap(image);
  }

  public event EventHandler? Changed;

  public int Width
  {
    get
    {
      ObjectDisposedException.ThrowIf(disposed, this);
      return current.Width;
    }
  }

  public int Height
  {
    get
    {
      ObjectDisposedException.ThrowIf(disposed, this);
      return current.Height;
    }
  }

  public bool CanUndo => !disposed && undoStates.Count > 0;

  public bool CanRedo => !disposed && redoStates.Count > 0;

  public bool HasChanges => !disposed && currentStateId != 0;

  internal Image CurrentImage
  {
    get
    {
      ObjectDisposedException.ThrowIf(disposed, this);
      return current;
    }
  }

  public Bitmap GetImageCopy()
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    return CopyBitmap(current);
  }

  public bool DrawRectangle(Rectangle bounds, Color? color = null)
  {
    if (!TryClip(bounds, out var clipped))
    {
      return false;
    }

    return Edit(next =>
    {
      using var graphics = Graphics.FromImage(next);
      graphics.SmoothingMode = SmoothingMode.AntiAlias;
      var stroke = StrokeWidth(next);
      using var pen = new Pen(color ?? Color.Red, stroke);
      var inset = stroke / 2F;
      graphics.DrawRectangle(
        pen,
        clipped.X + inset,
        clipped.Y + inset,
        Math.Max(0F, clipped.Width - stroke),
        Math.Max(0F, clipped.Height - stroke));
    });
  }

  public bool DrawArrow(Point start, Point end, Color? color = null)
  {
    start = Clamp(start);
    end = Clamp(end);
    if (start == end)
    {
      return false;
    }

    return Edit(next =>
    {
      using var graphics = Graphics.FromImage(next);
      graphics.SmoothingMode = SmoothingMode.AntiAlias;
      using var pen = new Pen(color ?? Color.Red, StrokeWidth(next));
      using var arrowCap = new AdjustableArrowCap(5F, 5F, true);
      pen.CustomEndCap = arrowCap;
      graphics.DrawLine(pen, start, end);
    });
  }

  public bool DrawText(string text, Point location, Color? color = null)
  {
    if (string.IsNullOrWhiteSpace(text))
    {
      return false;
    }

    location = Clamp(location);
    return Edit(next =>
    {
      using var graphics = Graphics.FromImage(next);
      graphics.SmoothingMode = SmoothingMode.AntiAlias;
      graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
      var fontSize = Math.Max(12F, Math.Min(next.Width, next.Height) / 25F);
      using var font = new Font(FontFamily.GenericSansSerif, fontSize, FontStyle.Bold, GraphicsUnit.Pixel);
      var measured = graphics.MeasureString(text.Trim(), font);
      var x = Math.Min(location.X, Math.Max(0F, next.Width - measured.Width - 6F));
      var y = Math.Min(location.Y, Math.Max(0F, next.Height - measured.Height - 4F));
      var labelColor = color ?? Color.Red;
      var labelBounds = new RectangleF(x, y, measured.Width + 6F, measured.Height + 4F);
      using var background = new SolidBrush(Color.FromArgb(210, Color.White));
      using var foreground = new SolidBrush(labelColor);
      using var border = new Pen(labelColor, Math.Max(1F, StrokeWidth(next) / 2F));
      graphics.FillRectangle(background, labelBounds);
      graphics.DrawRectangle(border, labelBounds.X, labelBounds.Y, labelBounds.Width, labelBounds.Height);
      graphics.DrawString(text.Trim(), font, foreground, x + 3F, y + 2F);
    });
  }

  public bool Mosaic(Rectangle bounds, int blockSize = 12)
  {
    if (blockSize < 2)
    {
      throw new ArgumentOutOfRangeException(nameof(blockSize));
    }

    if (!TryClip(bounds, out var clipped))
    {
      return false;
    }

    return Edit(next =>
    {
      var sampleWidth = Math.Max(1, (int)Math.Ceiling(clipped.Width / (double)blockSize));
      var sampleHeight = Math.Max(1, (int)Math.Ceiling(clipped.Height / (double)blockSize));
      using var sample = new Bitmap(sampleWidth, sampleHeight, PixelFormat.Format32bppPArgb);
      using (var sampleGraphics = Graphics.FromImage(sample))
      {
        sampleGraphics.InterpolationMode = InterpolationMode.HighQualityBilinear;
        sampleGraphics.DrawImage(
          next,
          new Rectangle(0, 0, sampleWidth, sampleHeight),
          clipped,
          GraphicsUnit.Pixel);
      }

      using var graphics = Graphics.FromImage(next);
      graphics.CompositingMode = CompositingMode.SourceCopy;
      graphics.InterpolationMode = InterpolationMode.NearestNeighbor;
      graphics.PixelOffsetMode = PixelOffsetMode.Half;
      graphics.DrawImage(sample, clipped);
    });
  }

  public bool Crop(Rectangle bounds)
  {
    if (!TryClip(bounds, out var clipped) || clipped.Width < 2 || clipped.Height < 2)
    {
      return false;
    }

    ObjectDisposedException.ThrowIf(disposed, this);
    var cropped = new Bitmap(clipped.Width, clipped.Height, PixelFormat.Format32bppPArgb);
    PreserveResolution(current, cropped);
    try
    {
      using (var graphics = Graphics.FromImage(cropped))
      {
        graphics.CompositingMode = CompositingMode.SourceCopy;
        graphics.DrawImage(
          current,
          new Rectangle(0, 0, cropped.Width, cropped.Height),
          clipped,
          GraphicsUnit.Pixel);
      }

      Commit(cropped, nextStateId++);
      cropped = null!;
      return true;
    }
    finally
    {
      cropped?.Dispose();
    }
  }

  public bool Undo()
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    if (undoStates.Count == 0)
    {
      return false;
    }

    redoStates.Add(new ImageState(current, currentStateId));
    var previous = TakeLast(undoStates);
    current = previous.Bitmap;
    currentStateId = previous.StateId;
    Changed?.Invoke(this, EventArgs.Empty);
    return true;
  }

  public bool Redo()
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    if (redoStates.Count == 0)
    {
      return false;
    }

    undoStates.Add(new ImageState(current, currentStateId));
    var next = TakeLast(redoStates);
    current = next.Bitmap;
    currentStateId = next.StateId;
    Changed?.Invoke(this, EventArgs.Empty);
    return true;
  }

  public bool Reset()
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    if (!HasChanges)
    {
      return false;
    }

    Commit(CopyBitmap(original), 0);
    return true;
  }

  public void Dispose()
  {
    if (disposed)
    {
      return;
    }

    disposed = true;
    current.Dispose();
    original.Dispose();
    DisposeStates(undoStates);
    DisposeStates(redoStates);
  }

  private static ImageState TakeLast(List<ImageState> states)
  {
    var index = states.Count - 1;
    var state = states[index];
    states.RemoveAt(index);
    return state;
  }

  private static void DisposeStates(List<ImageState> states)
  {
    foreach (var state in states)
    {
      state.Bitmap.Dispose();
    }

    states.Clear();
  }

  private static float StrokeWidth(Image image) =>
    Math.Max(2F, Math.Min(image.Width, image.Height) / 250F);

  private static Bitmap CopyBitmap(Image source)
  {
    var copy = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppPArgb);
    PreserveResolution(source, copy);
    using var graphics = Graphics.FromImage(copy);
    graphics.CompositingMode = CompositingMode.SourceCopy;
    graphics.DrawImageUnscaled(source, 0, 0);
    return copy;
  }

  private static void PreserveResolution(Image source, Bitmap target)
  {
    if (source.HorizontalResolution > 0 && source.VerticalResolution > 0)
    {
      target.SetResolution(source.HorizontalResolution, source.VerticalResolution);
    }
  }

  private Point Clamp(Point point) => new(
    Math.Clamp(point.X, 0, current.Width - 1),
    Math.Clamp(point.Y, 0, current.Height - 1));

  private bool TryClip(Rectangle bounds, out Rectangle clipped)
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    var normalized = Normalize(bounds);
    clipped = Rectangle.Intersect(normalized, new Rectangle(Point.Empty, current.Size));
    return clipped.Width > 0 && clipped.Height > 0;
  }

  private static Rectangle Normalize(Rectangle bounds)
  {
    var left = Math.Min(bounds.Left, bounds.Right);
    var top = Math.Min(bounds.Top, bounds.Bottom);
    var right = Math.Max(bounds.Left, bounds.Right);
    var bottom = Math.Max(bounds.Top, bounds.Bottom);
    return Rectangle.FromLTRB(left, top, right, bottom);
  }

  private bool Edit(Action<Bitmap> draw)
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    var next = CopyBitmap(current);
    try
    {
      draw(next);
      Commit(next, nextStateId++);
      next = null!;
      return true;
    }
    finally
    {
      next?.Dispose();
    }
  }

  private void Commit(Bitmap next, int stateId)
  {
    undoStates.Add(new ImageState(current, currentStateId));
    if (undoStates.Count > HistoryLimit)
    {
      undoStates[0].Bitmap.Dispose();
      undoStates.RemoveAt(0);
    }

    DisposeStates(redoStates);
    current = next;
    currentStateId = stateId;
    Changed?.Invoke(this, EventArgs.Empty);
  }

  private sealed record ImageState(Bitmap Bitmap, int StateId);
}
