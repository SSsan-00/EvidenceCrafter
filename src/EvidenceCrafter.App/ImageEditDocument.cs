using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace EvidenceCrafter.App;

internal sealed class ImageEditDocument : IDisposable
{
  private const int HistoryLimit = 20;
  private readonly Bitmap original;
  private readonly List<ImageState> undoStates = [];
  private readonly List<ImageState> redoStates = [];
  private IReadOnlyList<TextAnnotation> textAnnotations = [];
  private IReadOnlyList<RectangleAnnotation> rectangleAnnotations = [];
  private SharedBitmap current;
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
    current = new SharedBitmap(CopyBitmap(image));
  }

  public event EventHandler? Changed;

  public int Width
  {
    get
    {
      ObjectDisposedException.ThrowIf(disposed, this);
      return current.Bitmap.Width;
    }
  }

  public int Height
  {
    get
    {
      ObjectDisposedException.ThrowIf(disposed, this);
      return current.Bitmap.Height;
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
      return current.Bitmap;
    }
  }

  public Bitmap GetImageCopy()
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    return RenderCurrent();
  }

  public bool DrawRectangle(Rectangle bounds, Color? color = null)
  {
    if (!TryClip(bounds, out var clipped) || clipped.Width < 2 || clipped.Height < 2)
    {
      return false;
    }

    CommitAnnotations(nextStateId++, textAnnotations,
      [.. rectangleAnnotations, new RectangleAnnotation(Guid.NewGuid(), clipped, color ?? Color.Red)]);
    return true;
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
    }, preserveTextAnnotations: true);
  }

  public bool DrawText(string text, Point location, Color? color = null)
  {
    if (string.IsNullOrWhiteSpace(text))
    {
      return false;
    }

    var annotation = CreateTextAnnotation(Guid.NewGuid(), text.Trim(), location, color ?? Color.Red);
    CommitText(nextStateId++, [.. textAnnotations, annotation]);
    return true;
  }

  public bool TryGetTextAt(Point location, out Guid annotationId)
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    for (var index = textAnnotations.Count - 1; index >= 0; index--)
    {
      if (textAnnotations[index].Bounds.Contains(location))
      {
        annotationId = textAnnotations[index].Id;
        return true;
      }
    }

    annotationId = Guid.Empty;
    return false;
  }

  public bool TryGetMovedTextAnnotation(Guid annotationId, Point location, out TextAnnotation annotation)
  {
    if (!TryGetTextAnnotation(annotationId, out var existing))
    {
      annotation = default!;
      return false;
    }

    var moved = CreateTextAnnotation(existing.Id, existing.Text, location, existing.Color, existing.FontSize);
    annotation = moved with
    {
      MaskedRegions = existing.MaskedRegions.Select(region => new Rectangle(
        region.X + moved.Location.X - existing.Location.X,
        region.Y + moved.Location.Y - existing.Location.Y,
        region.Width, region.Height)).ToArray(),
    };
    return true;
  }

  public bool TryGetResizedTextAnnotation(Guid annotationId, float fontSize, out TextAnnotation annotation)
  {
    if (!TryGetTextAnnotation(annotationId, out var existing))
    {
      annotation = default!;
      return false;
    }

    annotation = CreateTextAnnotation(existing.Id, existing.Text, existing.Location, existing.Color, fontSize)
      with { MaskedRegions = existing.MaskedRegions };
    return true;
  }

  public bool TryGetTextAnnotation(Guid annotationId, out TextAnnotation annotation)
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    var existing = textAnnotations.FirstOrDefault(item => item.Id == annotationId);
    if (existing is null)
    {
      annotation = default!;
      return false;
    }

    annotation = existing;
    return true;
  }

  public bool MoveText(Guid annotationId, Point location)
  {
    if (!TryGetMovedTextAnnotation(annotationId, location, out var moved))
    {
      return false;
    }

    var index = textAnnotations.ToList().FindIndex(item => item.Id == annotationId);
    if (index < 0 || textAnnotations[index].Location == moved.Location)
    {
      return false;
    }

    var nextAnnotations = textAnnotations.ToList();
    nextAnnotations[index] = moved;
    CommitText(nextStateId++, nextAnnotations);
    return true;
  }

  public bool ResizeText(Guid annotationId, float fontSize)
  {
    if (!TryGetResizedTextAnnotation(annotationId, fontSize, out var resized)) return false;
    var index = textAnnotations.ToList().FindIndex(item => item.Id == annotationId);
    if (index < 0 || Math.Abs(textAnnotations[index].FontSize - resized.FontSize) < 0.01F) return false;

    var nextAnnotations = textAnnotations.ToList();
    nextAnnotations[index] = resized;
    CommitText(nextStateId++, nextAnnotations);
    return true;
  }

  public bool UpdateText(Guid annotationId, string text)
  {
    if (string.IsNullOrWhiteSpace(text) || !TryGetTextAnnotation(annotationId, out var existing) || existing.Text == text.Trim()) return false;
    var updated = CreateTextAnnotation(existing.Id, text.Trim(), existing.Location, existing.Color, existing.FontSize)
      with { MaskedRegions = existing.MaskedRegions };
    CommitText(nextStateId++, textAnnotations.Select(item => item.Id == annotationId ? updated : item).ToArray());
    return true;
  }

  public bool DeleteText(Guid annotationId)
  {
    if (!TryGetTextAnnotation(annotationId, out _)) return false;
    CommitText(nextStateId++, textAnnotations.Where(item => item.Id != annotationId).ToArray());
    return true;
  }

  public bool TryGetRectangleAt(Point location, out Guid annotationId)
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    var tolerance = Math.Max(4, (int)Math.Ceiling(StrokeWidth(current.Bitmap) * 2));
    for (var index = rectangleAnnotations.Count - 1; index >= 0; index--)
    {
      var annotation = rectangleAnnotations[index];
      if (annotation.MaskedRegions.Any(region => region.Contains(location))) continue;
      var bounds = annotation.Bounds;
      var outer = bounds;
      outer.Inflate(tolerance, tolerance);
      var inner = bounds;
      inner.Inflate(-tolerance, -tolerance);
      if (outer.Contains(location) && (inner.Width <= 0 || inner.Height <= 0 || !inner.Contains(location)))
      {
        annotationId = annotation.Id;
        return true;
      }
    }

    annotationId = Guid.Empty;
    return false;
  }

  public bool TryGetRectangleAnnotation(Guid annotationId, out RectangleAnnotation annotation)
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    var existing = rectangleAnnotations.FirstOrDefault(item => item.Id == annotationId);
    if (existing is null)
    {
      annotation = default!;
      return false;
    }

    annotation = existing;
    return true;
  }

  public bool TryGetMovedRectangleAnnotation(Guid annotationId, Point location, out RectangleAnnotation annotation)
  {
    if (!TryGetRectangleAnnotation(annotationId, out var existing))
    {
      annotation = default!;
      return false;
    }

    var bounds = new Rectangle(
      Math.Clamp(location.X, 0, Width - existing.Bounds.Width),
      Math.Clamp(location.Y, 0, Height - existing.Bounds.Height),
      existing.Bounds.Width,
      existing.Bounds.Height);
    annotation = existing with
    {
      Bounds = bounds,
      MaskedRegions = existing.MaskedRegions.Select(region => new Rectangle(
        region.X + bounds.X - existing.Bounds.X,
        region.Y + bounds.Y - existing.Bounds.Y,
        region.Width,
        region.Height)).ToArray(),
    };
    return true;
  }

  public bool TryGetResizedRectangleAnnotation(Guid annotationId, Rectangle bounds, out RectangleAnnotation annotation)
  {
    if (!TryGetRectangleAnnotation(annotationId, out var existing) ||
      !TryClip(bounds, out var clipped) || clipped.Width < 2 || clipped.Height < 2)
    {
      annotation = default!;
      return false;
    }

    annotation = existing with { Bounds = clipped };
    return true;
  }

  public bool MoveRectangle(Guid annotationId, Point location)
  {
    if (!TryGetMovedRectangleAnnotation(annotationId, location, out var moved) ||
      moved.Bounds == rectangleAnnotations.First(item => item.Id == annotationId).Bounds) return false;
    CommitAnnotations(nextStateId++, textAnnotations,
      rectangleAnnotations.Select(item => item.Id == annotationId ? moved : item).ToArray());
    return true;
  }

  public bool ResizeRectangle(Guid annotationId, Rectangle bounds)
  {
    if (!TryGetResizedRectangleAnnotation(annotationId, bounds, out var resized) ||
      resized.Bounds == rectangleAnnotations.First(item => item.Id == annotationId).Bounds) return false;
    CommitAnnotations(nextStateId++, textAnnotations,
      rectangleAnnotations.Select(item => item.Id == annotationId ? resized : item).ToArray());
    return true;
  }

  public bool DeleteRectangle(Guid annotationId)
  {
    if (!TryGetRectangleAnnotation(annotationId, out _)) return false;
    CommitAnnotations(nextStateId++, textAnnotations,
      rectangleAnnotations.Where(item => item.Id != annotationId).ToArray());
    return true;
  }

  internal void DrawTextAnnotations(
    Graphics graphics,
    Guid? excludedAnnotationId = null,
    TextAnnotation? preview = null)
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    foreach (var annotation in textAnnotations)
    {
      if (annotation.Id != excludedAnnotationId)
      {
        DrawTextAnnotation(graphics, annotation);
      }
    }

    if (preview is not null)
    {
      DrawTextAnnotation(graphics, preview);
    }
  }

  internal void DrawRectangleAnnotations(
    Graphics graphics,
    Guid? excludedAnnotationId = null,
    RectangleAnnotation? preview = null)
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    foreach (var annotation in rectangleAnnotations)
    {
      if (annotation.Id != excludedAnnotationId) DrawRectangleAnnotation(graphics, annotation);
    }

    if (preview is not null) DrawRectangleAnnotation(graphics, preview);
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

    using var rendered = RenderCurrent();
    var next = CopyBitmap(current.Bitmap);
    try
    {
      var sampleWidth = Math.Max(1, (int)Math.Ceiling(clipped.Width / (double)blockSize));
      var sampleHeight = Math.Max(1, (int)Math.Ceiling(clipped.Height / (double)blockSize));
      using var sample = new Bitmap(sampleWidth, sampleHeight, PixelFormat.Format32bppPArgb);
      using (var sampleGraphics = Graphics.FromImage(sample))
      {
        sampleGraphics.InterpolationMode = InterpolationMode.HighQualityBilinear;
        sampleGraphics.DrawImage(
          rendered,
          new Rectangle(0, 0, sampleWidth, sampleHeight),
          clipped,
          GraphicsUnit.Pixel);
      }

      using (var graphics = Graphics.FromImage(next))
      {
        graphics.CompositingMode = CompositingMode.SourceCopy;
        graphics.InterpolationMode = InterpolationMode.NearestNeighbor;
        graphics.PixelOffsetMode = PixelOffsetMode.Half;
        graphics.DrawImage(sample, clipped);
      }
      // Bake only the masked area, and keep later text moves from uncovering it.
      Commit(next, nextStateId++, textAnnotations.Select(item => item with
      { MaskedRegions = [.. item.MaskedRegions, clipped] }).ToArray(),
        rectangleAnnotations.Select(item => item with
        { MaskedRegions = [.. item.MaskedRegions, clipped] }).ToArray());
      next = null!;
      return true;
    }
    finally { next?.Dispose(); }
  }

  public bool Crop(Rectangle bounds)
  {
    if (!TryClip(bounds, out var clipped) || clipped.Width < 2 || clipped.Height < 2)
    {
      return false;
    }

    ObjectDisposedException.ThrowIf(disposed, this);
    var cropped = new Bitmap(clipped.Width, clipped.Height, PixelFormat.Format32bppPArgb);
    PreserveResolution(current.Bitmap, cropped);
    try
    {
      using (var graphics = Graphics.FromImage(cropped))
      {
        graphics.CompositingMode = CompositingMode.SourceCopy;
        graphics.DrawImage(
          current.Bitmap,
          new Rectangle(0, 0, cropped.Width, cropped.Height),
          clipped,
          GraphicsUnit.Pixel);
      }

      var annotations = textAnnotations.Where(item => item.Bounds.IntersectsWith(clipped)).Select(item => item with
      {
        Location = new Point(item.Location.X - clipped.X, item.Location.Y - clipped.Y),
        Bounds = new RectangleF(item.Bounds.X - clipped.X, item.Bounds.Y - clipped.Y, item.Bounds.Width, item.Bounds.Height),
        MaskedRegions = item.MaskedRegions.Select(region => new Rectangle(
          region.X - clipped.X, region.Y - clipped.Y, region.Width, region.Height)).ToArray(),
      }).ToArray();
      var rectangles = rectangleAnnotations
        .Where(item => item.Bounds.IntersectsWith(clipped))
        .Select(item =>
        {
          var bounds = Rectangle.Intersect(item.Bounds, clipped);
          bounds.Offset(-clipped.X, -clipped.Y);
          return item with
          {
            Bounds = bounds,
            MaskedRegions = item.MaskedRegions.Select(region => new Rectangle(
              region.X - clipped.X, region.Y - clipped.Y, region.Width, region.Height)).ToArray(),
          };
        }).ToArray();
      Commit(cropped, nextStateId++, annotations, rectangles);
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

    redoStates.Add(new ImageState(current.Retain(), currentStateId, textAnnotations, rectangleAnnotations));
    var previous = TakeLast(undoStates);
    current.Release();
    current = previous.Bitmap;
    currentStateId = previous.StateId;
    textAnnotations = previous.TextAnnotations;
    rectangleAnnotations = previous.RectangleAnnotations;
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

    undoStates.Add(new ImageState(current.Retain(), currentStateId, textAnnotations, rectangleAnnotations));
    var next = TakeLast(redoStates);
    current.Release();
    current = next.Bitmap;
    currentStateId = next.StateId;
    textAnnotations = next.TextAnnotations;
    rectangleAnnotations = next.RectangleAnnotations;
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

    Commit(CopyBitmap(original), 0, [], []);
    return true;
  }

  public void Dispose()
  {
    if (disposed)
    {
      return;
    }

    disposed = true;
    current.Release();
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
      state.Bitmap.Release();
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
    Math.Clamp(point.X, 0, current.Bitmap.Width - 1),
    Math.Clamp(point.Y, 0, current.Bitmap.Height - 1));

  private bool TryClip(Rectangle bounds, out Rectangle clipped)
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    var normalized = Normalize(bounds);
    clipped = Rectangle.Intersect(normalized, new Rectangle(Point.Empty, current.Bitmap.Size));
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

  private bool Edit(Action<Bitmap> draw, bool preserveTextAnnotations = false)
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    var next = preserveTextAnnotations ? CopyBitmap(current.Bitmap) : RenderCurrent();
    try
    {
      draw(next);
      Commit(next, nextStateId++, preserveTextAnnotations ? textAnnotations : [], rectangleAnnotations);
      next = null!;
      return true;
    }
    finally
    {
      next?.Dispose();
    }
  }

  private Bitmap RenderCurrent()
  {
    var rendered = CopyBitmap(current.Bitmap);
    using var graphics = Graphics.FromImage(rendered);
    DrawRectangleAnnotations(graphics);
    DrawTextAnnotations(graphics);
    return rendered;
  }

  private TextAnnotation CreateTextAnnotation(Guid id, string text, Point location, Color color, float? fontSize = null)
  {
    location = Clamp(location);
    using var graphics = Graphics.FromImage(current.Bitmap);
    using var font = CreateTextFont(fontSize);
    var measured = graphics.MeasureString(text, font);
    var x = Math.Min(location.X, Math.Max(0F, current.Bitmap.Width - measured.Width - 6F));
    var y = Math.Min(location.Y, Math.Max(0F, current.Bitmap.Height - measured.Height - 4F));
    var bounds = new RectangleF(x, y, measured.Width + 6F, measured.Height + 4F);
    return new TextAnnotation(id, text, new Point((int)x, (int)y), color, bounds) { FontSize = font.Size };
  }

  private void DrawTextAnnotation(Graphics graphics, TextAnnotation annotation)
  {
    var state = graphics.Save();
    try
    {
      foreach (var region in annotation.MaskedRegions) graphics.ExcludeClip(region);
      graphics.SmoothingMode = SmoothingMode.AntiAlias;
      graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
      using var font = CreateTextFont(annotation.FontSize);
      using var background = new SolidBrush(Color.FromArgb(210, Color.White));
      using var foreground = new SolidBrush(annotation.Color);
      using var border = new Pen(annotation.Color, Math.Max(1F, StrokeWidth(current.Bitmap) / 2F));
      graphics.FillRectangle(background, annotation.Bounds);
      graphics.DrawRectangle(border, annotation.Bounds.X, annotation.Bounds.Y, annotation.Bounds.Width, annotation.Bounds.Height);
      graphics.DrawString(annotation.Text, font, foreground, annotation.Bounds.X + 3F, annotation.Bounds.Y + 2F);
    }
    finally { graphics.Restore(state); }
  }

  private void DrawRectangleAnnotation(Graphics graphics, RectangleAnnotation annotation)
  {
    var state = graphics.Save();
    try
    {
      foreach (var region in annotation.MaskedRegions) graphics.ExcludeClip(region);
      graphics.SmoothingMode = SmoothingMode.AntiAlias;
      var stroke = StrokeWidth(current.Bitmap);
      using var pen = new Pen(annotation.Color, stroke);
      var inset = stroke / 2F;
      graphics.DrawRectangle(pen,
        annotation.Bounds.X + inset,
        annotation.Bounds.Y + inset,
        Math.Max(0F, annotation.Bounds.Width - stroke),
        Math.Max(0F, annotation.Bounds.Height - stroke));
    }
    finally { graphics.Restore(state); }
  }

  private Font CreateTextFont(float? size = null) => new(
    FontFamily.GenericSansSerif,
    Math.Clamp(
      size ?? Math.Max(12F, Math.Min(current.Bitmap.Width, current.Bitmap.Height) / 25F),
      8F,
      Math.Max(12F, Math.Min(current.Bitmap.Width, current.Bitmap.Height) / 2F)),
    FontStyle.Bold,
    GraphicsUnit.Pixel);

  private void Commit(
    Bitmap next,
    int stateId,
    IReadOnlyList<TextAnnotation> nextTextAnnotations,
    IReadOnlyList<RectangleAnnotation> nextRectangleAnnotations)
  {
    PushUndoState();
    DisposeStates(redoStates);
    current.Release();
    current = new SharedBitmap(next);
    currentStateId = stateId;
    textAnnotations = nextTextAnnotations;
    rectangleAnnotations = nextRectangleAnnotations;
    Changed?.Invoke(this, EventArgs.Empty);
  }

  private void CommitText(int stateId, IReadOnlyList<TextAnnotation> nextTextAnnotations)
    => CommitAnnotations(stateId, nextTextAnnotations, rectangleAnnotations);

  private void CommitAnnotations(
    int stateId,
    IReadOnlyList<TextAnnotation> nextTextAnnotations,
    IReadOnlyList<RectangleAnnotation> nextRectangleAnnotations)
  {
    PushUndoState();
    DisposeStates(redoStates);
    currentStateId = stateId;
    textAnnotations = nextTextAnnotations;
    rectangleAnnotations = nextRectangleAnnotations;
    Changed?.Invoke(this, EventArgs.Empty);
  }

  private void PushUndoState()
  {
    undoStates.Add(new ImageState(current.Retain(), currentStateId, textAnnotations, rectangleAnnotations));
    if (undoStates.Count > HistoryLimit)
    {
      undoStates[0].Bitmap.Release();
      undoStates.RemoveAt(0);
    }
  }

  private sealed record ImageState(
    SharedBitmap Bitmap,
    int StateId,
    IReadOnlyList<TextAnnotation> TextAnnotations,
    IReadOnlyList<RectangleAnnotation> RectangleAnnotations);

  private sealed class SharedBitmap(Bitmap bitmap)
  {
    private int references = 1;

    public Bitmap Bitmap { get; } = bitmap;

    public SharedBitmap Retain()
    {
      references++;
      return this;
    }

    public void Release()
    {
      if (--references == 0) Bitmap.Dispose();
    }
  }
}

internal sealed record TextAnnotation(
  Guid Id,
  string Text,
  Point Location,
  Color Color,
  RectangleF Bounds)
{
  internal float FontSize { get; init; } = 12;
  internal IReadOnlyList<Rectangle> MaskedRegions { get; init; } = [];
}

internal sealed record RectangleAnnotation(Guid Id, Rectangle Bounds, Color Color)
{
  internal IReadOnlyList<Rectangle> MaskedRegions { get; init; } = [];
}
