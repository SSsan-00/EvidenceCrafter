using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace EvidenceCrafter.App;

/// <summary>Windows-style, notification-free rectangular capture over the virtual desktop.</summary>
internal sealed class ScreenCaptureDialog : Form
{
  private readonly Bitmap desktop;
  private Point selectionStart;
  private Point selectionEnd;
  private bool selecting;
  private Bitmap? capturedImage;
  private Rectangle? capturedBounds;

  public ScreenCaptureDialog()
  {
    var virtualScreen = SystemInformation.VirtualScreen;
    if (virtualScreen.Width < 1 || virtualScreen.Height < 1)
    {
      throw new InvalidOperationException("画面領域を取得できませんでした。");
    }

    desktop = new Bitmap(virtualScreen.Width, virtualScreen.Height, PixelFormat.Format32bppPArgb);
    using (var graphics = Graphics.FromImage(desktop))
    {
      graphics.CopyFromScreen(virtualScreen.Location, Point.Empty, virtualScreen.Size, CopyPixelOperation.SourceCopy);
    }

    AutoScaleMode = AutoScaleMode.None;
    BackColor = Color.Black;
    Bounds = virtualScreen;
    Cursor = Cursors.Cross;
    DoubleBuffered = true;
    FormBorderStyle = FormBorderStyle.None;
    KeyPreview = true;
    ShowInTaskbar = false;
    StartPosition = FormStartPosition.Manual;
    TopMost = true;
  }

  public Bitmap TakeCapturedImage()
  {
    if (capturedImage is null)
    {
      throw new InvalidOperationException("キャプチャ画像がありません。");
    }

    var result = capturedImage;
    capturedImage = null;
    return result;
  }

  public Rectangle CapturedBounds => capturedBounds ?? throw new InvalidOperationException("キャプチャ範囲がありません。");

  internal static Rectangle NormalizeSelection(Point first, Point second)
  {
    var left = Math.Min(first.X, second.X);
    var top = Math.Min(first.Y, second.Y);
    var right = Math.Max(first.X, second.X);
    var bottom = Math.Max(first.Y, second.Y);
    return Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
  }

  protected override void OnShown(EventArgs eventArgs)
  {
    base.OnShown(eventArgs);
    Activate();
    Focus();
  }

  protected override void OnPaint(PaintEventArgs eventArgs)
  {
    base.OnPaint(eventArgs);
    eventArgs.Graphics.DrawImageUnscaled(desktop, Point.Empty);
    using (var dim = new SolidBrush(Color.FromArgb(105, Color.Black)))
    {
      eventArgs.Graphics.FillRectangle(dim, ClientRectangle);
    }

    var selection = CurrentSelection();
    if (selection.Width > 1 && selection.Height > 1)
    {
      eventArgs.Graphics.DrawImage(desktop, selection, selection, GraphicsUnit.Pixel);
      using var outline = new Pen(Color.White, 1F) { DashStyle = DashStyle.Solid };
      eventArgs.Graphics.DrawRectangle(outline, selection.X, selection.Y, selection.Width - 1, selection.Height - 1);
      DrawSizeLabel(eventArgs.Graphics, selection);
    }

    if (!selecting)
    {
      DrawInstruction(eventArgs.Graphics);
    }
  }

  protected override void OnMouseDown(MouseEventArgs eventArgs)
  {
    base.OnMouseDown(eventArgs);
    if (eventArgs.Button == MouseButtons.Right)
    {
      CancelCapture();
      return;
    }
    if (eventArgs.Button != MouseButtons.Left)
    {
      return;
    }

    selectionStart = Clamp(eventArgs.Location);
    selectionEnd = selectionStart;
    selecting = true;
    Capture = true;
    Invalidate();
  }

  protected override void OnMouseMove(MouseEventArgs eventArgs)
  {
    base.OnMouseMove(eventArgs);
    if (!selecting)
    {
      return;
    }

    selectionEnd = Clamp(eventArgs.Location);
    Invalidate();
  }

  protected override void OnMouseUp(MouseEventArgs eventArgs)
  {
    base.OnMouseUp(eventArgs);
    if (!selecting || eventArgs.Button != MouseButtons.Left)
    {
      return;
    }

    selectionEnd = Clamp(eventArgs.Location);
    selecting = false;
    Capture = false;
    var selection = CurrentSelection();
    if (selection.Width < 2 || selection.Height < 2)
    {
      Invalidate();
      return;
    }

    capturedImage = desktop.Clone(selection, PixelFormat.Format32bppPArgb);
    selection.Offset(Bounds.Location);
    capturedBounds = selection;
    DialogResult = DialogResult.OK;
    Close();
  }

  protected override bool ProcessCmdKey(ref Message message, Keys keyData)
  {
    if (keyData == Keys.Escape)
    {
      CancelCapture();
      return true;
    }

    return base.ProcessCmdKey(ref message, keyData);
  }

  protected override void Dispose(bool disposing)
  {
    if (disposing)
    {
      capturedImage?.Dispose();
      desktop.Dispose();
    }
    base.Dispose(disposing);
  }

  private Rectangle CurrentSelection() => NormalizeSelection(selectionStart, selectionEnd);

  private Point Clamp(Point point) => new(
    Math.Clamp(point.X, 0, Math.Max(0, ClientSize.Width - 1)),
    Math.Clamp(point.Y, 0, Math.Max(0, ClientSize.Height - 1)));

  private void CancelCapture()
  {
    DialogResult = DialogResult.Cancel;
    Close();
  }

  private void DrawInstruction(Graphics graphics)
  {
    const string instruction = "切り取る領域をドラッグしてください   Esc / 右クリック: キャンセル";
    using var font = new Font("Meiryo UI", 11F, FontStyle.Bold);
    var size = graphics.MeasureString(instruction, font);
    var bounds = new RectangleF(
      Math.Max(8F, (ClientSize.Width - size.Width) / 2F - 12F),
      20F,
      size.Width + 24F,
      size.Height + 14F);
    using var background = new SolidBrush(Color.FromArgb(220, 32, 32, 32));
    using var foreground = new SolidBrush(Color.White);
    graphics.FillRectangle(background, bounds);
    graphics.DrawString(instruction, font, foreground, bounds.X + 12F, bounds.Y + 7F);
  }

  private static void DrawSizeLabel(Graphics graphics, Rectangle selection)
  {
    var text = $"{selection.Width} × {selection.Height}";
    using var font = new Font("Segoe UI", 9F, FontStyle.Regular);
    var size = graphics.MeasureString(text, font);
    var x = selection.Left;
    var y = selection.Top >= size.Height + 8F
      ? selection.Top - size.Height - 6F
      : selection.Bottom + 4F;
    using var background = new SolidBrush(Color.FromArgb(220, 20, 20, 20));
    using var foreground = new SolidBrush(Color.White);
    graphics.FillRectangle(background, x, y, size.Width + 8F, size.Height + 2F);
    graphics.DrawString(text, font, foreground, x + 4F, y + 1F);
  }
}
