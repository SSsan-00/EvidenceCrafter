using System.ComponentModel;
using System.Drawing.Drawing2D;

namespace EvidenceCrafter.App;

internal sealed class ImageEditorDialog : Form
{
  private readonly ImageEditDocument document;
  private readonly ImageEditorCanvas canvas;
  private readonly ToolStripButton undoButton;
  private readonly ToolStripButton redoButton;
  private readonly ToolStripButton resetButton;
  private readonly ThemeColorPickerButton colorButton;
  private readonly ToolStripStatusLabel statusLabel;
  private int[] customColors;
  private bool confirmed;

  public ImageEditorDialog(Image image, int[]? customColors = null)
  {
    this.customColors = customColors?.Take(16).ToArray() ?? [];
    document = new ImageEditDocument(image);
    Text = "画像編集";
    StartPosition = FormStartPosition.CenterParent;
    MinimumSize = new Size(760, 560);
    Size = new Size(1000, 760);
    AutoScaleMode = AutoScaleMode.Dpi;
    Font = new Font("Meiryo UI", 9F);
    UiTheme.StyleForm(this);
    KeyPreview = true;

    var toolStrip = new ToolStrip
    {
      GripStyle = ToolStripGripStyle.Hidden,
      Dock = DockStyle.Top,
      Padding = new Padding(6, 3, 6, 3),
      BackColor = UiTheme.SurfaceMuted,
      ForeColor = UiTheme.TextOn(UiTheme.SurfaceMuted),
    };

    var rectangleButton = AddToolButton(toolStrip, "枠", ImageEditorTool.Rectangle);
    AddToolButton(toolStrip, "矢印", ImageEditorTool.Arrow);
    AddToolButton(toolStrip, "テキスト", ImageEditorTool.Text);
    AddToolButton(toolStrip, "モザイク", ImageEditorTool.Mosaic);
    AddToolButton(toolStrip, "トリミング", ImageEditorTool.Crop);
    rectangleButton.Checked = true;
    toolStrip.Items.Add(new ToolStripSeparator());

    colorButton = new ThemeColorPickerButton
    {
      AccessibleName = "注釈の色を選択",
      BackColor = Color.Red,
    };
    UiTheme.StyleThemeColorButton(colorButton, Font);
    colorButton.BackColor = Color.Red;
    colorButton.Click += (_, _) => SelectAnnotationColor();
    var colorHost = new ToolStripControlHost(colorButton)
    {
      AutoSize = false,
      Size = new Size(34, 34),
      Padding = new Padding(3),
      ToolTipText = "枠・矢印・テキストラベルの色を選択します",
    };
    toolStrip.Items.Add(colorHost);
    toolStrip.Items.Add(new ToolStripSeparator());

    undoButton = new ToolStripButton("元に戻す")
    {
      Enabled = false,
      ToolTipText = "直前の編集を元に戻します",
    };
    undoButton.Click += (_, _) => document.Undo();
    toolStrip.Items.Add(undoButton);

    redoButton = new ToolStripButton("やり直す")
    {
      Enabled = false,
      ToolTipText = "元に戻した編集をやり直します",
    };
    redoButton.Click += (_, _) => document.Redo();
    toolStrip.Items.Add(redoButton);

    resetButton = new ToolStripButton("元画像へ戻す")
    {
      Enabled = false,
      ToolTipText = "この編集セッション開始時の画像へ戻します",
    };
    resetButton.Click += (_, _) => document.Reset();
    toolStrip.Items.Add(resetButton);

    canvas = new ImageEditorCanvas(document)
    {
      Dock = DockStyle.Fill,
      Tool = ImageEditorTool.Rectangle,
      DrawingColor = Color.Red,
      AccessibleName = "画像編集キャンバス",
    };
    canvas.TextRequested += CanvasTextRequested;
    canvas.TextEditRequested += (_, annotationId) =>
    {
      if (!document.TryGetTextAnnotation(annotationId, out var annotation)) return;
      using var dialog = new ImageTextInputDialog(annotation.Text) { TopMost = TopMost };
      if (dialog.ShowDialog(this) == DialogResult.OK) document.UpdateText(annotationId, dialog.EnteredText);
    };

    var statusStrip = new StatusStrip { SizingGrip = false, BackColor = UiTheme.SurfaceMuted };
    statusLabel = new ToolStripStatusLabel
    {
      Spring = true,
      TextAlign = ContentAlignment.MiddleLeft,
      Text = string.Empty,
      ForeColor = UiTheme.TextMuted,
    };
    canvas.ActionRejected += (_, message) => statusLabel.Text = message;
    statusStrip.Items.Add(statusLabel);

    var applyButton = new Button
    {
      Text = "編集を反映",
      AutoSize = true,
      DialogResult = DialogResult.OK,
    };
    applyButton.Click += (_, _) => confirmed = true;
    var cancelButton = new Button
    {
      Text = "キャンセル",
      AutoSize = true,
      DialogResult = DialogResult.Cancel,
    };

    var bottomPanel = new FlowLayoutPanel
    {
      Dock = DockStyle.Bottom,
      AutoSize = true,
      FlowDirection = FlowDirection.RightToLeft,
      Padding = new Padding(8),
      BackColor = UiTheme.Canvas,
    };
    UiTheme.StyleButton(applyButton, Font, primary: true);
    applyButton.MinimumSize = new Size(132, 32);
    UiTheme.StyleButton(cancelButton, Font);
    cancelButton.MinimumSize = new Size(108, 32);
    bottomPanel.Controls.Add(cancelButton);
    bottomPanel.Controls.Add(applyButton);

    Controls.Add(canvas);
    Controls.Add(bottomPanel);
    Controls.Add(statusStrip);
    Controls.Add(toolStrip);
    AcceptButton = applyButton;
    CancelButton = cancelButton;

    document.Changed += DocumentChanged;
    FormClosing += ConfirmDiscardIfNeeded;
  }

  public bool HasChanges => document.HasChanges;

  public Bitmap GetEditedImage() => document.GetImageCopy();

  public int[] CustomColors => customColors.ToArray();

  protected override bool ProcessCmdKey(ref Message message, Keys keyData)
  {
    if (keyData == (Keys.Shift | Keys.Enter))
    {
      AddTextAtImageCenter();
      return true;
    }
    if (keyData == Keys.Enter)
    {
      if (!EnterShortcut.IsRepeat(message)) AcceptButton?.PerformClick();
      return true;
    }
    if (keyData == Keys.Escape && canvas.CancelDrag()) return true;
    if (keyData == Keys.Delete && canvas.DeleteSelectedText()) return true;
    if (keyData == Keys.F2 && canvas.EditSelectedText()) return true;
    if (keyData == (Keys.Control | Keys.Z))
    {
      document.Undo();
      return true;
    }

    if (keyData == (Keys.Control | Keys.Y))
    {
      document.Redo();
      return true;
    }

    return base.ProcessCmdKey(ref message, keyData);
  }

  protected override void Dispose(bool disposing)
  {
    if (disposing)
    {
      document.Changed -= DocumentChanged;
      canvas.TextRequested -= CanvasTextRequested;
      document.Dispose();
    }

    base.Dispose(disposing);
  }

  private ToolStripButton AddToolButton(
    ToolStrip toolStrip,
    string text,
    ImageEditorTool tool)
  {
    var button = new ToolStripButton(text)
    {
      CheckOnClick = true,
      Tag = tool,
      ToolTipText = InstructionFor(tool),
    };
    button.Click += (_, _) => SelectTool(button, toolStrip, tool);
    toolStrip.Items.Add(button);
    return button;
  }

  private void SelectTool(ToolStripButton selected, ToolStrip toolStrip, ImageEditorTool tool)
  {
    foreach (var item in toolStrip.Items.OfType<ToolStripButton>())
    {
      if (item.Tag is ImageEditorTool)
      {
        item.Checked = ReferenceEquals(item, selected);
      }
    }

    canvas.Tool = tool;
    statusLabel.Text = string.Empty;
  }

  private void CanvasTextRequested(object? sender, ImageTextRequestedEventArgs eventArgs)
  {
    using var dialog = new ImageTextInputDialog() { TopMost = TopMost };
    if (dialog.ShowDialog(this) == DialogResult.OK)
    {
      document.DrawText(dialog.EnteredText, eventArgs.ImageLocation, canvas.DrawingColor);
    }
  }

  private void AddTextAtImageCenter() =>
    CanvasTextRequested(canvas, new ImageTextRequestedEventArgs(new Point(document.Width / 2, document.Height / 2)));

  private void DocumentChanged(object? sender, EventArgs eventArgs)
  {
    undoButton.Enabled = document.CanUndo;
    redoButton.Enabled = document.CanRedo;
    resetButton.Enabled = document.HasChanges;
    canvas.Invalidate();
  }

  private void SelectAnnotationColor()
  {
    using var dialog = new ColorDialog
    {
      Color = canvas.DrawingColor,
      FullOpen = true,
      AnyColor = true,
      CustomColors = customColors.ToArray(),
    };
    var result = dialog.ShowDialog(this);
    customColors = dialog.CustomColors.Take(16).ToArray();
    if (result != DialogResult.OK)
    {
      return;
    }

    canvas.DrawingColor = dialog.Color;
    colorButton.BackColor = dialog.Color;
    canvas.Invalidate();
  }

  private void ConfirmDiscardIfNeeded(object? sender, FormClosingEventArgs eventArgs)
  {
    if (confirmed || !document.HasChanges || DialogResult == DialogResult.OK)
    {
      return;
    }

    if (MessageBox.Show(
      this,
      "画像の編集内容を破棄しますか？",
      "EvidenceCrafter",
      MessageBoxButtons.YesNo,
      MessageBoxIcon.Question,
      MessageBoxDefaultButton.Button2) != DialogResult.Yes)
    {
      eventArgs.Cancel = true;
    }
  }

  private static string InstructionFor(ImageEditorTool tool) => tool switch
  {
    ImageEditorTool.Rectangle => "ドラッグした範囲へ枠を追加します。",
    ImageEditorTool.Arrow => "矢印の始点から終点までドラッグします。",
    ImageEditorTool.Text => "クリックした位置にテキストを追加します。Shift+Enterで中央に追加します。",
    ImageEditorTool.Mosaic => "隠したい範囲をドラッグします。",
    ImageEditorTool.Crop => "残したい範囲をドラッグしてトリミングします。",
    _ => string.Empty,
  };

}

internal enum ImageEditorTool
{
  Rectangle,
  Arrow,
  Text,
  Mosaic,
  Crop,
}

internal sealed class ImageEditorCanvas : Control
{
  private const int CanvasPadding = 8;
  private readonly ImageEditDocument document;
  private Point dragStartClient;
  private Point dragCurrentClient;
  private Guid? selectedTextId;
  private Guid? movingTextId;
  private Guid? resizingTextId;
  private Point textDragStartImage;
  private Point textOriginalLocation;
  private Point textPreviewLocation;
  private RectangleF textOriginalBounds;
  private float textOriginalFontSize;
  private float textPreviewFontSize;
  private bool textDragThresholdPassed;
  private bool dragging;

  public ImageEditorCanvas(ImageEditDocument document)
  {
    this.document = document ?? throw new ArgumentNullException(nameof(document));
    DoubleBuffered = true;
    BackColor = Color.FromArgb(32, 32, 32);
    Cursor = Cursors.Cross;
    SetStyle(ControlStyles.ResizeRedraw, true);
    SetStyle(ControlStyles.Selectable, true);
    TabStop = true;
    AccessibleDescription = "テキストを選択後、F2またはダブルクリックで編集、×で削除。ドラッグで移動、右下のハンドルで拡大縮小。";
  }

  public event EventHandler<ImageTextRequestedEventArgs>? TextRequested;
  public event EventHandler<Guid>? TextEditRequested;

  public event EventHandler<string>? ActionRejected;

  [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
  public ImageEditorTool Tool { get; set; }

  [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
  public Color DrawingColor { get; set; } = Color.Red;

  protected override void OnPaint(PaintEventArgs eventArgs)
  {
    base.OnPaint(eventArgs);
    var imageBounds = GetImageBounds();
    if (imageBounds.Width <= 0 || imageBounds.Height <= 0)
    {
      return;
    }

    eventArgs.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
    eventArgs.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
    eventArgs.Graphics.DrawImage(document.CurrentImage, imageBounds);

    TextAnnotation? preview = null;
    if (movingTextId is Guid movingId &&
      !document.TryGetMovedTextAnnotation(movingId, textPreviewLocation, out preview))
    {
      movingTextId = null;
    }
    else if (resizingTextId is Guid resizingId &&
      !document.TryGetResizedTextAnnotation(resizingId, textPreviewFontSize, out preview))
    {
      resizingTextId = null;
    }

    var savedState = eventArgs.Graphics.Save();
    eventArgs.Graphics.TranslateTransform(imageBounds.X, imageBounds.Y);
    eventArgs.Graphics.ScaleTransform(
      imageBounds.Width / (float)document.Width,
      imageBounds.Height / (float)document.Height);
    document.DrawTextAnnotations(eventArgs.Graphics, movingTextId ?? resizingTextId, preview);
    DrawSelectedTextBounds(eventArgs.Graphics, preview);
    eventArgs.Graphics.Restore(savedState);
    var deleteBounds = GetDeleteBounds();
    if (!deleteBounds.IsEmpty)
    {
      using var brush = new SolidBrush(Color.Firebrick);
      eventArgs.Graphics.FillRectangle(brush, deleteBounds);
      TextRenderer.DrawText(eventArgs.Graphics, "×", Font, deleteBounds, Color.White,
        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
    }
    var resizeBounds = GetResizeBounds();
    if (!resizeBounds.IsEmpty)
    {
      using var brush = new SolidBrush(Color.DeepSkyBlue);
      eventArgs.Graphics.FillRectangle(brush, resizeBounds);
      eventArgs.Graphics.DrawRectangle(Pens.White, resizeBounds);
    }

    if (!dragging || movingTextId is not null || resizingTextId is not null || Tool is ImageEditorTool.Text)
    {
      return;
    }

    eventArgs.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
    using var pen = new Pen(
      Tool is ImageEditorTool.Rectangle or ImageEditorTool.Arrow ? DrawingColor : Color.DeepSkyBlue,
      Math.Max(1.5F, DeviceDpi / 96F * 2F))
    {
      DashStyle = Tool is ImageEditorTool.Mosaic or ImageEditorTool.Crop
        ? DashStyle.Dash
        : DashStyle.Solid,
    };
    if (Tool == ImageEditorTool.Arrow)
    {
      using var arrowCap = new AdjustableArrowCap(5F, 5F, true);
      pen.CustomEndCap = arrowCap;
      eventArgs.Graphics.DrawLine(pen, dragStartClient, dragCurrentClient);
    }
    else
    {
      eventArgs.Graphics.DrawRectangle(pen, Normalize(dragStartClient, dragCurrentClient));
    }
  }

  protected override void OnMouseDown(MouseEventArgs eventArgs)
  {
    base.OnMouseDown(eventArgs);
    Focus();
    if (eventArgs.Button == MouseButtons.Left && GetDeleteBounds().Contains(eventArgs.Location) && selectedTextId is Guid deleteId)
    {
      document.DeleteText(deleteId);
      selectedTextId = null;
      Invalidate();
      return;
    }
    if (eventArgs.Button == MouseButtons.Left && GetResizeBounds().Contains(eventArgs.Location))
    {
      BeginTextResize();
      return;
    }
    if (eventArgs.Button != MouseButtons.Left || !GetImageBounds().Contains(eventArgs.Location))
    {
      return;
    }

    if (document.TryGetTextAt(ToImagePoint(eventArgs.Location), out var textId))
    {
      selectedTextId = textId;
      if (eventArgs.Clicks == 2)
      {
        CancelDrag();
        Capture = false;
        TextEditRequested?.Invoke(this, textId);
      }
      else BeginTextMove(eventArgs.Location);
      return;
    }
    selectedTextId = null;
    if (Tool == ImageEditorTool.Text)
    {
      // MouseDown runs while WinForms owns mouse capture. Release it before
      // entering a modal message loop so clicks reach the text input window.
      Capture = false;
      TextRequested?.Invoke(this, new ImageTextRequestedEventArgs(ToImagePoint(eventArgs.Location)));
      return;
    }

    dragging = true;
    dragStartClient = eventArgs.Location;
    dragCurrentClient = eventArgs.Location;
    Capture = true;
    Invalidate();
  }

  protected override void OnMouseMove(MouseEventArgs eventArgs)
  {
    base.OnMouseMove(eventArgs);
    if (!dragging)
    {
      Cursor = GetResizeBounds().Contains(eventArgs.Location)
        ? Cursors.SizeNWSE
        : GetImageBounds().Contains(eventArgs.Location) &&
          document.TryGetTextAt(ToImagePoint(eventArgs.Location), out _)
          ? Cursors.SizeAll
          : Cursors.Cross;
      return;
    }

    if (resizingTextId is not null)
    {
      var point = ToImagePoint(ClampToImageBounds(eventArgs.Location));
      var scaleX = (point.X - textOriginalBounds.Left) / Math.Max(1F, textOriginalBounds.Width);
      var scaleY = (point.Y - textOriginalBounds.Top) / Math.Max(1F, textOriginalBounds.Height);
      textPreviewFontSize = textOriginalFontSize * Math.Max(0.1F, Math.Max(scaleX, scaleY));
      Invalidate();
      return;
    }

    if (movingTextId is not null)
    {
      if (!textDragThresholdPassed && Math.Abs(eventArgs.X - dragStartClient.X) < SystemInformation.DragSize.Width / 2 &&
          Math.Abs(eventArgs.Y - dragStartClient.Y) < SystemInformation.DragSize.Height / 2) return;
      textDragThresholdPassed = true;
      var currentImagePoint = ToImagePoint(ClampToImageBounds(eventArgs.Location));
      textPreviewLocation = new Point(
        textOriginalLocation.X + currentImagePoint.X - textDragStartImage.X,
        textOriginalLocation.Y + currentImagePoint.Y - textDragStartImage.Y);
      Invalidate();
      return;
    }

    dragCurrentClient = ClampToImageBounds(eventArgs.Location);
    Invalidate();
  }

  protected override void OnMouseUp(MouseEventArgs eventArgs)
  {
    base.OnMouseUp(eventArgs);
    if (!dragging || eventArgs.Button != MouseButtons.Left)
    {
      return;
    }

    if (movingTextId is not null) OnMouseMove(eventArgs);
    if (resizingTextId is not null) OnMouseMove(eventArgs);
    dragCurrentClient = ClampToImageBounds(eventArgs.Location);
    dragging = false;
    Capture = false;
    if (movingTextId is Guid movingId)
    {
      movingTextId = null;
      document.MoveText(movingId, textPreviewLocation);
      Invalidate();
      return;
    }
    if (resizingTextId is Guid resizingId)
    {
      resizingTextId = null;
      document.ResizeText(resizingId, textPreviewFontSize);
      Invalidate();
      return;
    }

    var start = ToImagePoint(dragStartClient);
    var end = ToImagePoint(dragCurrentClient);
    var changed = Tool switch
    {
      ImageEditorTool.Rectangle => document.DrawRectangle(Normalize(start, end), DrawingColor),
      ImageEditorTool.Arrow => document.DrawArrow(start, end, DrawingColor),
      ImageEditorTool.Mosaic => document.Mosaic(Normalize(start, end)),
      ImageEditorTool.Crop => document.Crop(Normalize(start, end)),
      _ => false,
    };

    if (!changed)
    {
      ActionRejected?.Invoke(this, "編集範囲を2ピクセル以上で指定してください。");
    }

    Invalidate();
  }

  protected override void OnMouseCaptureChanged(EventArgs eventArgs)
  {
    base.OnMouseCaptureChanged(eventArgs);
    if (Capture || !dragging)
    {
      return;
    }

    dragging = false;
    movingTextId = null;
    resizingTextId = null;
    Invalidate();
  }

  private void BeginTextMove(Point clientLocation)
  {
    var imageLocation = ToImagePoint(clientLocation);
    if (!document.TryGetTextAt(imageLocation, out var annotationId) ||
      !document.TryGetTextAnnotation(annotationId, out var annotation))
    {
      selectedTextId = null;
      ActionRejected?.Invoke(this, "移動するテキストをクリックしてください。");
      Invalidate();
      return;
    }

    selectedTextId = annotationId;
    movingTextId = annotationId;
    textDragThresholdPassed = false;
    dragStartClient = clientLocation;
    textDragStartImage = imageLocation;
    textOriginalLocation = annotation.Location;
    textPreviewLocation = annotation.Location;
    dragging = true;
    Capture = true;
    Invalidate();
  }

  private void BeginTextResize()
  {
    if (selectedTextId is not Guid id || !document.TryGetTextAnnotation(id, out var annotation)) return;
    resizingTextId = id;
    movingTextId = null;
    textOriginalBounds = annotation.Bounds;
    textOriginalFontSize = annotation.FontSize;
    textPreviewFontSize = annotation.FontSize;
    dragging = true;
    Capture = true;
    Invalidate();
  }

  internal bool CancelDrag()
  {
    if (!dragging) return false;
    dragging = false;
    movingTextId = null;
    resizingTextId = null;
    Capture = false;
    Invalidate();
    return true;
  }

  internal bool DeleteSelectedText()
  {
    if (selectedTextId is not Guid id) return false;
    CancelDrag();
    selectedTextId = null;
    var deleted = document.DeleteText(id);
    Invalidate();
    return deleted;
  }

  internal bool EditSelectedText()
  {
    if (selectedTextId is not Guid id || !document.TryGetTextAnnotation(id, out _)) return false;
    CancelDrag();
    TextEditRequested?.Invoke(this, id);
    return true;
  }

  private Rectangle GetDeleteBounds()
  {
    if (!TryGetSelectedTextPreview(out var annotation)) return Rectangle.Empty;
    var image = GetImageBounds();
    var size = Math.Max(20, (int)(22 * DeviceDpi / 96F));
    var x = image.X + (int)(annotation.Bounds.Right * image.Width / document.Width);
    var y = image.Y + (int)(annotation.Bounds.Top * image.Height / document.Height) - size;
    return new Rectangle(Math.Clamp(x, 0, Math.Max(0, Width - size)), Math.Clamp(y, 0, Math.Max(0, Height - size)), size, size);
  }

  private Rectangle GetResizeBounds()
  {
    if (!TryGetSelectedTextPreview(out var annotation)) return Rectangle.Empty;
    var image = GetImageBounds();
    var size = Math.Max(10, (int)(12 * DeviceDpi / 96F));
    var x = image.X + (int)(annotation.Bounds.Right * image.Width / document.Width) - size / 2;
    var y = image.Y + (int)(annotation.Bounds.Bottom * image.Height / document.Height) - size / 2;
    return new Rectangle(Math.Clamp(x, 0, Math.Max(0, Width - size)), Math.Clamp(y, 0, Math.Max(0, Height - size)), size, size);
  }

  private bool TryGetSelectedTextPreview(out TextAnnotation annotation)
  {
    annotation = default!;
    if (selectedTextId is not Guid id || !document.TryGetTextAnnotation(id, out annotation)) return false;
    if (movingTextId == id && document.TryGetMovedTextAnnotation(id, textPreviewLocation, out var moved)) annotation = moved;
    else if (resizingTextId == id && document.TryGetResizedTextAnnotation(id, textPreviewFontSize, out var resized)) annotation = resized;
    return true;
  }

  private void DrawSelectedTextBounds(Graphics graphics, TextAnnotation? preview)
  {
    if (selectedTextId is not Guid selectedId)
    {
      return;
    }

    TextAnnotation? selected = preview?.Id == selectedId ? preview : null;
    if (selected is null && !document.TryGetTextAnnotation(selectedId, out selected))
    {
      selectedTextId = null;
      return;
    }

    using var pen = new Pen(Color.DeepSkyBlue, Math.Max(1F, Math.Min(document.Width, document.Height) / 500F))
    {
      DashStyle = DashStyle.Dash,
    };
    graphics.DrawRectangle(pen, selected.Bounds.X, selected.Bounds.Y, selected.Bounds.Width, selected.Bounds.Height);
  }

  private Rectangle GetImageBounds()
  {
    var availableWidth = Math.Max(0, ClientSize.Width - (CanvasPadding * 2));
    var availableHeight = Math.Max(0, ClientSize.Height - (CanvasPadding * 2));
    if (availableWidth == 0 || availableHeight == 0)
    {
      return Rectangle.Empty;
    }

    var scale = Math.Min(
      availableWidth / (double)document.Width,
      availableHeight / (double)document.Height);
    var width = Math.Max(1, (int)Math.Round(document.Width * scale));
    var height = Math.Max(1, (int)Math.Round(document.Height * scale));
    return new Rectangle(
      (ClientSize.Width - width) / 2,
      (ClientSize.Height - height) / 2,
      width,
      height);
  }

  private Point ClampToImageBounds(Point point)
  {
    var bounds = GetImageBounds();
    return new Point(
      Math.Clamp(point.X, bounds.Left, bounds.Right - 1),
      Math.Clamp(point.Y, bounds.Top, bounds.Bottom - 1));
  }

  private Point ToImagePoint(Point clientPoint)
  {
    var bounds = GetImageBounds();
    var x = (clientPoint.X - bounds.Left) * document.Width / (double)bounds.Width;
    var y = (clientPoint.Y - bounds.Top) * document.Height / (double)bounds.Height;
    return new Point(
      Math.Clamp((int)Math.Floor(x), 0, document.Width - 1),
      Math.Clamp((int)Math.Floor(y), 0, document.Height - 1));
  }

  private static Rectangle Normalize(Point first, Point second)
  {
    var left = Math.Min(first.X, second.X);
    var top = Math.Min(first.Y, second.Y);
    var right = Math.Max(first.X, second.X);
    var bottom = Math.Max(first.Y, second.Y);
    return Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
  }
}

internal sealed class ImageTextRequestedEventArgs(Point imageLocation) : EventArgs
{
  public Point ImageLocation { get; } = imageLocation;
}

internal sealed class ImageTextInputDialog : Form
{
  private readonly TextBox textBox = new()
  {
    Dock = DockStyle.Fill,
    Multiline = true,
    AcceptsReturn = true,
    ScrollBars = ScrollBars.Vertical,
    MaxLength = 500,
  };

  public ImageTextInputDialog(string? existingText = null)
  {
    Text = existingText is null ? "テキストを追加" : "テキストを編集";
    textBox.Text = existingText ?? string.Empty;
    textBox.SelectionStart = textBox.TextLength;
    StartPosition = FormStartPosition.CenterParent;
    ClientSize = new Size(440, 180);
    MinimumSize = new Size(360, 160);
    AutoScaleMode = AutoScaleMode.Dpi;
    Font = new Font("Meiryo UI", 9F);
    UiTheme.StyleForm(this);
    ShowInTaskbar = false;

    var label = new Label
    {
      Text = "画像へ追加する文字:",
      Dock = DockStyle.Top,
      AutoSize = true,
      Padding = new Padding(0, 0, 0, 6),
      ForeColor = UiTheme.Text,
    };
    var okButton = new Button
    {
      Text = existingText is null ? "追加" : "変更",
      AutoSize = true,
      DialogResult = DialogResult.OK,
    };
    var cancelButton = new Button
    {
      Text = "キャンセル",
      AutoSize = true,
      DialogResult = DialogResult.Cancel,
    };
    var buttons = new FlowLayoutPanel
    {
      Dock = DockStyle.Bottom,
      AutoSize = true,
      FlowDirection = FlowDirection.RightToLeft,
      Padding = new Padding(0, 8, 0, 0),
      BackColor = UiTheme.Canvas,
    };
    UiTheme.StyleButton(okButton, Font, primary: true);
    okButton.MinimumSize = new Size(92, 32);
    UiTheme.StyleButton(cancelButton, Font);
    cancelButton.MinimumSize = new Size(108, 32);
    buttons.Controls.Add(cancelButton);
    buttons.Controls.Add(okButton);

    var layout = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12), BackColor = UiTheme.Canvas };
    layout.Controls.Add(textBox);
    layout.Controls.Add(label);
    layout.Controls.Add(buttons);
    Controls.Add(layout);
    AcceptButton = okButton;
    FormClosing += (_, args) =>
    {
      if (DialogResult == DialogResult.OK && string.IsNullOrWhiteSpace(textBox.Text))
      { args.Cancel = true; textBox.Focus(); }
    };
    CancelButton = cancelButton;
  }

  public string EnteredText => textBox.Text;

  protected override bool ProcessCmdKey(ref Message message, Keys keyData)
  {
    if (EnterShortcut.IsEditingInput(this)) return base.ProcessCmdKey(ref message, keyData);
    if (keyData == Keys.Enter)
    {
      if (!EnterShortcut.IsRepeat(message)) DialogResult = DialogResult.OK;
      return true;
    }
    if (keyData == (Keys.Shift | Keys.Enter))
    {
      textBox.SelectedText = Environment.NewLine;
      return true;
    }
    return base.ProcessCmdKey(ref message, keyData);
  }

  protected override void OnShown(EventArgs eventArgs)
  {
    base.OnShown(eventArgs);
    textBox.Focus();
  }
}
