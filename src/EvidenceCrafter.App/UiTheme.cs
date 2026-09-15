using System.Drawing.Drawing2D;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace EvidenceCrafter.App;

internal static class UiTheme
{
  private enum Role { Form, Canvas, Surface, MutedSurface, Text, MutedText, Border, Button, PrimaryButton, ThemeColorButton, TextBox, ComboBox, SideButton }
  private sealed class RoleHolder(Role value) { internal Role Value { get; } = value; }
  private static readonly ConditionalWeakTable<Control, RoleHolder> roles = new();

  private static int themeIntensity;
  private static Color themeColor = Color.FromArgb(EvidenceCrafterSettings.DefaultThemeColorArgb);

  internal static int ThemeIntensity => themeIntensity;
  internal static bool DarkMode => themeIntensity >= 50;
  internal static Color ThemeColor => themeColor;
  internal static Color Canvas { get; private set; }
  internal static Color Surface { get; private set; }
  internal static Color SurfaceMuted { get; private set; }
  internal static Color Text { get; private set; }
  internal static Color TextMuted { get; private set; }
  internal static Color Border { get; private set; }
  internal static Color Primary { get; private set; }
  internal static Color PrimaryHover { get; private set; }

  static UiTheme() => RebuildPalette();

  internal static void SetThemeIntensity(int value)
  {
    value = Math.Clamp(value, 0, 100);
    if (themeIntensity == value) return;
    themeIntensity = value;
    RebuildPalette();
  }

  internal static void SetThemeColor(Color color)
  {
    color = Color.FromArgb(color.R, color.G, color.B);
    if (themeColor == color) return;
    themeColor = color;
    RebuildPalette();
  }

  private static void RebuildPalette()
  {
    Canvas = Blend(Color.White, themeColor);
    Surface = Blend(Color.White, Mix(themeColor, Color.White, 0.14));
    SurfaceMuted = Blend(Color.FromArgb(244, 246, 248), Mix(themeColor, Color.White, 0.24));
    Text = TextOn(Canvas);
    TextMuted = MutedTextOn(SurfaceMuted);
    Border = Mix(Surface, TextOn(Surface), 0.20);
    Primary = Blend(Color.FromArgb(23, 105, 170), Accent(themeColor));
    PrimaryHover = Mix(Primary, TextOn(Primary), 0.16);
  }

  internal static void SetDarkMode(bool enabled) => SetThemeIntensity(enabled ? 100 : 0);

  internal static void StyleForm(Form form)
  {
    SetRole(form, Role.Form);
    Apply(form, Role.Form);
    form.Font = new Font("Meiryo UI", 9F);
  }

  internal static void StyleSurface(Control control, bool muted = false)
  {
    var role = muted ? Role.MutedSurface : Role.Surface;
    SetRole(control, role);
    Apply(control, role);
  }

  internal static void StyleCanvas(Control control)
  {
    SetRole(control, Role.Canvas);
    Apply(control, Role.Canvas);
  }

  internal static void StyleText(Control control, bool muted = false)
  {
    var role = muted ? Role.MutedText : Role.Text;
    SetRole(control, role);
    Apply(control, role);
  }

  internal static void StyleBorder(Control control)
  {
    SetRole(control, Role.Border);
    Apply(control, Role.Border);
  }

  internal static void StyleButton(Button button, Font font, bool primary = false)
  {
    var role = primary ? Role.PrimaryButton : Role.Button;
    SetRole(button, role);
    button.AutoSize = true;
    button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    button.Font = font;
    button.MinimumSize = new Size(64, 32);
    button.FlatStyle = FlatStyle.Flat;
    button.FlatAppearance.BorderSize = 1;
    button.UseVisualStyleBackColor = false;
    button.Margin = new Padding(3, 0, 3, 0);
    button.Padding = new Padding(12, 2, 12, 2);
    button.TextAlign = ContentAlignment.MiddleCenter;
    button.UseCompatibleTextRendering = false;
    Apply(button, role);
  }

  internal static void StyleTextBox(TextBox textBox)
  {
    SetRole(textBox, Role.TextBox);
    textBox.AutoSize = true;
    textBox.BorderStyle = BorderStyle.FixedSingle;
    Apply(textBox, Role.TextBox);
  }

  internal static void StyleComboBox(ComboBox comboBox)
  {
    SetRole(comboBox, Role.ComboBox);
    Apply(comboBox, Role.ComboBox);
  }

  internal static void StyleThemeColorButton(ThemeColorPickerButton button, Font font)
  {
    SetRole(button, Role.ThemeColorButton);
    button.AutoSize = false;
    button.Font = font;
    button.MinimumSize = new Size(28, 28);
    button.Size = new Size(28, 28);
    button.FlatStyle = FlatStyle.Flat;
    button.UseVisualStyleBackColor = false;
    button.Margin = new Padding(5, 0, 0, 0);
    button.Padding = Padding.Empty;
    button.TextAlign = ContentAlignment.MiddleCenter;
    button.UseCompatibleTextRendering = false;
    Apply(button, Role.ThemeColorButton);
  }

  internal static void StyleSideButton(RadioButton button)
  {
    SetRole(button, Role.SideButton);
    button.Appearance = Appearance.Button;
    button.AutoSize = true;
    button.MinimumSize = new Size(64, 32);
    button.Padding = new Padding(12, 2, 12, 2);
    button.TextAlign = ContentAlignment.MiddleCenter;
    button.UseCompatibleTextRendering = false;
    button.FlatStyle = FlatStyle.Flat;
    button.Margin = new Padding(0, 0, 4, 0);
    Apply(button, Role.SideButton);
  }

  internal static void Refresh(Control root)
  {
    if (roles.TryGetValue(root, out var holder)) Apply(root, holder.Value);
    foreach (Control child in root.Controls) Refresh(child);
  }

  private static void SetRole(Control control, Role role)
  {
    roles.Remove(control);
    roles.Add(control, new RoleHolder(role));
  }

  private static void Apply(Control control, Role role)
  {
    switch (role)
    {
      case Role.Form: control.BackColor = Canvas; control.ForeColor = Text; break;
      case Role.Canvas: control.BackColor = Canvas; control.ForeColor = Text; break;
      case Role.Surface: control.BackColor = Surface; control.ForeColor = TextOn(Surface); break;
      case Role.MutedSurface: control.BackColor = SurfaceMuted; control.ForeColor = TextOn(SurfaceMuted); break;
      case Role.Text: control.ForeColor = TextColorFor(control); break;
      case Role.MutedText: control.ForeColor = MutedTextColorFor(control); break;
      case Role.Border: control.BackColor = Border; break;
      case Role.Button: ApplyButton((Button)control, false); break;
      case Role.PrimaryButton: ApplyButton((Button)control, true); break;
      case Role.ThemeColorButton: ApplyThemeColorButton((ThemeColorPickerButton)control); break;
      case Role.TextBox: control.BackColor = Surface; control.ForeColor = TextOn(Surface); break;
      case Role.ComboBox: control.BackColor = Surface; control.ForeColor = TextOn(Surface); break;
      case Role.SideButton:
        var button = (RadioButton)control;
        button.FlatAppearance.CheckedBackColor = Primary;
        button.FlatAppearance.MouseOverBackColor = SurfaceMuted;
        break;
    }
  }

  private static void ApplyButton(Button button, bool primary)
  {
    button.BackColor = primary ? Primary : SurfaceMuted;
    button.ForeColor = TextOn(button.BackColor);
    button.FlatAppearance.BorderColor = primary ? PrimaryHover : Border;
    button.FlatAppearance.MouseOverBackColor = primary ? PrimaryHover : Surface;
    button.FlatAppearance.MouseDownBackColor = primary ? Primary : SurfaceMuted;
  }

  private static Color TextColorFor(Control control) =>
    TextOn(control.Parent?.BackColor ?? Canvas);

  private static Color MutedTextColorFor(Control control) =>
    MutedTextOn(control.Parent?.BackColor ?? SurfaceMuted);

  private static void ApplyThemeColorButton(ThemeColorPickerButton button)
  {
    button.BackColor = ThemeColor;
    button.Invalidate();
  }

  private static Color Blend(Color light, Color dark)
  {
    var amount = themeIntensity / 100d;
    return Color.FromArgb(
      Interpolate(light.R, dark.R, amount),
      Interpolate(light.G, dark.G, amount),
      Interpolate(light.B, dark.B, amount));
  }

  internal static Color TextOn(Color background) =>
    Luminance(background) < 0.179 ? Color.White : Color.Black;

  private static Color MutedTextOn(Color background)
  {
    var text = TextOn(background);
    var muted = Mix(text, background, 0.28);
    var first = Luminance(muted);
    var second = Luminance(background);
    return (Math.Max(first, second) + 0.05) / (Math.Min(first, second) + 0.05) >= 4.5 ? muted : text;
  }

  private static Color Accent(Color color) =>
    Luminance(color) < 0.48
      ? Mix(color, Color.White, 0.18)
      : Mix(color, Color.FromArgb(23, 32, 42), 0.22);

  private static Color Mix(Color first, Color second, double amount) =>
    Color.FromArgb(
      Interpolate(first.R, second.R, amount),
      Interpolate(first.G, second.G, amount),
      Interpolate(first.B, second.B, amount));

  private static int Interpolate(int light, int dark, double amount) =>
    (int)Math.Round(light + (dark - light) * amount, MidpointRounding.AwayFromZero);

  private static double Luminance(Color color)
  {
    static double Channel(byte value)
    {
      var normalized = value / 255d;
      return normalized <= 0.03928
        ? normalized / 12.92
        : Math.Pow((normalized + 0.055) / 1.055, 2.4);
    }

    return 0.2126 * Channel(color.R) + 0.7152 * Channel(color.G) + 0.0722 * Channel(color.B);
  }
}

internal sealed class ThemedCheckBox : CheckBox
{
  protected override bool ShowFocusCues => false;
}

/// <summary>Keeps compact header text on one line even when TableLayout constrains the cell.</summary>
internal sealed class SingleLineLabel : Label
{
  public override Size GetPreferredSize(Size proposedSize)
  {
    var measured = TextRenderer.MeasureText(
      Text,
      Font,
      Size.Empty,
      TextFormatFlags.SingleLine | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
    return new Size(Math.Max(MinimumSize.Width, measured.Width), Math.Max(MinimumSize.Height, measured.Height));
  }

  protected override void OnPaint(PaintEventArgs eventArgs)
  {
    eventArgs.Graphics.Clear(BackColor);
    TextRenderer.DrawText(
      eventArgs.Graphics,
      Text,
      Font,
      ClientRectangle,
      ForeColor,
      TextFormatFlags.SingleLine | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix |
      TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
  }
}

internal sealed class ThemeColorPickerButton : Button
{
  private bool hovered;

  internal ThemeColorPickerButton()
  {
    SetStyle(
      ControlStyles.UserPaint |
      ControlStyles.AllPaintingInWmPaint |
      ControlStyles.OptimizedDoubleBuffer |
      ControlStyles.ResizeRedraw,
      true);
    Cursor = Cursors.Hand;
  }

  protected override bool ShowFocusCues => false;

  protected override void OnMouseEnter(EventArgs e)
  {
    base.OnMouseEnter(e);
    hovered = true;
    Invalidate();
  }

  protected override void OnMouseLeave(EventArgs e)
  {
    base.OnMouseLeave(e);
    hovered = false;
    Invalidate();
  }

  protected override void OnPaint(PaintEventArgs e)
  {
    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
    e.Graphics.Clear(Parent?.BackColor ?? UiTheme.Canvas);

    var bounds = new Rectangle(1, 1, Math.Max(1, Width - 3), Math.Max(1, Height - 3));
    using var outline = new Pen(UiTheme.Border, hovered ? 2 : 1);
    using var fill = new SolidBrush(BackColor);
    e.Graphics.FillEllipse(fill, bounds);
    e.Graphics.DrawEllipse(outline, bounds);
  }
}

internal sealed class ThemeGradientSlider : Control
{
  private const int TrackHeight = 12;
  private const int TrackHorizontalInset = 4;
  private const int ThumbHorizontalInset = 8;
  private int value;
  private int minimumValue;
  private int maximumValue = 100;
  private bool dragging;

  internal ThemeGradientSlider()
  {
    SetStyle(
      ControlStyles.UserPaint |
      ControlStyles.AllPaintingInWmPaint |
      ControlStyles.OptimizedDoubleBuffer |
      ControlStyles.ResizeRedraw |
      ControlStyles.Selectable,
      true);
    AccessibleRole = AccessibleRole.Slider;
    AccessibleName = "テーマの明るさ";
    TabStop = true;
    Cursor = Cursors.Hand;
    MinimumSize = new Size(110, 30);
    Size = new Size(160, 30);
  }

  [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
  internal int Value
  {
    get => value;
    set => SetValue(value, raiseEvent: true);
  }

  [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
  internal int MinimumValue
  {
    get => minimumValue;
    set
    {
      minimumValue = Math.Clamp(value, 0, maximumValue - 1);
      SetValue(this.value, raiseEvent: false);
      Invalidate();
    }
  }

  [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
  internal int MaximumValue
  {
    get => maximumValue;
    set
    {
      maximumValue = Math.Clamp(value, minimumValue + 1, 100);
      SetValue(this.value, raiseEvent: false);
      Invalidate();
    }
  }

  internal event EventHandler? ValueChanged;

  protected override void OnPaint(PaintEventArgs e)
  {
    base.OnPaint(e);
    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
    e.Graphics.Clear(BackColor);

    var track = GetTrackBounds();
    using (var path = RoundedRectangle(track, TrackHeight / 2))
    using (var brush = new LinearGradientBrush(
      track,
      Color.White,
      UiTheme.ThemeColor,
      LinearGradientMode.Horizontal))
    {
      e.Graphics.FillPath(brush, path);
      using var border = new Pen(Color.FromArgb(110, 128, 145));
      e.Graphics.DrawPath(border, path);
    }

    var x = ValueToX(value);
    var thumb = new Rectangle(x - 8, track.Top - 4, 16, 16);
    using (var shadow = new SolidBrush(Color.FromArgb(55, 0, 0, 0)))
    {
      e.Graphics.FillEllipse(shadow, new Rectangle(thumb.X + 1, thumb.Y + 1, thumb.Width, thumb.Height));
    }
    using (var fill = new SolidBrush(UiTheme.Primary))
    using (var outline = new Pen(UiTheme.Surface, 2))
    {
      e.Graphics.FillEllipse(fill, thumb);
      e.Graphics.DrawEllipse(outline, thumb);
    }

  }

  protected override void OnMouseDown(MouseEventArgs e)
  {
    base.OnMouseDown(e);
    if (e.Button != MouseButtons.Left) return;
    Focus();
    dragging = true;
    Capture = true;
    SetValue(XToValue(e.X), raiseEvent: true);
  }

  protected override void OnMouseMove(MouseEventArgs e)
  {
    base.OnMouseMove(e);
    if (dragging) SetValue(XToValue(e.X), raiseEvent: true);
  }

  protected override void OnMouseUp(MouseEventArgs e)
  {
    base.OnMouseUp(e);
    if (e.Button == MouseButtons.Left)
    {
      dragging = false;
      Capture = false;
    }
  }

  protected override void OnMouseCaptureChanged(EventArgs e)
  {
    base.OnMouseCaptureChanged(e);
    if (!Capture) dragging = false;
  }

  protected override void OnKeyDown(KeyEventArgs e)
  {
    var next = e.KeyCode switch
    {
      Keys.Left or Keys.Down => value - 5,
      Keys.Right or Keys.Up => value + 5,
      Keys.Home => minimumValue,
      Keys.End => maximumValue,
      _ => value,
    };
    if (next != value)
    {
      SetValue(next, raiseEvent: true);
      e.Handled = true;
      e.SuppressKeyPress = true;
      return;
    }
    base.OnKeyDown(e);
  }

  protected override bool IsInputKey(Keys keyData) => keyData switch
  {
    Keys.Left or Keys.Right or Keys.Up or Keys.Down or Keys.Home or Keys.End => true,
    _ => base.IsInputKey(keyData),
  };

  protected override void OnBackColorChanged(EventArgs e)
  {
    base.OnBackColorChanged(e);
    Invalidate();
  }

  private void SetValue(int next, bool raiseEvent)
  {
    next = Math.Clamp(next, minimumValue, maximumValue);
    if (value == next) return;
    value = next;
    AccessibleDescription = $"{value}%";
    Invalidate();
    if (raiseEvent) ValueChanged?.Invoke(this, EventArgs.Empty);
  }

  private Rectangle GetTrackBounds()
  {
    var width = Math.Max(1, ClientSize.Width - TrackHorizontalInset * 2);
    return new Rectangle(
      TrackHorizontalInset,
      Math.Max(0, (ClientSize.Height - TrackHeight) / 2),
      width,
      TrackHeight);
  }

  private Rectangle GetThumbTrackBounds()
  {
    var width = Math.Max(1, ClientSize.Width - ThumbHorizontalInset * 2);
    return new Rectangle(
      ThumbHorizontalInset,
      Math.Max(0, (ClientSize.Height - TrackHeight) / 2),
      width,
      TrackHeight);
  }

  private int ValueToX(int currentValue)
  {
    var track = GetThumbTrackBounds();
    return track.Left + (int)Math.Round(
      track.Width * (currentValue - minimumValue) / (double)(maximumValue - minimumValue),
      MidpointRounding.AwayFromZero);
  }

  private int XToValue(int x)
  {
    var track = GetThumbTrackBounds();
    return track.Width <= 0
      ? minimumValue
      : minimumValue + (int)Math.Round(
        Math.Clamp(x - track.Left, 0, track.Width) * (maximumValue - minimumValue) / (double)track.Width,
        MidpointRounding.AwayFromZero);
  }

  private static GraphicsPath RoundedRectangle(Rectangle rectangle, int radius)
  {
    var diameter = radius * 2;
    var path = new GraphicsPath();
    path.AddArc(rectangle.Left, rectangle.Top, diameter, diameter, 180, 90);
    path.AddArc(rectangle.Right - diameter, rectangle.Top, diameter, diameter, 270, 90);
    path.AddArc(rectangle.Right - diameter, rectangle.Bottom - diameter, diameter, diameter, 0, 90);
    path.AddArc(rectangle.Left, rectangle.Bottom - diameter, diameter, diameter, 90, 90);
    path.CloseFigure();
    return path;
  }
}
