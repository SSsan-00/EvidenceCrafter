using System.Reflection;
using System.Runtime.ExceptionServices;
using EvidenceCrafter.App;

namespace EvidenceCrafter.Tests;

[TestClass]
[DoNotParallelize]
public sealed class ThemeReviewTests
{
  [TestMethod]
  public void Theme_CustomColorsKeepTextReadableOnEverySurface() => OnSta(() =>
  {
    var originalColor = UiTheme.ThemeColor;
    var originalIntensity = UiTheme.ThemeIntensity;
    try
    {
      using var form = new Form();
      using var card = new Panel();
      using var label = new Label();
      using var combo = new ComboBox();
      using var button = new Button();
      UiTheme.StyleForm(form);
      UiTheme.StyleSurface(card, muted: true);
      UiTheme.StyleText(label, muted: true);
      UiTheme.StyleComboBox(combo);
      UiTheme.StyleButton(button, form.Font, primary: true);
      form.Controls.Add(card);
      card.Controls.AddRange([label, combo, button]);
      foreach (var color in new[] { Color.Black, Color.White, Color.Gray, Color.Red, Color.Lime, Color.Blue, Color.Gold, Color.Teal })
      {
        UiTheme.SetThemeColor(color);
        for (var intensity = 0; intensity <= 100; intensity += 5)
        {
          UiTheme.SetThemeIntensity(intensity);
          UiTheme.Refresh(form);
          foreach (var control in new Control[] { form, card, label, combo, button })
          {
            var ratio = Contrast(control.ForeColor, control.BackColor);
            Assert.IsGreaterThanOrEqualTo(4.5, ratio, $"{color}/{intensity}/{control.GetType().Name}: {ratio:F2}");
          }
        }
        Assert.AreEqual(color.ToArgb(), UiTheme.Canvas.ToArgb(), "The selected endpoint must update cached colors.");
        UiTheme.SetThemeIntensity(0);
        Assert.AreEqual(Color.White.ToArgb(), UiTheme.Canvas.ToArgb());
      }
    }
    finally
    {
      UiTheme.SetThemeColor(originalColor);
      UiTheme.SetThemeIntensity(originalIntensity);
    }
  });

  [TestMethod]
  public void Slider_CaptureLossStopsDraggingAndEndpointsStayInRange() => OnSta(() =>
  {
    using var slider = new ThemeGradientSlider { MinimumValue = 50, Value = 70 };
    void Mouse(string name, int x) => typeof(ThemeGradientSlider)
      .GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic)!
      .Invoke(slider, [new MouseEventArgs(MouseButtons.Left, 1, x, 15, 0)]);
    Mouse("OnMouseDown", 0);
    Assert.AreEqual(50, slider.Value);
    Assert.IsTrue(slider.Capture);
    Mouse("OnMouseMove", 1000);
    Assert.AreEqual(100, slider.Value);
    slider.Capture = false;
    Mouse("OnMouseMove", 0);
    Assert.AreEqual(100, slider.Value, "Losing capture must stop a drag even without MouseUp.");
  });

  [TestMethod]
  public void SingleLineLabel_ReportsTheFullUnwrappedTitleWidth() => OnSta(() =>
  {
    using var label = new SingleLineLabel
    {
      Text = "EvidenceCrafter",
      Font = new Font("Meiryo UI", 15F, FontStyle.Bold),
      MinimumSize = new Size(1, 1),
    };

    var expected = TextRenderer.MeasureText(
      label.Text,
      label.Font,
      Size.Empty,
      TextFormatFlags.SingleLine | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
    Assert.IsGreaterThanOrEqualTo(expected.Width, label.GetPreferredSize(Size.Empty).Width);
  });

  [TestMethod]
  public void RainbowBackdrop_PaintsAStaticGradientAndOnlyAnimatesWhenRequested() => OnSta(() =>
  {
    using var backdrop = new RainbowBackdrop { Size = new Size(240, 80), BaseColor = Color.White };
    using var staticImage = new Bitmap(backdrop.Width, backdrop.Height);
    backdrop.Mode = RainbowBackgroundMode.Static;
    backdrop.DrawToBitmap(staticImage, backdrop.ClientRectangle);
    Assert.AreNotEqual(staticImage.GetPixel(5, 5).ToArgb(), staticImage.GetPixel(220, 70).ToArgb());

    using var firstAnimatedImage = new Bitmap(backdrop.Width, backdrop.Height);
    backdrop.Mode = RainbowBackgroundMode.Animated;
    backdrop.DrawToBitmap(firstAnimatedImage, backdrop.ClientRectangle);
    backdrop.AdvanceAnimation();
    using var nextAnimatedImage = new Bitmap(backdrop.Width, backdrop.Height);
    backdrop.DrawToBitmap(nextAnimatedImage, backdrop.ClientRectangle);
    Assert.AreNotEqual(firstAnimatedImage.GetPixel(120, 40).ToArgb(), nextAnimatedImage.GetPixel(120, 40).ToArgb());
  });

  [TestMethod]
  public void ThemeAnimatedBackdrop_UsesTheSelectedColorAndAnimates() => OnSta(() =>
  {
    using var backdrop = new RainbowBackdrop { Size = new Size(240, 80), BaseColor = Color.White, Mode = RainbowBackgroundMode.ThemeAnimated };
    using var redImage = new Bitmap(backdrop.Width, backdrop.Height);
    backdrop.ThemeColor = Color.Red;
    backdrop.DrawToBitmap(redImage, backdrop.ClientRectangle);

    using var blueImage = new Bitmap(backdrop.Width, backdrop.Height);
    backdrop.ThemeColor = Color.Blue;
    backdrop.DrawToBitmap(blueImage, backdrop.ClientRectangle);
    var redPixel = redImage.GetPixel(120, 40);
    var bluePixel = blueImage.GetPixel(120, 40);
    Assert.IsGreaterThan(redPixel.B, redPixel.R);
    Assert.IsGreaterThan(bluePixel.R, bluePixel.B);

    backdrop.AdvanceAnimation();
    using var animatedImage = new Bitmap(backdrop.Width, backdrop.Height);
    backdrop.DrawToBitmap(animatedImage, backdrop.ClientRectangle);
    Assert.AreNotEqual(blueImage.GetPixel(120, 40).ToArgb(), animatedImage.GetPixel(120, 40).ToArgb());
  });

  [TestMethod]
  public void RainbowBackground_ReachesTheSameCanvasAndSurfaceAreasAsTheTheme() => OnSta(() =>
  {
    using var form = new Form();
    using var card = new Panel();
    using var label = new Label();
    UiTheme.StyleForm(form);
    UiTheme.StyleSurface(card);
    UiTheme.StyleText(label);
    form.Controls.Add(card);
    card.Controls.Add(label);

    UiTheme.ApplyRainbowBackground(form, enabled: true);
    Assert.AreEqual(Color.Transparent.ToArgb(), card.BackColor.ToArgb());
    Assert.AreEqual(Color.Black.ToArgb(), label.ForeColor.ToArgb());

    UiTheme.ApplyRainbowBackground(form, enabled: false);
    Assert.AreEqual(UiTheme.Surface.ToArgb(), card.BackColor.ToArgb());
  });

  private static double Contrast(Color first, Color second)
  {
    static double Luminance(Color color)
    {
      static double Channel(byte b) => b / 255d <= 0.04045 ? b / 255d / 12.92 : Math.Pow((b / 255d + 0.055) / 1.055, 2.4);
      return .2126 * Channel(color.R) + .7152 * Channel(color.G) + .0722 * Channel(color.B);
    }
    var a = Luminance(first);
    var b = Luminance(second);
    return (Math.Max(a, b) + .05) / (Math.Min(a, b) + .05);
  }

  private static void OnSta(Action action)
  {
    Exception? failure = null;
    var thread = new Thread(() => { try { action(); } catch (Exception e) { failure = e; } }) { IsBackground = true };
    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(20)), "UI check timed out.");
    if (failure is not null) ExceptionDispatchInfo.Capture(failure).Throw();
  }
}
