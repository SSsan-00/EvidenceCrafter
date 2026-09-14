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
