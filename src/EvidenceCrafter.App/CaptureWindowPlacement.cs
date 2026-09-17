using System.Runtime.InteropServices;

namespace EvidenceCrafter.App;

/// <summary>Finds the window beneath a completed screen capture and centers dialogs over it.</summary>
internal static class CaptureWindowPlacement
{
  private const uint RootAncestor = 2;

  internal static Rectangle? TryGetWindowBounds(Point point)
  {
    var window = WindowFromPoint(new NativePoint(point.X, point.Y));
    if (window == 0) return null;

    var root = GetAncestor(window, RootAncestor);
    if (root != 0) window = root;
    if (!GetWindowRect(window, out var bounds) || bounds.Right <= bounds.Left || bounds.Bottom <= bounds.Top)
      return null;

    return Rectangle.FromLTRB(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom);
  }

  internal static void CenterInWindow(Form dialog, Rectangle windowBounds)
  {
    var workingArea = Screen.FromPoint(windowBounds.Location + new Size(windowBounds.Width / 2, windowBounds.Height / 2)).WorkingArea;
    dialog.Location = CenterLocation(windowBounds, dialog.Size, workingArea);
  }

  internal static Point CenterLocation(Rectangle windowBounds, Size dialogSize, Rectangle workingArea)
  {
    var desiredX = windowBounds.Left + (windowBounds.Width - dialogSize.Width) / 2;
    var desiredY = windowBounds.Top + (windowBounds.Height - dialogSize.Height) / 2;
    var maximumX = Math.Max(workingArea.Left, workingArea.Right - dialogSize.Width);
    var maximumY = Math.Max(workingArea.Top, workingArea.Bottom - dialogSize.Height);
    return new Point(
      Math.Clamp(desiredX, workingArea.Left, maximumX),
      Math.Clamp(desiredY, workingArea.Top, maximumY));
  }

  [DllImport("user32.dll")]
  private static extern nint WindowFromPoint(NativePoint point);

  [DllImport("user32.dll")]
  private static extern nint GetAncestor(nint window, uint flags);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  private static extern bool GetWindowRect(nint window, out NativeRectangle rectangle);

  [StructLayout(LayoutKind.Sequential)]
  private struct NativePoint
  {
    internal NativePoint(int x, int y)
    {
      X = x;
      Y = y;
    }

    internal int X;
    internal int Y;
  }

  [StructLayout(LayoutKind.Sequential)]
  private struct NativeRectangle
  {
    internal int Left;
    internal int Top;
    internal int Right;
    internal int Bottom;
  }
}
