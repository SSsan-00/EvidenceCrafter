using System.Runtime.InteropServices;

namespace EvidenceCrafter.App;

internal static class NativeClipboard
{
  internal const int ClipboardUpdateMessage = 0x031D;

  [DllImport("user32.dll")]
  internal static extern uint GetClipboardSequenceNumber();

  [DllImport("user32.dll", SetLastError = true)]
  [return: MarshalAs(UnmanagedType.Bool)]
  internal static extern bool AddClipboardFormatListener(nint windowHandle);

  [DllImport("user32.dll", SetLastError = true)]
  [return: MarshalAs(UnmanagedType.Bool)]
  internal static extern bool RemoveClipboardFormatListener(nint windowHandle);

  internal static bool IsExcelCellRangeCopy(IEnumerable<string>? formats)
  {
    if (formats is null) return false;
    foreach (var format in formats)
    {
      if (string.IsNullOrEmpty(format)) continue;
      if (format.StartsWith("Biff", StringComparison.OrdinalIgnoreCase) ||
        format.Equals("XML Spreadsheet", StringComparison.OrdinalIgnoreCase) ||
        format.Equals("Excel 12.0", StringComparison.OrdinalIgnoreCase))
      {
        return true;
      }
    }
    return false;
  }
}
