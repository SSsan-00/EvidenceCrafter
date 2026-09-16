using EvidenceCrafter.App;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class ClipboardFormatTests
{
  [TestMethod]
  public void ExcelCellFormats_AreRecognizedWithoutClassifyingImages()
  {
    Assert.IsTrue(NativeClipboard.IsExcelCellRangeCopy(["Biff12", DataFormats.Bitmap]));
    Assert.IsTrue(NativeClipboard.IsExcelCellRangeCopy(["XML Spreadsheet", DataFormats.EnhancedMetafile]));
    Assert.IsFalse(NativeClipboard.IsExcelCellRangeCopy([DataFormats.Bitmap, DataFormats.Tiff]));
  }
}
