using EvidenceCrafter.App;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class ImageEditDocumentTests
{
  [TestMethod]
  public void CropUndoRedo_RestoresEachRasterState()
  {
    using var source = CreateQuadrantImage();
    using var document = new ImageEditDocument(source);

    Assert.IsTrue(document.Crop(new Rectangle(4, 0, 4, 4)));
    Assert.AreEqual(4, document.Width);
    Assert.AreEqual(Color.Green.ToArgb(), document.GetImageCopy().GetPixel(1, 1).ToArgb());

    Assert.IsTrue(document.Undo());
    Assert.AreEqual(8, document.Width);
    Assert.AreEqual(Color.Red.ToArgb(), document.GetImageCopy().GetPixel(1, 1).ToArgb());

    Assert.IsTrue(document.Redo());
    Assert.AreEqual(4, document.Width);
    Assert.AreEqual(Color.Green.ToArgb(), document.GetImageCopy().GetPixel(1, 1).ToArgb());
  }

  [TestMethod]
  public void ResetThenUndo_RestoresEditedImage()
  {
    using var source = new Bitmap(20, 20);
    using (var graphics = Graphics.FromImage(source))
    {
      graphics.Clear(Color.White);
    }

    using var document = new ImageEditDocument(source);
    Assert.IsTrue(document.DrawRectangle(new Rectangle(2, 2, 12, 12)));
    using var edited = document.GetImageCopy();
    Assert.IsTrue(document.HasChanges);

    Assert.IsTrue(document.Reset());
    Assert.IsFalse(document.HasChanges);
    using var reset = document.GetImageCopy();
    Assert.AreEqual(Color.White.ToArgb(), reset.GetPixel(2, 2).ToArgb());

    Assert.IsTrue(document.Undo());
    Assert.IsTrue(document.HasChanges);
    using var restored = document.GetImageCopy();
    Assert.AreEqual(edited.GetPixel(2, 2).ToArgb(), restored.GetPixel(2, 2).ToArgb());
  }

  [TestMethod]
  public void NewEditAfterUndo_DiscardsRedoState()
  {
    using var source = new Bitmap(30, 30);
    using var document = new ImageEditDocument(source);

    Assert.IsTrue(document.DrawArrow(new Point(1, 1), new Point(20, 20)));
    Assert.IsTrue(document.Undo());
    Assert.IsTrue(document.CanRedo);

    Assert.IsTrue(document.DrawText("test", new Point(2, 2)));
    Assert.IsFalse(document.CanRedo);
  }

  [TestMethod]
  public void TextOnlyHistory_ReusesRasterUntilPixelsChange()
  {
    using var source = new Bitmap(800, 600);
    using var document = new ImageEditDocument(source);
    var raster = document.CurrentImage;

    Assert.IsTrue(document.DrawText("label", new Point(10, 10)));
    Assert.AreSame(raster, document.CurrentImage);
    Assert.IsTrue(document.TryGetTextAt(new Point(10, 10), out var id));
    Assert.IsTrue(document.MoveText(id, new Point(100, 100)));
    Assert.AreSame(raster, document.CurrentImage);
    Assert.IsTrue(document.UpdateText(id, "updated"));
    Assert.AreSame(raster, document.CurrentImage);
    Assert.IsTrue(document.DeleteText(id));
    Assert.AreSame(raster, document.CurrentImage);
    Assert.IsTrue(document.Undo());
    Assert.AreSame(raster, document.CurrentImage);

    Assert.IsTrue(document.DrawRectangle(new Rectangle(20, 20, 100, 80)));
    Assert.AreNotSame(raster, document.CurrentImage);
  }

  [TestMethod]
  public void Mosaic_ProducesUniformBlocksWithinSelection()
  {
    using var source = new Bitmap(8, 8);
    for (var y = 0; y < source.Height; y++)
    {
      for (var x = 0; x < source.Width; x++)
      {
        source.SetPixel(x, y, Color.FromArgb(255, x * 20, y * 20, (x + y) * 10));
      }
    }

    using var document = new ImageEditDocument(source);
    Assert.IsTrue(document.Mosaic(new Rectangle(0, 0, 8, 8), blockSize: 4));
    using var result = document.GetImageCopy();

    Assert.AreEqual(result.GetPixel(0, 0).ToArgb(), result.GetPixel(3, 3).ToArgb());
    Assert.AreEqual(result.GetPixel(4, 4).ToArgb(), result.GetPixel(7, 7).ToArgb());
  }

  [TestMethod]
  public void InvalidSelections_DoNotCreateHistory()
  {
    using var source = new Bitmap(10, 10);
    using var document = new ImageEditDocument(source);

    Assert.IsFalse(document.DrawArrow(new Point(2, 2), new Point(2, 2)));
    Assert.IsFalse(document.Crop(new Rectangle(1, 1, 1, 1)));
    Assert.IsFalse(document.DrawText("   ", Point.Empty));
    Assert.IsFalse(document.CanUndo);
  }

  [TestMethod]
  public void AnnotationColor_IsAppliedToFrameArrowAndTextLabel()
  {
    using var source = new Bitmap(120, 80);
    using (var graphics = Graphics.FromImage(source))
    {
      graphics.Clear(Color.White);
    }

    using var document = new ImageEditDocument(source);
    Assert.IsTrue(document.DrawRectangle(new Rectangle(5, 5, 50, 40), Color.Blue));
    Assert.IsTrue(document.DrawArrow(new Point(60, 5), new Point(100, 30), Color.Green));
    Assert.IsTrue(document.DrawText("label", new Point(5, 50), Color.Purple));
    using var result = document.GetImageCopy();

    Assert.IsTrue(ContainsColor(result, Color.Blue));
    Assert.IsTrue(ContainsColor(result, Color.Green));
    Assert.IsTrue(ContainsColor(result, Color.Purple));
  }

  [TestMethod]
  public void MoveText_RepositionsTextAndSupportsUndoRedo()
  {
    using var source = new Bitmap(240, 120);
    using (var graphics = Graphics.FromImage(source))
    {
      graphics.Clear(Color.White);
    }

    using var document = new ImageEditDocument(source);
    Assert.IsTrue(document.DrawText("label", new Point(8, 8), Color.Purple));
    Assert.IsTrue(document.TryGetTextAt(new Point(8, 8), out var annotationId));

    Assert.IsTrue(document.MoveText(annotationId, new Point(120, 60)));
    Assert.IsFalse(document.TryGetTextAt(new Point(8, 8), out _));
    Assert.IsTrue(document.TryGetTextAt(new Point(120, 60), out var movedId));
    Assert.AreEqual(annotationId, movedId);

    Assert.IsTrue(document.Undo());
    Assert.IsTrue(document.TryGetTextAt(new Point(8, 8), out var restoredId));
    Assert.AreEqual(annotationId, restoredId);

    Assert.IsTrue(document.Redo());
    Assert.IsTrue(document.TryGetTextAt(new Point(120, 60), out var redoneId));
    Assert.AreEqual(annotationId, redoneId);
  }

  [TestMethod]
  public void FrameAndArrow_DoNotPreventTextFromBeingMoved()
  {
    using var source = new Bitmap(240, 120);
    using var document = new ImageEditDocument(source);
    Assert.IsTrue(document.DrawText("label", new Point(8, 8)));
    Assert.IsTrue(document.TryGetTextAt(new Point(8, 8), out var annotationId));

    Assert.IsTrue(document.DrawRectangle(new Rectangle(80, 10, 60, 40)));
    Assert.IsTrue(document.DrawArrow(new Point(80, 80), new Point(160, 80)));

    Assert.IsTrue(document.MoveText(annotationId, new Point(120, 50)));
    Assert.IsTrue(document.TryGetTextAt(new Point(120, 50), out var movedId));
    Assert.AreEqual(annotationId, movedId);
  }

  [TestMethod]
  public void EditAndDeleteText_PreserveIdentityAndUndoRedo()
  {
    using var source = new Bitmap(400, 200);
    using var document = new ImageEditDocument(source);
    document.DrawText("short", new Point(10, 10), Color.Blue);
    Assert.IsTrue(document.TryGetTextAt(new Point(10, 10), out var id));
    Assert.IsTrue(document.UpdateText(id, "a longer label"));
    Assert.IsTrue(document.TryGetTextAnnotation(id, out var updated));
    Assert.AreEqual(Color.Blue, updated.Color);
    Assert.AreEqual(new Point(10, 10), updated.Location);
    Assert.IsFalse(document.UpdateText(id, "   "));
    Assert.IsTrue(document.DeleteText(id));
    Assert.IsFalse(document.TryGetTextAnnotation(id, out _));
    Assert.IsTrue(document.Undo());
    Assert.IsTrue(document.TryGetTextAnnotation(id, out var restored));
    Assert.AreEqual("a longer label", restored.Text);
    document.Undo();
    Assert.IsTrue(document.TryGetTextAnnotation(id, out restored));
    Assert.AreEqual("short", restored.Text);
    document.Redo();
    document.Redo();
    Assert.IsFalse(document.TryGetTextAnnotation(id, out _));
  }

  [TestMethod]
  public void Crop_KeepsTextEditableAndPreservesFontSize()
  {
    using var source = new Bitmap(600, 400);
    using var document = new ImageEditDocument(source);
    document.DrawText("label", new Point(100, 100));
    document.TryGetTextAt(new Point(100, 100), out var id);
    document.TryGetTextAnnotation(id, out var original);
    Assert.IsTrue(document.Crop(new Rectangle(50, 50, 400, 250)));
    Assert.IsTrue(document.TryGetTextAnnotation(id, out var cropped));
    Assert.AreEqual(new Point(50, 50), cropped.Location);
    Assert.AreEqual(original.FontSize, cropped.FontSize);
    Assert.IsTrue(document.UpdateText(id, "changed"));
    Assert.IsTrue(document.MoveText(id, new Point(20, 20)));
    Assert.IsTrue(document.DeleteText(id));
  }

  [TestMethod]
  public void Mosaic_TextMoveEditDelete_DoesNotUncoverMaskedPixels()
  {
    using var source = new Bitmap(300, 160);
    using (var graphics = Graphics.FromImage(source)) graphics.Clear(Color.White);
    using var document = new ImageEditDocument(source);
    document.DrawText("secret", new Point(10, 10));
    document.TryGetTextAt(new Point(10, 10), out var id);
    var mask = new Rectangle(5, 5, 90, 40);
    Assert.IsTrue(document.Mosaic(mask));
    using var masked = document.GetImageCopy();
    Assert.IsTrue(document.MoveText(id, new Point(150, 90)));
    Assert.IsTrue(document.UpdateText(id, "edited"));
    using var moved = document.GetImageCopy();
    for (var y = 90; y < 110; y++)
      for (var x = 150; x < 190; x++)
        Assert.AreEqual(Color.White.ToArgb(), moved.GetPixel(x, y).ToArgb(), "Moving text must not reveal its masked characters.");
    Assert.IsTrue(document.DeleteText(id));
    using var deleted = document.GetImageCopy();
    for (var y = mask.Top; y < mask.Bottom; y++)
      for (var x = mask.Left; x < mask.Right; x++)
        Assert.AreEqual(masked.GetPixel(x, y), deleted.GetPixel(x, y));
    document.Undo();
    Assert.IsTrue(document.TryGetTextAnnotation(id, out _));
  }

  private static bool ContainsColor(Bitmap image, Color expected)
  {
    for (var y = 0; y < image.Height; y++)
    {
      for (var x = 0; x < image.Width; x++)
      {
        var actual = image.GetPixel(x, y);
        if (Math.Abs(actual.R - expected.R) <= 5 &&
          Math.Abs(actual.G - expected.G) <= 5 &&
          Math.Abs(actual.B - expected.B) <= 5)
        {
          return true;
        }
      }
    }

    return false;
  }

  private static Bitmap CreateQuadrantImage()
  {
    var image = new Bitmap(8, 8);
    using var graphics = Graphics.FromImage(image);
    graphics.Clear(Color.White);
    graphics.FillRectangle(Brushes.Red, 0, 0, 4, 4);
    graphics.FillRectangle(Brushes.Green, 4, 0, 4, 4);
    graphics.FillRectangle(Brushes.Blue, 0, 4, 4, 4);
    graphics.FillRectangle(Brushes.Yellow, 4, 4, 4, 4);
    return image;
  }
}
