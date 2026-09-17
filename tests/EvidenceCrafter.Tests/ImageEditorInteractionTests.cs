using System.Reflection;
using System.Runtime.ExceptionServices;
using EvidenceCrafter.App;

namespace EvidenceCrafter.Tests;

[TestClass]
[DoNotParallelize]
public sealed class ImageEditorInteractionTests
{
  [TestMethod]
  public void TextInput_ModalAddReleasesMouseAndKeepsEditorOpen() => OnSta(() =>
  {
    using var bitmap = new Bitmap(400, 200);
    using var editor = new ImageEditorDialog(bitmap) { TopMost = true };
    editor.Show();
    var canvas = editor.Controls.OfType<RainbowBackdrop>().Single().Controls.OfType<ImageEditorCanvas>().Single();
    canvas.Tool = ImageEditorTool.Text;
    canvas.Capture = true;
    bool? captured = null;
    bool? inputTopmost = null;
    using var timer = new System.Windows.Forms.Timer { Interval = 50 };
    timer.Tick += (_, _) =>
    {
      var input = Application.OpenForms.OfType<ImageTextInputDialog>().SingleOrDefault();
      if (input is null) return;
      timer.Stop();
      captured = canvas.Capture;
      inputTopmost = input.TopMost;
      var field = (TextBox)typeof(ImageTextInputDialog).GetField("textBox", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(input)!;
      field.Text = "追加テキスト";
      typeof(ImageTextInputDialog).GetMethod("ProcessCmdKey", BindingFlags.Instance | BindingFlags.NonPublic)!
        .Invoke(input, [Message.Create(input.Handle, 0x100, (nint)13, 1), Keys.Enter]);
    };
    timer.Start();
    var bounds = Invoke<Rectangle>(canvas, "GetImageBounds");
    Mouse(canvas, "OnMouseDown", new Point(bounds.Left + 20, bounds.Top + 20));
    Assert.IsFalse(captured, "The modal text dialog must receive its own mouse input.");
    Assert.IsTrue(inputTopmost);
    Assert.IsTrue(editor.Visible);
    Assert.AreEqual(DialogResult.None, editor.DialogResult);
    Assert.IsTrue(editor.HasChanges);
  });

  [TestMethod]
  public void TextInput_ShiftEnterAddsLineBreakAndEnterConfirms() => OnSta(() =>
  {
    using var dialog = new ImageTextInputDialog("first");
    var message = Message.Create(0, 0x100, (nint)13, 1);
    var arguments = new object[] { message, Keys.Shift | Keys.Enter };
    var handled = (bool)typeof(ImageTextInputDialog).GetMethod(
      "ProcessCmdKey", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(dialog, arguments)!;
    Assert.IsTrue(handled);
    Assert.AreEqual("first" + Environment.NewLine, dialog.EnteredText);

    message = Message.Create(0, 0x100, (nint)13, 1);
    arguments = [message, Keys.Enter];
    handled = (bool)typeof(ImageTextInputDialog).GetMethod(
      "ProcessCmdKey", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(dialog, arguments)!;
    Assert.IsTrue(handled);
    Assert.AreEqual(DialogResult.OK, dialog.DialogResult);
  });

  [TestMethod]
  public void Editor_ShiftEnterAddsTextAtTheImageCenter() => OnSta(() =>
  {
    using var bitmap = new Bitmap(400, 200);
    using var editor = new ImageEditorDialog(bitmap) { TopMost = true };
    editor.Show();
    using var timer = new System.Windows.Forms.Timer { Interval = 50 };
    timer.Tick += (_, _) =>
    {
      var input = Application.OpenForms.OfType<ImageTextInputDialog>().SingleOrDefault();
      if (input is null) return;
      timer.Stop();
      var field = (TextBox)typeof(ImageTextInputDialog).GetField("textBox", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(input)!;
      field.Text = "中央テキスト";
      typeof(ImageTextInputDialog).GetMethod("ProcessCmdKey", BindingFlags.Instance | BindingFlags.NonPublic)!
        .Invoke(input, [Message.Create(input.Handle, 0x100, (nint)13, 1), Keys.Enter]);
    };
    timer.Start();
    var handled = (bool)typeof(ImageEditorDialog).GetMethod("ProcessCmdKey", BindingFlags.Instance | BindingFlags.NonPublic)!
      .Invoke(editor, [Message.Create(editor.Handle, 0x100, (nint)13, 1), Keys.Shift | Keys.Enter])!;
    Assert.IsTrue(handled);
    var document = (ImageEditDocument)typeof(ImageEditorDialog).GetField("document", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(editor)!;
    Assert.IsTrue(document.TryGetTextAt(new Point(200, 100), out _));
  });

  [TestMethod]
  public void Canvas_TextMovesWithoutMoveTool_DoubleClickEditsAndCrossDeletes() => OnSta(() =>
  {
    using var bitmap = new Bitmap(400, 200);
    using var document = new ImageEditDocument(bitmap);
    document.DrawText("label", new Point(20, 20));
    document.TryGetTextAt(new Point(20, 20), out var id);
    using var canvas = new ImageEditorCanvas(document) { Size = new Size(800, 500), Tool = ImageEditorTool.Rectangle };
    var bounds = Invoke<Rectangle>(canvas, "GetImageBounds");
    Point Client(Point point) => new(bounds.Left + point.X * bounds.Width / document.Width, bounds.Top + point.Y * bounds.Height / document.Height);
    var start = Client(new Point(25, 25));
    var finish = Client(new Point(145, 85));
    Mouse(canvas, "OnMouseDown", start);
    Mouse(canvas, "OnMouseMove", finish);
    Mouse(canvas, "OnMouseUp", finish);
    Assert.IsTrue(document.TryGetTextAnnotation(id, out var moved));
    Assert.IsTrue(Math.Abs(moved.Location.X - 140) <= 1 && Math.Abs(moved.Location.Y - 80) <= 1,
      "Dragging must preserve the grab offset within one image pixel of display rounding.");
    var originalFontSize = moved.FontSize;
    var resizeHandle = Invoke<Rectangle>(canvas, "GetResizeBounds");
    var resizeStart = new Point(resizeHandle.Left + resizeHandle.Width / 2, resizeHandle.Top + resizeHandle.Height / 2);
    var resizeEnd = new Point(resizeStart.X + 80, resizeStart.Y + 40);
    Mouse(canvas, "OnMouseDown", resizeStart);
    Mouse(canvas, "OnMouseMove", resizeEnd);
    Mouse(canvas, "OnMouseUp", resizeEnd);
    Assert.IsTrue(document.TryGetTextAnnotation(id, out var resized));
    Assert.IsGreaterThan(originalFontSize, resized.FontSize);
    document.Undo();
    Assert.IsTrue(document.TryGetTextAnnotation(id, out var restored));
    Assert.AreEqual(originalFontSize, restored.FontSize);
    document.Redo();
    Guid? editedId = null;
    canvas.TextEditRequested += (_, selected) => editedId = selected;
    Mouse(canvas, "OnMouseDown", finish, clicks: 2);
    Assert.AreEqual(id, editedId);
    var cross = Invoke<Rectangle>(canvas, "GetDeleteBounds");
    Assert.IsTrue(canvas.ClientRectangle.Contains(cross));
    Mouse(canvas, "OnMouseDown", new Point(cross.Left + 2, cross.Top + 2));
    Assert.IsFalse(document.TryGetTextAnnotation(id, out _));
    document.Undo();
    Assert.IsTrue(document.TryGetTextAnnotation(id, out _));
  });

  [TestMethod]
  public void Canvas_CancelDragPreservesTextAndHistory() => OnSta(() =>
  {
    using var bitmap = new Bitmap(400, 200);
    using var document = new ImageEditDocument(bitmap);
    document.DrawText("label", new Point(20, 20));
    document.TryGetTextAt(new Point(20, 20), out var id);
    using var canvas = new ImageEditorCanvas(document) { Size = new Size(416, 216) };
    Mouse(canvas, "OnMouseDown", new Point(30, 30));
    Mouse(canvas, "OnMouseMove", new Point(150, 90));
    Assert.IsTrue(canvas.CancelDrag());
    document.TryGetTextAnnotation(id, out var annotation);
    Assert.AreEqual(new Point(20, 20), annotation.Location);
    document.Undo();
    Assert.IsFalse(document.TryGetTextAnnotation(id, out _), "Cancelled dragging must not create a history entry.");
  });

  private static void Mouse(ImageEditorCanvas canvas, string method, Point point, int clicks = 1) =>
    typeof(ImageEditorCanvas).GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic)!
      .Invoke(canvas, [new MouseEventArgs(MouseButtons.Left, clicks, point.X, point.Y, 0)]);

  private static T Invoke<T>(object target, string method) =>
    (T)target.GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(target, null)!;

  private static void OnSta(Action action)
  {
    Exception? failure = null;
    var thread = new Thread(() => { try { action(); } catch (Exception exception) { failure = exception; } });
    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(15)), "Editor test timed out.");
    if (failure is not null) ExceptionDispatchInfo.Capture(failure).Throw();
  }
}
