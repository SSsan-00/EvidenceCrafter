using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Excel;

/// <summary>Inspects and mutates only the single EvidenceCrafter image selected in Excel.</summary>
public sealed class ExcelManagedShapeService
{
  private const int MsoFalse = 0;
  private const int MsoTrue = -1;
  private const int MsoLinkedPicture = 11;
  private const int MsoPicture = 13;
  private const int XlMove = 2;
  private readonly ImageSizingService imageSizingService = new();

  public ManagedShapeSelectionResult Inspect(WorkbookIdentity workbook, string worksheetName, string shapeName)
  {
    var validation = ValidateSession(workbook, requiresWrite: false);
    if (validation is not null) return ManagedShapeSelectionResult.Failed(validation);
    return WithWorkbook(workbook, ManagedShapeSelectionResult.Failed, (_, candidate) =>
    {
      object? sheet = null, shapes = null, shape = null;
      try
      {
        sheet = ResolveWorksheet(candidate, worksheetName);
        shapes = GetRequiredProperty(sheet, "Shapes");
        shape = InvokeMethod(shapes, "Item", shapeName)!;
        return TryReadManagedTarget(shape, workbook, out var target, out var reason)
          ? new ManagedShapeSelectionResult(true, target, string.Empty)
          : ManagedShapeSelectionResult.Failed(reason);
      }
      finally { ComRelease.Release(shape); ComRelease.Release(shapes); ComRelease.Release(sheet); }
    });
  }

  /// <summary>Changes geometry in place, preserving picture data and validating the expected state.</summary>
  public ManagedShapeMutationResult Resize(WorkbookIdentity workbook, ManagedShapeTarget expected, ManagedShapeTarget desired)
  {
    var validation = ValidateSession(workbook, requiresWrite: true);
    if (validation is not null) return ManagedShapeMutationResult.Failed(validation);
    if (!IsValidImage(new(desired.WidthPoints, desired.HeightPoints)))
      return ManagedShapeMutationResult.Failed("画像サイズが不正です。");
    return WithWorkbook(workbook, ManagedShapeMutationResult.Failed, (application, candidate) =>
    {
      object? sheet = null, shapes = null, shape = null;
      var eventsEnabled = ReadBoolean(application, "EnableEvents");
      var changed = false;
      void Apply(ManagedShapeTarget target)
      {
        SetProperty(shape!, "LockAspectRatio", MsoFalse);
        SetProperty(shape!, "Width", target.WidthPoints);
        SetProperty(shape!, "Height", target.HeightPoints);
        SetProperty(shape!, "AlternativeText", target.AlternativeText);
        SetProperty(shape!, "LockAspectRatio", MsoTrue);
      }
      try
      {
        sheet = ResolveWorksheet(candidate, expected.WorksheetName);
        if (!CanMutateWorkbook(candidate, workbook) || IsWorksheetProtected(sheet))
          return ManagedShapeMutationResult.Failed("対象ブックまたはシートは変更できません。");
        shapes = GetRequiredProperty(sheet, "Shapes");
        shape = InvokeMethod(shapes, "Item", expected.ShapeName)!;
        if (!TryReadManagedTarget(shape, workbook, out var current, out _) || !TargetUnchanged(expected, current))
          return ManagedShapeMutationResult.Failed("参照画像が変更されたため倍率変更を停止しました。") with { ReferenceChanged = true };
        var count = Convert.ToInt32(GetRequiredProperty(shapes, "Count"), CultureInfo.InvariantCulture);
        for (var i = 1; i <= count; i++)
        {
          var other = InvokeMethod(shapes, "Item", i)!;
          try
          {
            if (Equals(GetRequiredProperty(other, "Name"), expected.ShapeName)) continue;
            var left = ReadFiniteDouble(other, "Left");
            var top = ReadFiniteDouble(other, "Top");
            if (current.LeftPoints < left + ReadFiniteDouble(other, "Width") - 0.05 &&
              current.LeftPoints + desired.WidthPoints > left + 0.05 &&
              current.TopPoints < top + ReadFiniteDouble(other, "Height") - 0.05 &&
              current.TopPoints + desired.HeightPoints > top + 0.05)
              return ManagedShapeMutationResult.Failed("倍率変更後の画像が既存の図形と重なるため停止しました。");
          }
          finally { ComRelease.Release(other); }
        }
        SetProperty(application, "EnableEvents", false);
        changed = true;
        Apply(desired);
        if (!TryReadManagedTarget(shape, workbook, out var after, out _) ||
          !NearlyEqual(after.WidthPoints, desired.WidthPoints) || !NearlyEqual(after.HeightPoints, desired.HeightPoints))
          throw new InvalidOperationException("画像サイズを検証できません。");
        return new ManagedShapeMutationResult(true, true, current, after, "参照画像の倍率を変更しました。");
      }
      catch (Exception exception) when (IsAutomationFailure(exception))
      {
        if (changed)
        {
          try { Apply(expected); }
          catch (Exception restoreError) when (IsAutomationFailure(restoreError))
          { return ManagedShapeMutationResult.Failed("参照画像の倍率変更と復元に失敗しました。保存せず状態を確認してください。") with { CompensationSucceeded = false }; }
        }
        return ManagedShapeMutationResult.Failed($"参照画像の倍率変更に失敗しました: {exception.Message}");
      }
      finally
      {
        TryRestoreEvents(application, eventsEnabled);
        ComRelease.Release(shape); ComRelease.Release(shapes); ComRelease.Release(sheet);
      }
    });
  }

  public ManagedShapeSelectionResult InspectSelection(WorkbookIdentity workbook)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    var validation = ValidateSession(workbook, requiresWrite: false);
    if (validation is not null)
    {
      return ManagedShapeSelectionResult.Failed(validation);
    }

    return WithWorkbook(
      workbook,
      ManagedShapeSelectionResult.Failed,
      (application, candidate) => InspectWorkbookSelection(application, candidate, workbook));
  }

  public ManagedShapeMutationResult Replace(
    WorkbookIdentity workbook,
    ManagedShapeTarget expectedTarget,
    string imagePath,
    ImageDimensions imageDimensions,
    ManagedShapeTarget? exactGeometry = null)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentNullException.ThrowIfNull(expectedTarget);
    ArgumentException.ThrowIfNullOrWhiteSpace(imagePath);
    if (!File.Exists(imagePath))
    {
      return ManagedShapeMutationResult.Failed("差し替え画像が見つかりません。");
    }

    if (!IsValidImage(imageDimensions))
    {
      throw new ArgumentOutOfRangeException(nameof(imageDimensions));
    }

    var validation = ValidateSession(workbook, requiresWrite: true);
    if (validation is not null)
    {
      return ManagedShapeMutationResult.Failed(validation);
    }

    return WithWorkbook(
      workbook,
      ManagedShapeMutationResult.Failed,
      (application, candidate) => ReplaceInWorkbook(
        application,
        candidate,
        workbook,
        expectedTarget,
        imagePath,
        imageDimensions,
        exactGeometry));
  }

  public ManagedShapeMutationResult Delete(
    WorkbookIdentity workbook,
    ManagedShapeTarget expectedTarget)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentNullException.ThrowIfNull(expectedTarget);
    var validation = ValidateSession(workbook, requiresWrite: true);
    if (validation is not null)
    {
      return ManagedShapeMutationResult.Failed(validation);
    }

    return WithWorkbook(
      workbook,
      ManagedShapeMutationResult.Failed,
      (application, candidate) => DeleteFromWorkbook(application, candidate, workbook, expectedTarget));
  }

  /// <summary>Exports the exact managed picture so a destructive operation can be undone.</summary>
  public ManagedShapeExportResult Export(
    WorkbookIdentity workbook,
    ManagedShapeTarget expectedTarget,
    string outputPath)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentNullException.ThrowIfNull(expectedTarget);
    ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
    var validation = ValidateSession(workbook, requiresWrite: true);
    if (validation is not null)
    {
      return ManagedShapeExportResult.Failed(validation);
    }

    return WithWorkbook(
      workbook,
      ManagedShapeExportResult.Failed,
      (application, candidate) => ExportFromWorkbook(application, candidate, workbook, expectedTarget, outputPath));
  }

  /// <summary>Restores an exported picture at its original geometry and metadata.</summary>
  public ManagedShapeMutationResult Restore(
    WorkbookIdentity workbook,
    ManagedShapeTarget originalTarget,
    string imagePath)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentNullException.ThrowIfNull(originalTarget);
    ArgumentException.ThrowIfNullOrWhiteSpace(imagePath);
    if (!File.Exists(imagePath))
    {
      return ManagedShapeMutationResult.Failed("復元画像が見つかりません。");
    }

    var validation = ValidateSession(workbook, requiresWrite: true);
    if (validation is not null)
    {
      return ManagedShapeMutationResult.Failed(validation);
    }

    return WithWorkbook(
      workbook,
      ManagedShapeMutationResult.Failed,
      (application, candidate) => RestoreInWorkbook(application, candidate, workbook, originalTarget, imagePath));
  }

  private static ManagedShapeSelectionResult InspectWorkbookSelection(
    object application,
    object workbook,
    WorkbookIdentity identity)
  {
    if (!WorkbookWindowMatchesIdentity(workbook, identity))
    {
      return ManagedShapeSelectionResult.Failed(
        "Workbookが閉じられたか再オープンされています。更新して選択し直してください。");
    }

    object? activeWorkbook = null;
    object? selection = null;
    object? shapeRange = null;
    object? shape = null;
    try
    {
      activeWorkbook = GetRequiredProperty(application, "ActiveWorkbook");
      if (!WorkbookMatches(activeWorkbook, identity))
      {
        return ManagedShapeSelectionResult.Failed("対象WorkbookをExcelでアクティブにして画像を1つ選択してください。");
      }

      selection = GetRequiredProperty(application, "Selection");
      if (!TryGetProperty(selection, "ShapeRange", out shapeRange) || shapeRange is null)
      {
        return ManagedShapeSelectionResult.Failed("ExcelでEvidenceCrafter画像を1つ選択してください。");
      }

      var count = Convert.ToInt32(GetRequiredProperty(shapeRange, "Count"), CultureInfo.InvariantCulture);
      if (count != 1)
      {
        return ManagedShapeSelectionResult.Failed("複数選択では変更できません。EvidenceCrafter画像を1つだけ選択してください。");
      }

      shape = InvokeMethod(shapeRange, "Item", 1);
      var reason = "選択されたShapeを検証できません。";
      if (shape is null || !TryReadManagedTarget(shape, identity, out var target, out reason))
      {
        return ManagedShapeSelectionResult.Failed(reason);
      }

      return new ManagedShapeSelectionResult(true, target, "管理画像を選択しました。");
    }
    finally
    {
      ComRelease.Release(shape);
      ComRelease.Release(shapeRange);
      ComRelease.Release(selection);
      ComRelease.Release(activeWorkbook);
    }
  }

  private ManagedShapeMutationResult ReplaceInWorkbook(
    object application,
    object workbook,
    WorkbookIdentity identity,
    ManagedShapeTarget expected,
    string imagePath,
    ImageDimensions imageDimensions,
    ManagedShapeTarget? exactGeometry)
  {
    if (!CanMutateWorkbook(workbook, identity))
    {
      return ManagedShapeMutationResult.Failed("Workbookの接続世代または読み取り専用状態が変わりました。");
    }

    var eventsWereEnabled = ReadBoolean(application, "EnableEvents");

    object? worksheet = null;
    object? shapes = null;
    object? oldShape = null;
    object? replacementShape = null;
    var replacementCreated = false;
    var oldShapeDeleted = false;
    try
    {
      SetProperty(application, "EnableEvents", false);
      worksheet = ResolveWorksheet(workbook, expected.WorksheetName);
      if (IsWorksheetProtected(worksheet))
      {
        return ManagedShapeMutationResult.Failed($"シート {expected.WorksheetName} は保護されています。");
      }

      shapes = GetRequiredProperty(worksheet, "Shapes");
      oldShape = InvokeMethod(shapes, "Item", expected.ShapeName);
      var reason = "差し替え対象を検証できません。";
      if (oldShape is null ||
        !TryReadManagedTarget(oldShape, identity, out var current, out reason) ||
        !TargetUnchanged(expected, current))
      {
        return ManagedShapeMutationResult.Failed(
          oldShape is null ? "差し替え対象が見つかりません。" : $"差し替え対象が変更されています。{reason}");
      }

      if (!CanMutateWorkbook(workbook, identity))
      {
        return ManagedShapeMutationResult.Failed("差し替え直前にWorkbookの状態が変わりました。");
      }

      var fitted = exactGeometry is null
        ? imageSizingService.FitToWidth(imageDimensions, current.WidthPoints)
        : new FittedImage(exactGeometry.WidthPoints, exactGeometry.HeightPoints, 1);
      var replacementName = $"EST_IMG_{Guid.NewGuid():N}";
      replacementShape = InvokeMethod(
        shapes,
        "AddPicture",
        imagePath,
        MsoFalse,
        MsoTrue,
        exactGeometry?.LeftPoints ?? current.LeftPoints,
        exactGeometry?.TopPoints ?? current.TopPoints,
        fitted.WidthPoints,
        fitted.HeightPoints);
      if (replacementShape is null)
      {
        return ManagedShapeMutationResult.Failed("Excelが差し替え画像を作成できませんでした。");
      }

      replacementCreated = true;
      SetProperty(replacementShape, "Name", replacementName);
      var replacementMetadata = (exactGeometry?.Metadata ?? current.Metadata) with
      {
        SourceDimensions = imageDimensions,
        AppliedScale = exactGeometry is null
          ? fitted.Scale
          : exactGeometry.WidthPoints / imageDimensions.WidthPoints,
      };
      SetProperty(replacementShape, "AlternativeText", replacementMetadata.Serialize());
      SetProperty(replacementShape, "LockAspectRatio", MsoTrue);
      SetProperty(replacementShape, "Placement", XlMove);
      if (!TryReadManagedTarget(replacementShape, identity, out var replacement, out _) ||
        !string.Equals(replacement.ShapeName, replacementName, StringComparison.Ordinal))
      {
        return ManagedShapeMutationResult.Failed("差し替え画像を検証できませんでした。");
      }

      if (!CanMutateWorkbook(workbook, identity) ||
        !TryReadManagedTarget(oldShape, identity, out var finalCurrent, out _) ||
        !TargetUnchanged(expected, finalCurrent))
      {
        return ManagedShapeMutationResult.Failed("差し替え確定前に対象またはWorkbookが変更されました。");
      }

      InvokeMethod(oldShape, "Delete");
      oldShapeDeleted = true;
      return new ManagedShapeMutationResult(
        true,
        true,
        current,
        replacement,
        $"{expected.WorksheetName} の管理画像を差し替えました（未保存）。");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return ManagedShapeMutationResult.Failed(
        $"管理画像の差し替えに失敗しました (0x{GetAutomationHResult(exception):X8})。変更結果を確認してください。");
    }
    finally
    {
      if (replacementCreated && !oldShapeDeleted)
      {
        TryDeleteShape(replacementShape);
      }

      TryRestoreEvents(application, eventsWereEnabled);
      ComRelease.Release(replacementShape);
      ComRelease.Release(oldShape);
      ComRelease.Release(shapes);
      ComRelease.Release(worksheet);
    }
  }

  private static ManagedShapeMutationResult DeleteFromWorkbook(
    object application,
    object workbook,
    WorkbookIdentity identity,
    ManagedShapeTarget expected)
  {
    if (!CanMutateWorkbook(workbook, identity))
    {
      return ManagedShapeMutationResult.Failed("Workbookの接続世代または読み取り専用状態が変わりました。");
    }

    var eventsWereEnabled = ReadBoolean(application, "EnableEvents");

    object? worksheet = null;
    object? shapes = null;
    object? shape = null;
    try
    {
      SetProperty(application, "EnableEvents", false);
      worksheet = ResolveWorksheet(workbook, expected.WorksheetName);
      if (IsWorksheetProtected(worksheet))
      {
        return ManagedShapeMutationResult.Failed($"シート {expected.WorksheetName} は保護されています。");
      }

      shapes = GetRequiredProperty(worksheet, "Shapes");
      shape = InvokeMethod(shapes, "Item", expected.ShapeName);
      var reason = "削除対象を検証できません。";
      if (shape is null ||
        !TryReadManagedTarget(shape, identity, out var current, out reason) ||
        !TargetUnchanged(expected, current))
      {
        return ManagedShapeMutationResult.Failed(
          shape is null ? "削除対象が見つかりません。" : $"削除対象が変更されています。{reason}");
      }

      if (!CanMutateWorkbook(workbook, identity))
      {
        return ManagedShapeMutationResult.Failed("削除直前にWorkbookの状態が変わりました。");
      }

      InvokeMethod(shape, "Delete");
      return new ManagedShapeMutationResult(
        true,
        true,
        current,
        null,
        $"{expected.WorksheetName} の管理画像を削除しました（未保存）。");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return ManagedShapeMutationResult.Failed(
        $"管理画像の削除に失敗しました (0x{GetAutomationHResult(exception):X8})。変更結果を確認してください。");
    }
    finally
    {
      TryRestoreEvents(application, eventsWereEnabled);
      ComRelease.Release(shape);
      ComRelease.Release(shapes);
      ComRelease.Release(worksheet);
    }
  }

  private static ManagedShapeExportResult ExportFromWorkbook(
    object application,
    object workbook,
    WorkbookIdentity identity,
    ManagedShapeTarget expected,
    string outputPath)
  {
    if (!CanMutateWorkbook(workbook, identity))
    {
      return ManagedShapeExportResult.Failed("Workbookの接続世代または状態が変わりました。");
    }

    var eventsWereEnabled = ReadBoolean(application, "EnableEvents");
    var workbookWasSaved = ReadBoolean(workbook, "Saved");
    object? worksheet = null;
    object? shapes = null;
    object? shape = null;
    object? chartObjects = null;
    object? chartObject = null;
    object? chart = null;
    var exportStage = "validating the managed image";
    try
    {
      SetProperty(application, "EnableEvents", false);
      worksheet = ResolveWorksheet(workbook, expected.WorksheetName);
      if (IsWorksheetProtected(worksheet))
      {
        return ManagedShapeExportResult.Failed($"シート {expected.WorksheetName} は保護されています。");
      }
      shapes = GetRequiredProperty(worksheet, "Shapes");
      shape = InvokeMethod(shapes, "Item", expected.ShapeName);
      var reason = "出力対象を検証できません。";
      if (shape is null ||
        !TryReadManagedTarget(shape, identity, out var current, out reason) ||
        !TargetUnchanged(expected, current))
      {
        return ManagedShapeExportResult.Failed(reason);
      }

      Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath))!);
      exportStage = "copying the managed image";
      InvokeMethod(shape, "CopyPicture", 1, 2);
      exportStage = "creating the temporary export chart";
      chartObjects = InvokeMethod(worksheet, "ChartObjects");
      chartObject = InvokeMethod(chartObjects!, "Add", 0d, 0d, current.WidthPoints, current.HeightPoints);
      InvokeMethod(chartObject!, "Activate");
      chart = GetRequiredProperty(chartObject!, "Chart");
      exportStage = "pasting the managed image into the temporary chart";
      InvokeMethod(chart, "Paste");
      exportStage = "exporting the temporary chart as PNG";
      var exported = Convert.ToBoolean(InvokeMethod(chart, "Export", outputPath, "PNG"), CultureInfo.InvariantCulture);
      return exported && File.Exists(outputPath) && new FileInfo(outputPath).Length > 0
        ? new ManagedShapeExportResult(true, current, outputPath, "管理画像を一時退避しました。")
        : ManagedShapeExportResult.Failed("管理画像を復元用に出力できませんでした。");
    }
    catch (Exception exception) when (IsAutomationFailure(exception) || exception is IOException)
    {
      return ManagedShapeExportResult.Failed(
        $"管理画像の一時退避に失敗しました: {exportStage} (0x{GetAutomationHResult(exception):X8})。");
    }
    finally
    {
      if (chartObject is not null)
      {
        try { InvokeMethod(chartObject, "Delete"); } catch (Exception exception) when (IsAutomationFailure(exception)) { }
      }

      if (workbookWasSaved)
      {
        try { SetProperty(workbook, "Saved", true); } catch (Exception exception) when (IsAutomationFailure(exception)) { }
      }
      TryRestoreEvents(application, eventsWereEnabled);

      ComRelease.Release(chart);
      ComRelease.Release(chartObject);
      ComRelease.Release(chartObjects);
      ComRelease.Release(shape);
      ComRelease.Release(shapes);
      ComRelease.Release(worksheet);
    }
  }

  private static ManagedShapeMutationResult RestoreInWorkbook(
    object application,
    object workbook,
    WorkbookIdentity identity,
    ManagedShapeTarget original,
    string imagePath)
  {
    if (!CanMutateWorkbook(workbook, identity))
    {
      return ManagedShapeMutationResult.Failed("Workbookの接続世代または読み取り専用状態が変わりました。");
    }

    var eventsWereEnabled = ReadBoolean(application, "EnableEvents");
    object? worksheet = null;
    object? shapes = null;
    object? shape = null;
    try
    {
      SetProperty(application, "EnableEvents", false);
      worksheet = ResolveWorksheet(workbook, original.WorksheetName);
      if (IsWorksheetProtected(worksheet))
      {
        return ManagedShapeMutationResult.Failed($"シート {original.WorksheetName} は保護されています。");
      }

      shapes = GetRequiredProperty(worksheet, "Shapes");
      var restoredName = $"EST_IMG_{Guid.NewGuid():N}";
      shape = InvokeMethod(
        shapes,
        "AddPicture",
        imagePath,
        MsoFalse,
        MsoTrue,
        original.LeftPoints,
        original.TopPoints,
        original.WidthPoints,
        original.HeightPoints);
      if (shape is null)
      {
        return ManagedShapeMutationResult.Failed("管理画像を復元できませんでした。");
      }

      SetProperty(shape, "Name", restoredName);
      SetProperty(shape, "AlternativeText", original.AlternativeText);
      SetProperty(shape, "LockAspectRatio", MsoTrue);
      SetProperty(shape, "Placement", XlMove);
      if (!TryReadManagedTarget(shape, identity, out var restored, out _))
      {
        TryDeleteShape(shape);
        return ManagedShapeMutationResult.Failed("復元した管理画像を検証できませんでした。");
      }

      return new ManagedShapeMutationResult(true, true, null, restored, "管理画像を復元しました（未保存）。");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      TryDeleteShape(shape);
      return ManagedShapeMutationResult.Failed($"管理画像の復元に失敗しました (0x{GetAutomationHResult(exception):X8})。");
    }
    finally
    {
      TryRestoreEvents(application, eventsWereEnabled);
      ComRelease.Release(shape);
      ComRelease.Release(shapes);
      ComRelease.Release(worksheet);
    }
  }

  private static bool TryReadManagedTarget(
    object shape,
    WorkbookIdentity identity,
    out ManagedShapeTarget target,
    out string reason)
  {
    target = null!;
    reason = "選択対象はEvidenceCrafter管理画像ではありません。";
    var shapeType = Convert.ToInt32(GetRequiredProperty(shape, "Type"), CultureInfo.InvariantCulture);
    if (shapeType is not (MsoPicture or MsoLinkedPicture))
    {
      return false;
    }

    var name = Convert.ToString(GetRequiredProperty(shape, "Name"), CultureInfo.InvariantCulture) ?? string.Empty;
    var alternativeText = Convert.ToString(
      GetRequiredProperty(shape, "AlternativeText"),
      CultureInfo.InvariantCulture) ?? string.Empty;
    if (!ManagedShapeMetadata.IsManagedName(name) ||
      !ManagedShapeMetadata.TryParse(alternativeText, out var metadata))
    {
      return false;
    }

    object? worksheet = null;
    object? parentWorkbook = null;
    object? topLeftCell = null;
    object? pictureFormat = null;
    try
    {
      worksheet = GetRequiredProperty(shape, "Parent");
      parentWorkbook = GetRequiredProperty(worksheet, "Parent");
      if (!WorkbookMatches(parentWorkbook, identity))
      {
        reason = "選択対象は指定Workbookに属していません。";
        return false;
      }

      var worksheetName = Convert.ToString(
        GetRequiredProperty(worksheet, "Name"),
        CultureInfo.CurrentCulture) ?? string.Empty;
      topLeftCell = GetRequiredProperty(shape, "TopLeftCell");
      pictureFormat = GetRequiredProperty(shape, "PictureFormat");
      var rotation = ReadFiniteDouble(shape, "Rotation");
      var cropLeft = ReadFiniteDouble(pictureFormat, "CropLeft");
      var cropTop = ReadFiniteDouble(pictureFormat, "CropTop");
      var cropRight = ReadFiniteDouble(pictureFormat, "CropRight");
      var cropBottom = ReadFiniteDouble(pictureFormat, "CropBottom");
      if (!NearlyEqual(rotation, 0) ||
        !NearlyEqual(cropLeft, 0) || !NearlyEqual(cropTop, 0) ||
        !NearlyEqual(cropRight, 0) || !NearlyEqual(cropBottom, 0))
      {
        reason = "回転またはトリミングされた管理画像は安全に復元できないため変更できません。";
        return false;
      }
      target = new ManagedShapeTarget(
        name,
        worksheetName,
        alternativeText,
        metadata,
        new CellReference(
          Convert.ToInt32(GetRequiredProperty(topLeftCell, "Row"), CultureInfo.InvariantCulture),
          Convert.ToInt32(GetRequiredProperty(topLeftCell, "Column"), CultureInfo.InvariantCulture)),
        ReadFiniteDouble(shape, "Left"),
        ReadFiniteDouble(shape, "Top"),
        ReadPositiveDouble(shape, "Width"),
        ReadPositiveDouble(shape, "Height"),
        rotation,
        cropLeft,
        cropTop,
        cropRight,
        cropBottom);
      reason = string.Empty;
      return true;
    }
    finally
    {
      ComRelease.Release(topLeftCell);
      ComRelease.Release(pictureFormat);
      ComRelease.Release(parentWorkbook);
      ComRelease.Release(worksheet);
    }
  }

  internal static bool TargetUnchanged(ManagedShapeTarget expected, ManagedShapeTarget current) =>
    string.Equals(expected.ShapeName, current.ShapeName, StringComparison.Ordinal) &&
    string.Equals(expected.WorksheetName, current.WorksheetName, StringComparison.Ordinal) &&
    string.Equals(expected.AlternativeText, current.AlternativeText, StringComparison.Ordinal) &&
    expected.TopLeftCell == current.TopLeftCell &&
    NearlyEqual(expected.LeftPoints, current.LeftPoints) &&
    NearlyEqual(expected.TopPoints, current.TopPoints) &&
    NearlyEqual(expected.WidthPoints, current.WidthPoints) &&
    NearlyEqual(expected.HeightPoints, current.HeightPoints) &&
    NearlyEqual(expected.Rotation, current.Rotation) &&
    NearlyEqual(expected.CropLeft, current.CropLeft) &&
    NearlyEqual(expected.CropTop, current.CropTop) &&
    NearlyEqual(expected.CropRight, current.CropRight) &&
    NearlyEqual(expected.CropBottom, current.CropBottom);

  private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.05;

  private static object ResolveWorksheet(object workbook, string worksheetName)
  {
    var worksheets = GetRequiredProperty(workbook, "Worksheets");
    try
    {
      return InvokeProperty(worksheets, "Item", worksheetName) ??
        throw new InvalidOperationException($"Worksheet was not found: {worksheetName}");
    }
    finally
    {
      ComRelease.Release(worksheets);
    }
  }

  private static T WithWorkbook<T>(
    WorkbookIdentity identity,
    Func<string, T> failed,
    Func<object, object, T> action)
  {
    IRunningObjectTable? runningObjectTable = null;
    IEnumMoniker? monikerEnumerator = null;
    IBindCtx? bindContext = null;
    try
    {
      Marshal.ThrowExceptionForHR(NativeMethods.GetRunningObjectTable(0, out runningObjectTable));
      runningObjectTable.EnumRunning(out monikerEnumerator);
      Marshal.ThrowExceptionForHR(NativeMethods.CreateBindCtx(0, out bindContext));
      var monikers = new IMoniker[1];
      while (monikerEnumerator.Next(1, monikers, IntPtr.Zero) == 0)
      {
        object? runningObject = null;
        try
        {
          monikers[0].GetDisplayName(bindContext, null, out var displayName);
          if (!string.Equals(displayName, identity.RotMonikerDisplayName, StringComparison.Ordinal))
          {
            continue;
          }

          runningObjectTable.GetObject(monikers[0], out runningObject);
          if (runningObject is null)
          {
            return failed("対象WorkbookはExcelで利用できなくなりました。");
          }

          var result = TryUseRunningObject(runningObject, identity, action);
          return result.HasValue
            ? result.Value
            : failed("対象WorkbookをExcelインスタンス内で再検証できませんでした。");
        }
        finally
        {
          ComRelease.Release(runningObject);
          ComRelease.Release(monikers[0]);
          monikers[0] = null!;
        }
      }

      return failed("対象WorkbookはRunning Object Tableから切断されています。");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return failed($"Excel操作に失敗しました (0x{GetAutomationHResult(exception):X8})。");
    }
    finally
    {
      ComRelease.Release(bindContext);
      ComRelease.Release(monikerEnumerator);
      ComRelease.Release(runningObjectTable);
    }
  }

  private static Optional<T> TryUseRunningObject<T>(
    object runningObject,
    WorkbookIdentity identity,
    Func<object, object, T> action)
  {
    if (TryGetProperty(runningObject, "Workbooks", out var workbooks) && workbooks is not null)
    {
      try
      {
        if (!ApplicationMatches(runningObject, identity))
        {
          return default;
        }

        var count = Convert.ToInt32(GetRequiredProperty(workbooks, "Count"), CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
          object? candidate = null;
          try
          {
            candidate = InvokeProperty(workbooks, "Item", index);
            if (candidate is not null && WorkbookMatches(candidate, identity))
            {
              return new Optional<T>(action(runningObject, candidate));
            }
          }
          finally
          {
            ComRelease.Release(candidate);
          }
        }

        return default;
      }
      finally
      {
        ComRelease.Release(workbooks);
      }
    }

    if (!TryGetProperty(runningObject, "Application", out var application) || application is null)
    {
      return default;
    }

    try
    {
      return ApplicationMatches(application, identity) && WorkbookMatches(runningObject, identity)
        ? new Optional<T>(action(application, runningObject))
        : default;
    }
    finally
    {
      ComRelease.Release(application);
    }
  }

  private static string? ValidateSession(WorkbookIdentity workbook, bool requiresWrite)
  {
    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return "Excel操作はSTAスレッドで実行する必要があります。";
    }

    if (requiresWrite && workbook.IsReadOnly)
    {
      return "対象Workbookは読み取り専用です。";
    }

    if (string.IsNullOrWhiteSpace(workbook.RotMonikerDisplayName))
    {
      return "Workbookの接続を検証できません。一覧を更新してください。";
    }

    if (workbook.WindowSessionToken == IntPtr.Zero)
    {
      return "Workbook接続トークンがありません。一覧を更新してください。";
    }

    return null;
  }

  private static bool IsValidImage(ImageDimensions dimensions) =>
    double.IsFinite(dimensions.WidthPoints) &&
    double.IsFinite(dimensions.HeightPoints) &&
    dimensions.WidthPoints > 0 &&
    dimensions.HeightPoints > 0;

  private static bool CanMutateWorkbook(object workbook, WorkbookIdentity identity) =>
    WorkbookWindowMatchesIdentity(workbook, identity) &&
    !ReadBoolean(workbook, "ReadOnly");

  private static bool IsWorksheetProtected(object worksheet) =>
    ReadBoolean(worksheet, "ProtectContents") ||
    ReadBoolean(worksheet, "ProtectDrawingObjects") ||
    ReadBoolean(worksheet, "ProtectScenarios");

  private static bool ApplicationMatches(object application, WorkbookIdentity identity)
  {
    var windowHandle = new IntPtr(Convert.ToInt64(
      GetRequiredProperty(application, "Hwnd"),
      CultureInfo.InvariantCulture));
    var threadId = NativeMethods.GetWindowThreadProcessId(windowHandle, out var processId);
    return threadId != 0 && processId == identity.ProcessId;
  }

  private static bool WorkbookMatches(object workbook, WorkbookIdentity identity)
  {
    var name = Convert.ToString(GetRequiredProperty(workbook, "Name"), CultureInfo.CurrentCulture);
    var fullName = Convert.ToString(GetRequiredProperty(workbook, "FullName"), CultureInfo.CurrentCulture);
    return string.Equals(name, identity.Name, StringComparison.OrdinalIgnoreCase) &&
      string.Equals(fullName, identity.FullPath, StringComparison.OrdinalIgnoreCase);
  }

  private static bool WorkbookWindowMatchesIdentity(object workbook, WorkbookIdentity identity)
  {
    if (identity.WindowSessionToken == IntPtr.Zero)
    {
      return false;
    }

    object? windows = null;
    object? window = null;
    try
    {
      windows = GetRequiredProperty(workbook, "Windows");
      var count = Convert.ToInt32(GetRequiredProperty(windows, "Count"), CultureInfo.InvariantCulture);
      if (count < 1)
      {
        return false;
      }

      window = InvokeProperty(windows, "Item", 1);
      if (window is null)
      {
        return false;
      }

      var windowHandle = new IntPtr(Convert.ToInt64(
        GetRequiredProperty(window, "Hwnd"),
        CultureInfo.InvariantCulture));
      var desktopWindow = NativeMethods.FindWindowEx(windowHandle, IntPtr.Zero, "XLDESK", null);
      var documentWindowHandle = desktopWindow == IntPtr.Zero
        ? IntPtr.Zero
        : NativeMethods.FindWindowEx(desktopWindow, IntPtr.Zero, "EXCEL7", null);
      NativeMethods.GetWindowThreadProcessId(windowHandle, out var processId);
      return windowHandle == identity.ExcelWindowHandle &&
        documentWindowHandle == identity.ExcelDocumentWindowHandle &&
        processId == identity.ProcessId &&
        NativeMethods.GetProp(documentWindowHandle, WorkbookSessionTokenRegistry.WindowPropertyName) ==
          identity.WindowSessionToken;
    }
    finally
    {
      ComRelease.Release(window);
      ComRelease.Release(windows);
    }
  }

  private static void TryRestoreEvents(object application, bool value)
  {
    try
    {
      SetProperty(application, "EnableEvents", value);
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      // Excel may have disconnected after the mutation.
    }
  }

  private static void TryDeleteShape(object? shape)
  {
    if (shape is null)
    {
      return;
    }

    try
    {
      InvokeMethod(shape, "Delete");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      // Compensation is best effort; the caller reports the original failure.
    }
  }

  private static bool ReadBoolean(object target, string propertyName) =>
    Convert.ToBoolean(GetRequiredProperty(target, propertyName), CultureInfo.InvariantCulture);

  private static double ReadFiniteDouble(object target, string propertyName)
  {
    var value = Convert.ToDouble(GetRequiredProperty(target, propertyName), CultureInfo.InvariantCulture);
    return double.IsFinite(value) ? value : throw new InvalidOperationException($"Invalid Shape {propertyName}.");
  }

  private static double ReadPositiveDouble(object target, string propertyName)
  {
    var value = ReadFiniteDouble(target, propertyName);
    return value > 0 ? value : throw new InvalidOperationException($"Invalid Shape {propertyName}.");
  }

  private static object GetRequiredProperty(object target, string propertyName, params object[] arguments) =>
    InvokeProperty(target, propertyName, arguments) ??
    throw new InvalidOperationException($"COM property returned null: {propertyName}");

  private static bool TryGetProperty(object target, string propertyName, out object? value)
  {
    try
    {
      value = InvokeProperty(target, propertyName);
      return true;
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      value = null;
      return false;
    }
  }

  private static object? InvokeProperty(object target, string propertyName, params object[] arguments) =>
    target.GetType().InvokeMember(
      propertyName,
      BindingFlags.GetProperty,
      binder: null,
      target,
      arguments,
      CultureInfo.CurrentCulture);

  private static object? InvokeMethod(object target, string methodName, params object[] arguments) =>
    target.GetType().InvokeMember(
      methodName,
      BindingFlags.InvokeMethod,
      binder: null,
      target,
      arguments,
      CultureInfo.CurrentCulture);

  private static void SetProperty(object target, string propertyName, object value) =>
    _ = target.GetType().InvokeMember(
      propertyName,
      BindingFlags.SetProperty,
      binder: null,
      target,
      [value],
      CultureInfo.CurrentCulture);

  private static bool IsAutomationFailure(Exception exception) =>
    exception is COMException or TargetInvocationException or MissingMemberException or InvalidOperationException;

  private static int GetAutomationHResult(Exception exception) =>
    exception is TargetInvocationException { InnerException: not null } invocationException
      ? invocationException.InnerException!.HResult
      : exception.HResult;

  private readonly record struct Optional<T>(T Value)
  {
    public bool HasValue { get; } = true;
  }

  private static class NativeMethods
  {
    [DllImport("ole32.dll")]
    internal static extern int GetRunningObjectTable(
      int reserved,
      [MarshalAs(UnmanagedType.Interface)] out IRunningObjectTable runningObjectTable);

    [DllImport("ole32.dll")]
    internal static extern int CreateBindCtx(
      int reserved,
      [MarshalAs(UnmanagedType.Interface)] out IBindCtx bindContext);

    [DllImport("user32.dll")]
    internal static extern uint GetWindowThreadProcessId(nint windowHandle, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern nint FindWindowEx(
      nint parentWindow,
      nint childAfter,
      string className,
      string? windowName);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern nint GetProp(nint windowHandle, string propertyName);
  }
}

public sealed record ManagedShapeMetadata(int Version, EvidenceSide Side, CellReference AnchorCell)
{
  public ImageDimensions? SourceDimensions { get; init; }
  public double? AppliedScale { get; init; }
  public const string Prefix = "CraftEvidence:v1|";

  public string Serialize() => $"{Prefix}Side={Side}|Cell=R{AnchorCell.Row}C{AnchorCell.Column}" +
    (SourceDimensions is { } size && AppliedScale is { } scale
      ? FormattableString.Invariant($"|SourceW={size.WidthPoints:R}|SourceH={size.HeightPoints:R}|Scale={scale:R}") : "");

  public static bool IsManagedName(string name) =>
    name.StartsWith("EST_IMG_", StringComparison.Ordinal) &&
    Guid.TryParseExact(name["EST_IMG_".Length..], "N", out _);

  public static bool TryParse(string? value, out ManagedShapeMetadata metadata)
  {
    metadata = null!;
    if (string.IsNullOrWhiteSpace(value) || !value.StartsWith(Prefix, StringComparison.Ordinal))
    {
      return false;
    }

    EvidenceSide? side = null;
    CellReference? anchor = null;
    foreach (var segment in value[Prefix.Length..].Split('|', StringSplitOptions.RemoveEmptyEntries))
    {
      if (segment.StartsWith("Side=", StringComparison.Ordinal) &&
        Enum.TryParse<EvidenceSide>(segment["Side=".Length..], ignoreCase: false, out var parsedSide) &&
        Enum.IsDefined(parsedSide))
      {
        side = parsedSide;
      }
      else if (segment.StartsWith("Cell=R", StringComparison.Ordinal))
      {
        var cell = segment["Cell=R".Length..];
        var columnMarker = cell.IndexOf('C', StringComparison.Ordinal);
        if (columnMarker > 0 &&
          int.TryParse(cell[..columnMarker], NumberStyles.None, CultureInfo.InvariantCulture, out var row) &&
          int.TryParse(cell[(columnMarker + 1)..], NumberStyles.None, CultureInfo.InvariantCulture, out var column) &&
          row is >= 1 and <= ExcelWorksheetLimits.MaximumRow &&
          column is >= 1 and <= ExcelWorksheetLimits.MaximumColumn)
        {
          anchor = new CellReference(row, column);
        }
      }
    }

    if (side is null || anchor is null)
    {
      return false;
    }

    double? Number(string key)
    {
      var field = value.Split('|').FirstOrDefault(part => part.StartsWith(key + "=", StringComparison.Ordinal));
      return field is not null && double.TryParse(field[(key.Length + 1)..], NumberStyles.Float,
        CultureInfo.InvariantCulture, out var number) && double.IsFinite(number) && number > 0 ? number : null;
    }
    metadata = new ManagedShapeMetadata(1, side.Value, anchor.Value)
    {
      SourceDimensions = Number("SourceW") is { } width && Number("SourceH") is { } height
        ? new ImageDimensions(width, height) : null,
      AppliedScale = Number("Scale"),
    };
    return true;
  }
}

public sealed record ManagedShapeTarget(
  string ShapeName,
  string WorksheetName,
  string AlternativeText,
  ManagedShapeMetadata Metadata,
  CellReference TopLeftCell,
  double LeftPoints,
  double TopPoints,
  double WidthPoints,
  double HeightPoints,
  double Rotation = 0,
  double CropLeft = 0,
  double CropTop = 0,
  double CropRight = 0,
  double CropBottom = 0);

public sealed record ManagedShapeSelectionResult(
  bool Succeeded,
  ManagedShapeTarget? Shape,
  string Message)
{
  public static ManagedShapeSelectionResult Failed(string message) => new(false, null, message);
}

public sealed record ManagedShapeMutationResult(
  bool Succeeded,
  bool Changed,
  ManagedShapeTarget? Before,
  ManagedShapeTarget? After,
  string Message)
{
  public bool ReferenceChanged { get; init; }
  public bool CompensationSucceeded { get; init; } = true;
  public static ManagedShapeMutationResult Failed(string message) => new(false, false, null, null, message);
}

public sealed record ManagedShapeExportResult(
  bool Succeeded,
  ManagedShapeTarget? Target,
  string? ImagePath,
  string Message)
{
  public static ManagedShapeExportResult Failed(string message) => new(false, null, null, message);
}
