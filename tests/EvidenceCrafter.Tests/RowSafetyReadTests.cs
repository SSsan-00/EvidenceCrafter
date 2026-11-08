using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Runtime.CompilerServices;
using EvidenceCrafter.Excel;
using EvidenceCrafter.Core.Models;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class RowSafetyReadTests
{
  [TestMethod]
  public void RowContentReadFailure_PreventsDeletion()
  {
    var read = typeof(ExcelRowMutationService).GetMethod("HasAnyCellContent", BindingFlags.Static | BindingFlags.NonPublic)!;
    Assert.IsTrue((bool)read.Invoke(null, [new UnreadableRow()])!);
  }

  [TestMethod]
  public void RecoverySnapshot_DisposeKeepsNativeBackupAndMetadata()
  {
    var path = Path.Combine(Path.GetTempPath(), $"row-recovery-test-{Guid.NewGuid():N}.xlsx");
    File.WriteAllText(path, "owned test backup");
    var snapshot = new RowDeletionSnapshot(new WorkbookIdentity("test", "test.xlsx", "test.xlsx", 0, 0, false), "Sheet1", 3, 2, path);
    try
    {
      snapshot.RetainForRecovery(true, "ambiguous Delete");
      snapshot.Dispose();
      Assert.IsFalse(snapshot.IsDisposed);
      Assert.AreEqual(path, snapshot.RecoveryBackupPath);
      Assert.IsTrue(File.Exists(path));
      Assert.IsTrue(File.Exists(path + ".recovery.json"));
      Assert.IsFalse(new ExcelRowMutationService().RestoreDeletedRows(new WorkbookIdentity("test", "test.xlsx", "test.xlsx", 0, 0, false), snapshot).Succeeded);
    }
    finally { File.Delete(path); File.Delete(path + ".recovery.json"); }
  }

  [TestMethod]
  public void FingerprintMatrix_RejectsUnexpectedBulkShape()
  {
    var check = typeof(ExcelRowMutationService).GetMethod("IsFingerprintMatrix", BindingFlags.Static | BindingFlags.NonPublic)!;
    Assert.IsTrue((bool)check.Invoke(null, [new object[2, 3], 2, 3])!);
    Assert.IsFalse((bool)check.Invoke(null, [new object[2], 2, 1])!);
    Assert.IsFalse((bool)check.Invoke(null, [new object[3, 2], 2, 3])!);
    Assert.IsFalse((bool)check.Invoke(null, ["unexpected scalar", 2, 3])!);
    Assert.IsTrue((bool)check.Invoke(null, [false, 1, 1])!);
  }

  [TestMethod]
  public void UniformFingerprintProperty_UsesOnlyKnownValuesAndFallsBackOnFailure()
  {
    var read = typeof(ExcelRowMutationService).GetMethod("ReadUniformFingerprintProperty", BindingFlags.Static | BindingFlags.NonPublic)!;
    var append = typeof(ExcelRowMutationService).GetMethod("AppendFingerprintProperty", BindingFlags.Static | BindingFlags.NonPublic)!;
    foreach (var value in new object?[] { false, true, null, DBNull.Value, "True", 0 })
    {
      var block = new FingerprintProperty(value);
      var uniform = read.MakeGenericMethod(typeof(bool)).Invoke(null, [block, "Value"]);
      Assert.AreEqual(value is bool ? value : null, uniform);
      var cell = new FingerprintProperty(true);
      var builder = new StringBuilder();
      append.Invoke(null, [builder, cell, "Value", false, uniform]);
      Assert.AreEqual(value is false ? "False;" : "True;", builder.ToString());
      Assert.AreEqual(uniform is null ? 1 : 0, cell.Reads);
    }

    var unavailable = new UnreadableRow();
    Assert.IsNull(read.MakeGenericMethod(typeof(bool)).Invoke(null, [unavailable, "Value2"]));
    var failed = new StringBuilder();
    append.Invoke(null, [failed, unavailable, "Value2", false, null]);
    Assert.AreEqual("<unsupported>;", failed.ToString());
  }

  public sealed class FingerprintProperty(object? value)
  {
    public int Reads { get; private set; }
    public object? Value { get { Reads++; return value; } }
  }

  [TestMethod]
  [DataRow(true, true)]
  [DataRow(true, false)]
  [DataRow(false, true)]
  public void ColumnFingerprintFormats_PreserveUnknownsAndReadOnlyMissingProperties(bool readFormat, bool readWrap)
  {
    var read = typeof(ExcelRowMutationService).GetMethod("ReadColumnFingerprintFormats", BindingFlags.Static | BindingFlags.NonPublic)!;
    var block = new ColumnFormatBlock();
    var result = ((object? NumberFormat, object? WrapText)[])read.Invoke(null, [block, 4, readFormat, readWrap])!;
    Assert.AreEqual(readFormat ? "@" : null, result[0].NumberFormat);
    Assert.AreEqual(readWrap ? (object)false : null, result[0].WrapText);
    Assert.AreEqual((null, null), result[1]);
    Assert.AreEqual((null, null), result[2]);
    Assert.IsNull(result[3].NumberFormat);
    Assert.AreEqual(readWrap ? (object)true : null, result[3].WrapText);
    Assert.AreEqual(readFormat ? 1 : 0, block.Uniform.FormatReads);
    Assert.AreEqual(readWrap ? 1 : 0, block.Uniform.WrapReads);
  }

  public sealed class ColumnFormatBlock
  {
    public ColumnFormat Uniform { get; } = new("@", false);
    [IndexerName("Columns")]
    public ColumnFormat this[int index] => index switch
    {
      1 => Uniform,
      2 => throw new COMException("Column unavailable"),
      3 => new(null, DBNull.Value),
      4 => new(null, true, unreadableFormat: true),
      _ => throw new ArgumentOutOfRangeException(nameof(index)),
    };
  }

  public sealed class ColumnFormat(object? format, object? wrap, bool unreadableFormat = false)
  {
    public int FormatReads { get; private set; }
    public int WrapReads { get; private set; }
    public object? NumberFormat
    {
      get { FormatReads++; return unreadableFormat ? throw new COMException("Format unavailable") : format; }
    }
    public object? WrapText { get { WrapReads++; return wrap; } }
  }

  public sealed class UnreadableRow
  {
    public object Value2 => throw new COMException("Read unavailable", unchecked((int)0x80010001));
    public object Formula => throw new COMException("Read unavailable", unchecked((int)0x80010001));
  }

  [TestMethod]
  public void ValidationAbsence_RequiresTheExplicitExcelNoMatchError()
  {
    var read = typeof(ExcelRowMutationService).GetMethod("RangeHasNoValidation", BindingFlags.Static | BindingFlags.NonPublic)!;
    foreach (var (error, expected) in new (Exception?, bool)[]
    {
      (new COMException("No cells were found.", unchecked((int)0x800A03EC)), true),
      (new COMException("該当するセルが見つかりません。", unchecked((int)0x800A03EC)), true),
      (new COMException("Unknown Excel failure", unchecked((int)0x800A03EC)), false),
      (new COMException("No cells were found.", unchecked((int)0x80010001)), false),
      (new COMException("Aucun résultat", unchecked((int)0x800A03EC)), false),
      (new InvalidOperationException("No cells were found."), false),
      (null, false),
    })
      Assert.AreEqual(expected, (bool)read.Invoke(null, [new ValidationRange(error)])!);
    Assert.IsFalse((bool)read.Invoke(null, [new ValidationRange(null, new object())])!);
  }

  public sealed class ValidationRange(Exception? error, object? result = null)
  {
    public object? SpecialCells(int type)
    {
      Assert.AreEqual(-4174, type);
      if (error is not null) throw error;
      return result;
    }
  }
}
