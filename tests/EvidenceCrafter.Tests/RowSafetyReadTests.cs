using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
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

  public sealed class UnreadableRow
  {
    public object Value2 => throw new COMException("Read unavailable", unchecked((int)0x80010001));
    public object Formula => throw new COMException("Read unavailable", unchecked((int)0x80010001));
  }
}
