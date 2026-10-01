using System.Reflection;
using System.Runtime.InteropServices;
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

  public sealed class UnreadableRow
  {
    public object Value2 => throw new COMException("Read unavailable", unchecked((int)0x80010001));
    public object Formula => throw new COMException("Read unavailable", unchecked((int)0x80010001));
  }
}
