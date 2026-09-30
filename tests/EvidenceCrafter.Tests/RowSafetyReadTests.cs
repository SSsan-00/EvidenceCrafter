using System.Reflection;
using System.Runtime.InteropServices;
using EvidenceCrafter.Excel;

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

  public sealed class UnreadableRow
  {
    public object Value2 => throw new COMException("Read unavailable", unchecked((int)0x80010001));
    public object Formula => throw new COMException("Read unavailable", unchecked((int)0x80010001));
  }
}
