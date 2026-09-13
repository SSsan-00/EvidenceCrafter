namespace EvidenceCrafter.App;

static class Program
{
  [STAThread]
  static void Main()
  {
    using var instance = new SingleInstance();
    if (!instance.IsPrimary)
    {
      instance.ActivateExistingInstance();
      return;
    }

    ApplicationConfiguration.Initialize();
    using var mainForm = new MainForm(new Excel.ExcelSessionCatalog());
    instance.Run(mainForm);
  }
}
