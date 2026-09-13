using EvidenceCrafter.App;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class SingleInstanceTests
{
  [TestMethod]
  public void SecondLaunch_RestoresMinimizedWindow_AndKeepsOwnedModalActive()
  {
    Exception? failure = null;
    var thread = new Thread(() =>
    {
      try
      {
        var name = @"Local\EvidenceCrafter.Tests." + Guid.NewGuid().ToString("N");
        using var primary = new SingleInstance(name);
        using var main = new Form { Text = "Single-instance activation test" };
        using var modal = new Form { Text = "Owned preview/editor test" };
        using var timeout = new System.Windows.Forms.Timer { Interval = 10000 };
        timeout.Tick += (_, _) =>
        {
          failure ??= new TimeoutException("The activation UI test did not finish.");
          modal.Close();
          main.Close();
        };

        Task SecondLaunch() => Task.Run(() =>
        {
          using var secondary = new SingleInstance(name);
          Ensure(!secondary.IsPrimary, "Secondary launch must not become primary.");
          secondary.ActivateExistingInstance();
        });

        modal.Shown += async (_, _) =>
        {
          try
          {
            modal.Activate();
            Ensure(modal.Handle == GetActiveWindow(), "Modal window must be active.");
            Ensure(!IsWindowEnabled(main.Handle), "ShowDialog must disable its owner.");
            await SecondLaunch();
            await Task.Delay(350); // Allow the primary's 100 ms activation timer to consume the request.
            Ensure(!primary.TakeActivationRequest(), "The request must be handled during the modal loop.");
            Ensure(modal.Handle == GetActiveWindow(), "Activation must preserve the modal window.");
            Ensure(!IsWindowEnabled(main.Handle), "Modal owner must remain disabled.");
          }
          catch (Exception exception) { failure = exception; }
          finally { modal.Close(); }
        };

        main.Shown += async (_, _) =>
        {
          try
          {
            main.WindowState = FormWindowState.Minimized;
            Ensure(main.WindowState == FormWindowState.Minimized, "Main form must minimize.");
            await SecondLaunch();
            await Task.Delay(350);
            Ensure(main.WindowState == FormWindowState.Normal, "Second launch must restore the main form.");
            modal.ShowDialog(main);
          }
          catch (Exception exception) { failure = exception; }
          finally { main.Close(); }
        };
        timeout.Start();
        primary.Run(main);
      }
      catch (Exception exception) { failure = exception; }
    }) { IsBackground = true };
    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(15)), "The STA UI thread must terminate.");
    if (failure is not null) ExceptionDispatchInfo.Capture(failure).Throw();
  }

  [DllImport("user32.dll")]
  private static extern nint GetActiveWindow();

  [DllImport("user32.dll")]
  private static extern bool IsWindowEnabled(nint window);

  private static void Ensure(bool condition, string message)
  {
    if (!condition) throw new InvalidOperationException(message);
  }

  [TestMethod]
  public void SecondLaunch_SignalsBeforeWindowExists_AndCanOwnAfterShutdown()
  {
    var name = @"Local\EvidenceCrafter.Tests." + Guid.NewGuid().ToString("N");
    using (var primary = new SingleInstance(name))
    {
      Assert.IsTrue(primary.IsPrimary);
      Exception? secondaryFailure = null;
      var secondaryThread = new Thread(() =>
      {
        try
        {
          using var secondary = new SingleInstance(name);
          Ensure(!secondary.IsPrimary, "Secondary launch must not become primary.");
          secondary.ActivateExistingInstance();
        }
        catch (Exception exception) { secondaryFailure = exception; }
      }) { IsBackground = true };
      secondaryThread.Start();
      Ensure(secondaryThread.Join(TimeSpan.FromSeconds(3)), "Secondary launch must finish.");
      if (secondaryFailure is not null) ExceptionDispatchInfo.Capture(secondaryFailure).Throw();
      Assert.IsTrue(primary.TakeActivationRequest(), "Startup must not lose activation requests.");
      Assert.IsFalse(primary.TakeActivationRequest(), "One activation is consumed once.");
    }
    using var next = new SingleInstance(name);
    Assert.IsTrue(next.IsPrimary);
  }

  [TestMethod]
  public void AbandonedOwner_AllowsNextLaunch()
  {
    var name = @"Local\EvidenceCrafter.Tests." + Guid.NewGuid().ToString("N");
    Mutex? abandoned = null;
    var thread = new Thread(() =>
    {
      abandoned = new Mutex(true, name + ".Mutex");
      // Exiting without releasing simulates a crashed owner; keep the handle alive.
    });
    thread.Start();
    thread.Join();
    try
    {
      using var next = new SingleInstance(name);
      Assert.IsTrue(next.IsPrimary);
    }
    finally { abandoned?.Dispose(); }
  }
}
