using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Operations;
using System;
using System.ComponentModel.Composition;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace VSTextApiInvestigator
{
  [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
  [Guid(VSTextApiInvestigatorPackage.PackageGuidString)]
  [ProvideMenuResource("Menus.ctmenu", 1)]
  [ProvideToolWindow(typeof(InvestigatorToolWindow))]
  public sealed class VSTextApiInvestigatorPackage : AsyncPackage
  {
    public const string PackageGuidString = "120da8b9-77a4-48f8-b379-87d428a667ff";

    protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
    {
      // When initialized asynchronously, the current thread may be a background thread at this point.
      // Do any initialization that requires the UI thread after switching to the UI thread.
      await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
        await InvestigatorToolWindowCommand.InitializeAsync(this);
    }
  }
}
