using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;

namespace VSTextApiInvestigator
{
  /// <summary>
  /// This class implements the tool window exposed by this package and hosts a user control.
  /// </summary>
  /// <remarks>
  /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
  /// usually implemented by the package implementer.
  /// <para>
  /// This class derives from the ToolWindowPane class provided from the MPF in order to use its
  /// implementation of the IVsUIElementPane interface.
  /// </para>
  /// </remarks>
  [Guid("598f970a-5d8d-468e-a265-353b7e763c8f")]
  public class InvestigatorToolWindow : ToolWindowPane
  {
    /// <summary>
    /// Initializes a new instance of the <see cref="InvestigatorToolWindow"/> class.
    /// </summary>
    public InvestigatorToolWindow() : base(null)
    {
      this.Caption = "VSTextApiInvestigator";

      // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
      // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
      // the object returned by the Content property.
      this.Content = new InvestigatorToolWindowControl();
    }
  }
}
