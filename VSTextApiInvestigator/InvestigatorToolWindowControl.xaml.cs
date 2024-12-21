using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.TextManager.Interop;
using System.Diagnostics.CodeAnalysis;
using System.Windows;
using System.Windows.Controls;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.ComponentModelHost;
using System;
using EnvDTE;
using System.Linq;

namespace VSTextApiInvestigator
{
  public class MyTextManagerEvents : IVsTextManagerEvents
  {
    public MyTextManagerEvents(InvestigatorToolWindowControl control)
    {
      mControl = control;
    }

    public void OnRegisterMarkerType(int iMarkerType)
    {
    }

    public void OnRegisterView(IVsTextView pView)
    {
      mControl.OnNewTextView(pView);
    }

    public void OnUnregisterView(IVsTextView pView)
    {
    }

    public void OnUserPreferencesChanged(VIEWPREFERENCES[] pViewPrefs, FRAMEPREFERENCES[] pFramePrefs, LANGPREFERENCES[] pLangPrefs, FONTCOLORPREFERENCES[] pColorPrefs)
    {
    }

    private readonly InvestigatorToolWindowControl mControl;
  }


  /// <summary>
  /// Interaction logic for InvestigatorToolWindowControl.
  /// </summary>
  public partial class InvestigatorToolWindowControl : UserControl
  {
    /// <summary>
    /// Initializes a new instance of the <see cref="InvestigatorToolWindowControl"/> class.
    /// </summary>
    public InvestigatorToolWindowControl()
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      this.InitializeComponent();

      IVsTextManager textManager = ServiceProvider.GlobalProvider?.GetService(typeof(SVsTextManager)) as IVsTextManager;
      if (textManager == null) {
        return;
      }
      var connectionPointContainer = textManager as Microsoft.VisualStudio.OLE.Interop.IConnectionPointContainer;
      if (connectionPointContainer == null) {
        return;
      }

      Microsoft.VisualStudio.OLE.Interop.IConnectionPoint textManagerEvents = null;
      var eventGuid = typeof(IVsTextManagerEvents).GUID;
      connectionPointContainer.FindConnectionPoint(ref eventGuid, out textManagerEvents);
      if (textManagerEvents != null) {
        var myEventReceiver = new MyTextManagerEvents(this);
        textManagerEvents.Advise(myEventReceiver, out uint textManagerCookie);
      }
    }


    public void OnNewTextView(IVsTextView pView)
    {
      IVsUserData userData = pView as IVsUserData;
      if (userData == null) {
        return;
      }
      
      object holder;
      var guidViewHost = Microsoft.VisualStudio.Editor.DefGuidList.guidIWpfTextViewHost;
      userData.GetData(ref guidViewHost, out holder);
      IWpfTextViewHost viewHost = holder as IWpfTextViewHost;
      IWpfTextView textView = viewHost?.TextView;
      ITextSelection textSelection = textView?.Selection;
      if (textSelection != null) {
        textSelection.SelectionChanged += SelectionInTextViewChanged;
      }
    }

    private void OnInvestigateRadioButtonChecked(object sender, RoutedEventArgs e)
    {
      SelectionInTextViewChanged(GetCurrentTextSelection(), EventArgs.Empty);
    }


    private ITextSelection GetCurrentTextSelection()
    {
      IWpfTextViewHost viewHost = GetCurrentViewHost(); 
      IWpfTextView textView = viewHost?.TextView;
      ITextSelection textSelection = textView?.Selection;
      return textSelection;
    }


    private void SelectionInTextViewChanged(object sender, EventArgs e)
    {
      if (mInfoTextBox == null) {
        return;
      }

      ITextSelection textSelection = sender as ITextSelection;
      ITextView textView = textSelection?.TextView;
      ITextBuffer textBuffer = textView?.TextBuffer;
      if (textBuffer == null) {
        mInfoTextBox.Text = "textBuffer == null";
        return;
      }

      var selectedSpans = textSelection.SelectedSpans;
      if (selectedSpans.Count == 0) {
        mInfoTextBox.Text = "Nothing selected";
        return;
      }

      SnapshotSpan firstSelectedSpan = selectedSpans[0];
      try {
        if (mRadioTextStructureNavigator.IsChecked == true) {
          mInfoTextBox.Text = GetInfosFromNavigator(textBuffer, firstSelectedSpan);
        }
        else {
          mInfoTextBox.Text = "NOT IMPLEMENTED";
        }
      }
      catch (Exception ex) {
        mInfoTextBox.Text = $"Exception:\n{ex}";
      }
    }


    private ITextStructureNavigator GetNavigator(ITextBuffer textBuffer)
    {
      IComponentModel mefCompositionContainer = ServiceProvider.GlobalProvider?.GetService(typeof(SComponentModel)) as IComponentModel;
      ITextStructureNavigatorSelectorService navigatorService = mefCompositionContainer?.GetService<ITextStructureNavigatorSelectorService>();
      ITextStructureNavigator navigator = navigatorService?.GetTextStructureNavigator(textBuffer);
      return navigator;
    }


    private string GetInfosFromNavigator(ITextBuffer textBuffer, SnapshotSpan s)
    {
      var navigator = GetNavigator(textBuffer);
      if (navigator == null) {
        return "navigator == null";
      }
      return GetInfosFromNavigator(navigator, s);
    }


    private string GetInfosFromNavigator(ITextStructureNavigator n, SnapshotSpan s)
    {
      return $@"Selected:
  {SpanToStr(s)}

GetSpanOfEnclosing:
  {SpanToStr(n.GetSpanOfEnclosing(s))}

GetSpanOfEnclosing(GetSpanOfEnclosing):
  {SpanToStr(n.GetSpanOfEnclosing(n.GetSpanOfEnclosing(s)))}

GetSpanOfPreviousSibling:
  {SpanToStr(n.GetSpanOfPreviousSibling(s))}

GetSpanOfPreviousSibling(GetSpanOfPreviousSibling):
  {SpanToStr(n.GetSpanOfPreviousSibling(n.GetSpanOfPreviousSibling(s)))}

GetSpanOfPreviousSibling(GetSpanOfEnclosing):
  {SpanToStr(n.GetSpanOfPreviousSibling(n.GetSpanOfEnclosing(s)))}

GetSpanOfNextSibling:
  {SpanToStr(n.GetSpanOfNextSibling(s))}

GetSpanOfNextSibling(GetSpanOfNextSibling):
  {SpanToStr(n.GetSpanOfNextSibling(n.GetSpanOfNextSibling(s)))}

GetSpanOfFirstChild:
  {SpanToStr(n.GetSpanOfFirstChild(s))}

GetSpanOfFirstChild(GetSpanOfFirstChild):
  {SpanToStr(n.GetSpanOfFirstChild(n.GetSpanOfFirstChild(s)))}

";
    }


    private string MakeSpacesVisible(string str)
    {
      return str.Replace(' ', '•').Replace('\t', '→').Replace('\n', '⤶').Replace('\r', '↻');
    }


    private string SpanToStr(SnapshotSpan span)
    {
      return $"[{span.Start.Position},{span.End.Position}): " + MakeSpacesVisible(span.GetText());
    }


    // https://stackoverflow.com/a/6823111/3740047
    private IWpfTextViewHost GetCurrentViewHost()
    {
      // code to get access to the editor's currently selected text cribbed from
      // http://msdn.microsoft.com/en-us/library/dd884850.aspx
      IVsTextManager txtMgr = (IVsTextManager)ServiceProvider.GlobalProvider?.GetService(typeof(SVsTextManager));
      IVsTextView vTextView = null;
      int mustHaveFocus = 1;
      txtMgr?.GetActiveView(mustHaveFocus, null, out vTextView);
      IVsUserData userData = vTextView as IVsUserData;
      if (userData == null) {
        return null;
      }
      else {
        IWpfTextViewHost viewHost;
        object holder;
        var guidViewHost = Microsoft.VisualStudio.Editor.DefGuidList.guidIWpfTextViewHost;
        userData.GetData(ref guidViewHost, out holder);
        viewHost = (IWpfTextViewHost)holder;
        return viewHost;
      }
    }


    /// Given an IWpfTextViewHost representing the currently selected editor pane,
    /// return the ITextDocument for that view. That's useful for learning things 
    /// like the filename of the document, its creation date, and so on.
    private ITextDocument GetTextDocumentForView(IWpfTextViewHost viewHost)
    {
      ITextDocument document;
      viewHost.TextView.TextDataModel.DocumentBuffer.Properties.TryGetProperty(typeof(ITextDocument), out document);
      return document;
    }


    /// Get the current editor selection
    private ITextSelection GetSelection(IWpfTextViewHost viewHost)
    {
      return viewHost.TextView.Selection;
    }

  }
}