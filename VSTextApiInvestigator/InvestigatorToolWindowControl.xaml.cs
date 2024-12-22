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
using EnvDTE80;
using System.Diagnostics;
using Microsoft.VisualStudio.Editor;
using System.Collections.Generic;
using Microsoft.VisualStudio.VCCodeModel;

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
      ThreadHelper.ThrowIfNotOnUIThread();
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
      ThreadHelper.ThrowIfNotOnUIThread();
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
          mInfoTextBox.Text = GetInfosFromNavigator(firstSelectedSpan);
        }
        else if (mRadioCodeModel.IsChecked == true) {
          mInfoTextBox.Text = GetInfosFromCodeModelAtSpanStart(firstSelectedSpan);
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
      ITextStructureNavigator navigator = NavigatorService?.GetTextStructureNavigator(textBuffer);
      return navigator;
    }


    private string GetInfosFromNavigator(SnapshotSpan selection)
    {
      var navigator = GetNavigator(selection.Snapshot.TextBuffer);
      if (navigator == null) {
        return "navigator == null";
      }
      return GetInfosFromNavigator(navigator, selection);
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


    private string GetInfosFromCodeModelAtSpanStart(SnapshotSpan selection)
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      EditPoint editPoint = GetEditPointFromSnapshotPoint(selection.Start);
      return GetInfosFromCodeModelAtEditPoint(editPoint);
    }


    private EditPoint GetEditPointFromSnapshotPoint(SnapshotPoint point)
    {
      ThreadHelper.ThrowIfNotOnUIThread();

      var adapterService = AdapterService;
      if (adapterService == null) {
        throw new Exception("IVsEditorAdaptersFactoryService is null.");
      }

      ITextBuffer textBuffer = point.Snapshot.TextBuffer;
      var mapper = new VisualStudioNewToOldTextBufferMapper(adapterService, textBuffer);
      if (mapper.VsTextLines == null) {
        throw new Exception("VsTextLines is null.");
      }

      ITextSnapshotLine lineInfo = point.GetContainingLine();
      var offsetInLine0Based = point.Position - lineInfo.Start;
      var offsetInLine1Based = offsetInLine0Based + 1;
      var lineNumber0Based = lineInfo.LineNumber;
      var lineNumber1Based = lineNumber0Based + 1;

      mapper.VsTextLines.CreateEditPoint(lineNumber0Based, offsetInLine0Based, out object pointObj);
      EditPoint pt = pointObj as EditPoint;
      if (pt == null) {
        throw new Exception("Failed to get EditPoint.");
      }
      if (pt.Line != lineNumber1Based || pt.LineCharOffset != offsetInLine1Based) {
        throw new Exception("Wrong EditPoint.");
      }
      return pt;
    }


    private string GetInfosFromCodeModelAtEditPoint(EditPoint pt)
    {
      ThreadHelper.ThrowIfNotOnUIThread();

      string s = $"Start at point: Line={pt.Line}, LineCharOffset={pt.LineCharOffset}, AbsoluteCharOffset={pt.AbsoluteCharOffset}\n\n";
      
      var kindsAlreadyFound = new HashSet<vsCMElement>();
      var availableKinds = Enum.GetValues(typeof(vsCMElement));
      var invalidKinds = new List<vsCMElement>();
      foreach (vsCMElement kind in availableKinds) {
        CodeElement elem = null;
        try {
          elem = pt.CodeElement[kind];
        }
        catch (Exception) {
          invalidKinds.Add(kind);
        }

        if (elem != null && !kindsAlreadyFound.Contains(elem.Kind)) {
          kindsAlreadyFound.Add(elem.Kind);
          s += $"CodeElement[{kind}]: " + GetInfosForCodeElement(elem) + '\n';
        }
      }

      if (invalidKinds.Count > 0) {
        s += "Invalid kinds: " + string.Join(", ", invalidKinds);
      }
      return s;
    }


    private string GetInfosForCodeElement(CodeElement elem)
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      List<object> infos = GetInfosForCodeElementAsNestedLists(elem);
      string s = "";
      foreach (object info in infos) {
        s += ConcatCodeElementInfos(info, "");
      }
      return s;
    }


    private string ConcatCodeElementInfos(object info, string prefix)
    {
      if (info is string str) {
        return prefix + str + '\n';
      }
      else if (info is List<object> list) {
        string s = "";
        foreach (object innerInfo in list) {
          string fullPrefix = prefix + "   ";
          s += ConcatCodeElementInfos(innerInfo, fullPrefix);
        }
        return s;
      }
      else {
        throw new Exception("Unknown info type.");
      }
    }


    private List<object> GetInfosForCodeElementAsNestedLists(CodeElement elem)
    {
      ThreadHelper.ThrowIfNotOnUIThread();

      var outer = new List<object>();

      string name;
      try {
        name = elem.Name;
      }
      catch {
        name = "";
      }
      outer.Add($"Name=\"{name}\", Kind=\"{elem.Kind}\", IsCodeType={elem.IsCodeType}");

      var inner = new List<object>();
      outer.Add(inner);

      try {
        inner.Add($"FullName: \"{MakeSpacesVisible(elem.FullName)}\"");
      }
      catch { }

      try {
        string elemText = MakeSpacesVisible(elem.StartPoint.CreateEditPoint().GetText(elem.EndPoint));
        inner.Add($"Text: {elemText}");
      }
      catch { }

      try {
        inner.Add($"StartPoint: Line={elem.StartPoint.Line}, LineCharOffset={elem.StartPoint.LineCharOffset}, AbsoluteCharOffset={elem.StartPoint.AbsoluteCharOffset}");
        inner.Add($"EndPoint: Line={elem.EndPoint.Line}, LineCharOffset={elem.EndPoint.LineCharOffset}, AbsoluteCharOffset={elem.EndPoint.AbsoluteCharOffset}");
      }
      catch { }

      try {
        inner.Add($"Language: {elem.Language}");
      }
      catch { }

      try {
        inner.Add($"InfoLocation: {elem.InfoLocation}");
      }
      catch { }

      var kindInfos = new List<object>();
      switch (elem.Kind) {
        case vsCMElement.vsCMElementFunction:
          var func = elem as CodeFunction;
          if (func != null) {
            kindInfos.Add("Function infos:");
            var funcInfos = new List<object>();
            kindInfos.Add(funcInfos);
            funcInfos.Add($"FunctionKind: {func.FunctionKind}");
            funcInfos.Add($"InfoLocation: {func.InfoLocation}");
            funcInfos.Add($"Comment: {MakeSpacesVisible(func.Comment)}");
            funcInfos.Add($"DocComment: {MakeSpacesVisible(func.DocComment)}");
            funcInfos.Add($"Parameters: {func.Parameters.Count}");
            foreach (CodeElement param in func.Parameters) {
              funcInfos.Add(GetInfosForCodeElementAsNestedLists(param));
            }

            var vcFunc = elem as VCCodeFunction;
            if (vcFunc != null) {
              funcInfos.Add($"C++ specific:");
              var vcInfos = new List<object>();
              funcInfos.Add(vcInfos);
              vcInfos.Add($"DisplayName: {MakeSpacesVisible(vcFunc.DisplayName)}");
              vcInfos.Add($"DeclarationText: {MakeSpacesVisible(vcFunc.DeclarationText)}");
              vcInfos.Add($"IsTemplate: {vcFunc.IsTemplate}");
              vcInfos.Add($"BodyText: {MakeSpacesVisible(vcFunc.BodyText)}");
              vcInfos.Add($"TemplateParameters: {vcFunc.TemplateParameters.Count}");
              foreach (CodeElement param in vcFunc.TemplateParameters) {
                vcInfos.Add(GetInfosForCodeElementAsNestedLists(param));
              }
            }
          }
          break;
      }

      inner.AddRange(kindInfos);

      if (elem.Children.Count > 0) {
        inner.Add($"Children:");
        foreach (CodeElement child in elem.Children) {
          inner.Add(GetInfosForCodeElementAsNestedLists(child));
        }
      }
      else {
        inner.Add($"No children.");
      }

      return outer;
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


    private IComponentModel MefCompositionContainer => ServiceProvider.GlobalProvider?.GetService(typeof(SComponentModel)) as IComponentModel;
    private ITextStructureNavigatorSelectorService NavigatorService => MefCompositionContainer?.GetService<ITextStructureNavigatorSelectorService>();
    private IVsEditorAdaptersFactoryService AdapterService => MefCompositionContainer?.GetService<IVsEditorAdaptersFactoryService>();
  }



  //==============================================================================
  // VisualStudioNewToOldTextBufferMapper
  //==============================================================================

  /// <summary>
  /// As far as I understand, Microsoft.VisualStudio.Text.ITextBuffer and similar classes are the "new" .NET managed classes.
  /// On the other hand, the stuff in the EnvDTE namespace (e.g. EnvDTE.Document and EnvDTE.TextDocument) represent 'old' classes,
  /// predating the .NET implementations. They are always COM interfaces. However, they are still relevant for certain things,
  /// e.g. the FileCodeModel. Things like IVsTextBuffer seem to be wrappers/adapters around the old classes. We can get from a 
  /// "new world" object (such as ITextBuffer) to the adapter via the IVsEditorAdaptersFactoryService service, resulting in e.g. 
  /// a IVsTextBuffer. (I think that service is just getting some object from ITextBuffer.Properties.) Digging through decompiled 
  /// VS .NET code, from the adapter we get to the "old world" object via IExtensibleObject. Note that the documentation of 
  /// IExtensibleObject states that it is Microsoft internal. We ignore this warning here. The only valid arguments to
  /// IExtensibleObject.GetAutomationObject() seem to be "Document" (giving an EnvDTE.Document) and "TextDocument" (giving
  /// an EnvDTE.TextDocument).
  /// </summary>
  struct VisualStudioNewToOldTextBufferMapper
  {
    public IVsTextBuffer VsTextBuffer { get; private set; }
    public IVsTextLines VsTextLines { get; private set; }
    public IExtensibleObject ExtensibleObject { get; private set; }

    public EnvDTE.Document Document {
      get {
        ThreadHelper.ThrowIfNotOnUIThread();
        object docObj = null;
        ExtensibleObject?.GetAutomationObject("Document", null, out docObj);
        return docObj as EnvDTE.Document;
      }
    }

    public EnvDTE.TextDocument TextDocument {
      get {
        ThreadHelper.ThrowIfNotOnUIThread();
        object docObj = null;
        ExtensibleObject?.GetAutomationObject("TextDocument", null, out docObj);
        return docObj as EnvDTE.TextDocument;
      }
    }

    public VisualStudioNewToOldTextBufferMapper(IVsEditorAdaptersFactoryService adapterService, ITextBuffer textBuffer)
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      VsTextBuffer = adapterService?.GetBufferAdapter(textBuffer);
      VsTextLines = VsTextBuffer as IVsTextLines;
      ExtensibleObject = VsTextBuffer as IExtensibleObject;
    }
  }

}