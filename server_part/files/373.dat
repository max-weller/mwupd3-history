
''es_temp2.vb

'NEU...
#Reference System.Windows.Forms.dll
#Reference System.Data.dll
#Reference System.Drawing.dll
#Reference WeifenLuo.WinFormsUI.Docking.dll
#Reference TenTec.Windows.iGridLib.iGrid.dll

#Reference ICSharpCode.NRefactory.dll
#Reference ICSharpCode.SharpDevelop.Dom.dll

#Imports System.Windows.Forms
#Imports System.Data
#Imports System.Text
#Imports System.Drawing
#Imports ScriptIDE.Core
#Imports ScriptIDE.ScriptHost
#Imports ScriptIDE.ScriptWindowHelper
#Imports Tentec.Windows.iGridLib

#Imports System.Collections.Generic
#Imports System.IO
#Imports System.Threading

#Imports ICSharpCode
#Imports ICSharpCode.SharpDevelop
#Imports ICSharpCode.SharpDevelop.Dom
#Imports ICSharpCode.SharpDevelop.Dom.NRefactoryResolver

#runtime exe

'-->
'--> G L O B 
#Para DebugMode extern
'#Para SilentMode true



'--> P A R A M E T E R -------------------------
Const  winID = "__CLASSNAME__.ccTester"
const  globHideOnClose as boolean =  false           'bei true merkt er sich die Position
public winCaption as string="Code-Completion Test"
public glob_defaultPath as string = "C:\yPara\scriptIDE4\scriptPara\ccTest\"
public glob_scriptPath as string = "C:\yPara\scriptIDE4\scriptClass\"

dim toolBarColor as string="#4B4C8B"


public WithEvents FormRef As Form
public PanelRef As ScriptedPanel

public shared toolBar As ScriptedPanel
public shared statBar As ScriptedPanel
public shared panelMain As ScriptedPanel

dim myTextArea as textbox
dim Withevents myListBox as ListBox

Public Const Auto = -2
shared myScriptClass


'-->
sub test1()
  'trace ("test1()[es_temp2....xxxxxxxxxxxxxx.....]")
  monitorAdd ("test1()[.]")
  'exit sub
  monitorAdd ("test1()[..]")
  monitorAdd ("test1()[...]")
exit sub

end sub




sub test2()
  'msgBox("OK - 2")
  
  'demoCodeInsert()
end sub



sub test3()
  'msgBox("OK - 3")
end sub







'-->
'--> FORM-HELPER/...Factory ________________________


:function getOuterWindowRef(panelRef as object) as object
  on error resume next
  if panelRef.form.parent.parent is Nothing then : err.number=0
    '...falls es kein DockPanel-Fenster ist
    :return panelRef.form
  else : err.number=0
    '...DockPanel-Fenster haben noch weitere Schichten dazwischen:
    :return panelRef.form.parent.parent
  end if
End function



Sub globAddHandler()
   AddHandler  myTextArea.KeyDown, AddressOf onMyTextArea_KeyDown
   AddHandler  myTextArea.KeyPress, AddressOf onMyTextArea_KeyPress
  'AddHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  'AddHandler Timer1.Tick, AddressOf Timer1_Tick
  'AddHandler formRef.Resize, AddressOf Form_Resize
  'AddHandler TT.TraceWrite, AddressOf OnTrace
  'AddHandler TT.PrintLineChanged, AddressOf OnPrintLine
  'AddHandler  myTextArea.KeyDown, AddressOf OnPrintLine
end sub


Sub globRemoveHandler()
  trace "globRemoveHandler","try..."
  'if formRef is nothing then exit sub
  RemoveHandler  myTextArea.KeyDown, AddressOf onMyTextArea_KeyDown
   RemoveHandler  myTextArea.KeyPress, AddressOf onMyTextArea_KeyPress
  'RemoveHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  'RemoveHandler Timer1.Tick, AddressOf Timer1_Tick
  'RemoveHandler formRef.Resize, AddressOf Form_Resize
  'RemoveHandler TT.TraceWrite, AddressOf OnTrace
  'RemoveHandler TT.PrintLineChanged, AddressOf OnPrintLine
   trace "globRemoveHandler","DONE"
end sub


sub onInitialize()
  globAddHandler
end sub


sub onTerminate()
  globRemoveHandler
  'stopTimer()
end sub


Function GetPanelRef()
  'onTerminate() 'aufruf, um das alte leben ein bischen anzuhalten 
  ' If PanelRef IsNot Nothing Then Exit Sub

  '--> ... --- dockWindow / normales Fenster
  'VERALTET: PanelRef = ZZ.IDEHelper.CreateDockWindow(winID, 1) '1=popupWin
  PanelRef = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)'  : err.number=0
  'PanelRef = ZZ.CreateWindow(winID, DWndFlags.StdWindow, 1)'  : err.number=0

  'formRef = getOuterWindowRef(panelRef)
  formRef = panelRef.form
  formRef.text=winCaption
  : CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = globHideOnClose : err.number=0
  return panelRef
End Function



'-->
'--> A U T O S T A R T ___________________________

Sub AutoStart()
  IO.Directory.createDirectory(glob_defaultPath)
  
  monitorClear()
  monitorAdd("AutoStart(): mw_ideTest")
  GetPanelRef()
  dim el as object
  
  '--> ... containerMain
  With PanelRef
    .resetControls ()
    .activateEvents = "|IconMouseDown|   |TextboxKeyDown|"
    panelMain  =.addPanel("containerMain"   , DockStyle.Fill)
    toolBar     =.addPanel("toolBar"        , DockStyle.Top, intHeight:=33)
    statBar     =.addPanel("statBar"        , DockStyle.Bottom, intHeight:=33)
  end with
  
  '--> ... --- toolBar
  with toolBar
    .resetControls  (10,3)
    .visible=true
    .height=30
    .BackColor = ColorTranslator.FromHtml(toolBarColor)
    .addButton ("cmdExit"                , "E N D E"       , "#F0C0C0")
    '.BackColor = ColorTranslator.FromHtml("#f88")
  end with
    
  with panelMain 
    .resetControls  (0,0)
    .BackColor = ColorTranslator.FromHtml(toolBarColor)
    .BackColor = ColorTranslator.FromHtml("#E3E0DB")
  end with
  
  '--> ... --- statBar
  with statbar
    .resetControls  (10,5)
    .visible=true
    .height=30
    .BackColor = ColorTranslator.FromHtml(toolBarColor)
    '.BackColor = ColorTranslator.FromHtml("#88f")
    statBar.height=30
    .addButton ("cmdDoParserStep"        , "Parse"          , "#AD9C01")
    .addButton ("cmdGetExp"              , "get Expression"       , "#C0C0C0")

  end with
   
  '--> ... --- panelMain
  With panelMain
    myListBox = New ListBox()
    myListBox.Size = New size(140,220)
    .Controls.add(myListBox)
    
    .addLabel("myInfoLabel", "infoLabel", "#fea", "#000", , , 300, 60)
    
    '.insertX = 5 :.insertY = 0' 110
    .addTextbox ("myTextArea" ,  -2         , "inBox"    , 0  , "#FFFF99", , , "multiline=240")
     el=panelRef.element("txt_myTextArea")
     myTextArea=el
     el.Dock = DockStyle.Fill
     el.WordWrap=false
     el.scrollbars = System.Windows.Forms.ScrollBars.Both
     el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
     el.bringToFront()
     
     dim miniHelp as string="" 
     miniHelp=miniHelp+vbNewLine+"'...hier Code eingeben und dann contr.ENTER drücken"
     miniHelp=miniHelp+vbNewLine+"'--------------------------------------------------"
     miniHelp=miniHelp+vbNewLine+"Imports System.Drawing"
     miniHelp=miniHelp+vbNewLine+""
     miniHelp=miniHelp+vbNewLine+"Public Class TestSC"
     miniHelp=miniHelp+vbNewLine+"'z.B."
     miniHelp=miniHelp+vbNewLine+"Public Shared Sub HalloWelt()"
     miniHelp=miniHelp+vbNewLine+"  msgBox(""hello World"")"
     miniHelp=miniHelp+vbNewLine+"  "
     miniHelp=miniHelp+vbNewLine+"  '...happy coding ;-)"
     miniHelp=miniHelp+vbNewLine+"  "
     miniHelp=miniHelp+vbNewLine+"End Sub"
     miniHelp=miniHelp+vbNewLine+"End Class"
     miniHelp=miniHelp+vbNewLine+""
     miniHelp=miniHelp+vbNewLine+"Public Class Test222"
     miniHelp=miniHelp+vbNewLine+"Private _color As Color"
     miniHelp=miniHelp+vbNewLine+"Private Sub new(myColor As String)"
     miniHelp=miniHelp+vbNewLine+"  _color = ColorTranslator.FromHtml(myColor)"
     miniHelp=miniHelp+vbNewLine+"End Sub"
     miniHelp=miniHelp+vbNewLine+""
     miniHelp=miniHelp+vbNewLine+"Public ReadOnly Property ColorValue() As Color"
     miniHelp=miniHelp+vbNewLine+"  Get"
     miniHelp=miniHelp+vbNewLine+"    Return _color"
     miniHelp=miniHelp+vbNewLine+"  End Get"
     miniHelp=miniHelp+vbNewLine+"End Property"
     miniHelp=miniHelp+vbNewLine+""
     miniHelp=miniHelp+vbNewLine+"Public Shared Function GetClass(theColor As String) As Test222"
     miniHelp=miniHelp+vbNewLine+"  Return New Test222(theColor)"
     miniHelp=miniHelp+vbNewLine+"End Function"
     miniHelp=miniHelp+vbNewLine+""
     miniHelp=miniHelp+vbNewLine+"End Class"
     miniHelp=miniHelp+vbNewLine+""
     myTextArea.text=miniHelp
  end with

  globAddHandler()

  myTextArea.SelectionStart=myTextArea.text.length
  myTextArea.focus()

end sub


'-->
'--> E V E N T S ___________________________

dim skipNextKeyPress as boolean=false


Sub onButtonEvent(e)
  setMonitorRef()
  clearAll
  statustext("OK")
  errorText("")
  printLine 11, "e.Sender.Name" , e.Sender.Name
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  MonitorAdd("======================== onLabelEvent")
  MonitorAdd("SenderFullName: " + e.Sender.Name)
  MonitorAdd("___MouseButton: " + e.MouseButton)
  MonitorAdd("________Action: " + action)
  
  '' if action.contains("cmdConnectServer") then connectToServer()
  '' if action.contains("cmdSendTestMail") then sendTestMail()
 
  callCmdByName(e)

end Sub


Sub cmdExit(e)
  FormRef.Close
End Sub


Friend pcRegistry As Dom.ProjectContentRegistry
Friend myProjectContent As Dom.DefaultProjectContent
Friend parseInformation As New Dom.ParseInformation()
Private lastCompilationUnit As Dom.ICompilationUnit
Shared ReadOnly CurrentLanguageProperties As Dom.LanguageProperties = Dom.LanguageProperties.VBNet
'If(IsVisualBasic, Dom.LanguageProperties.VBNet, Dom.LanguageProperties.CSharp)



Sub cmdDoParserStep(e)  
  pcRegistry = New Dom.ProjectContentRegistry()
  ' Default .NET 2.0 registry
  ' Persistence lets SharpDevelop.Dom create a cache file on disk so that
  ' future starts are faster.
  ' It also caches XML documentation files in an on-disk hash table, thus
  ' reducing memory usage.
  pcRegistry.ActivatePersistence(Path.Combine(glob_defaultPath, "CSharpCodeCompletion"))
  myProjectContent = New Dom.DefaultProjectContent()
  myProjectContent.Language = CurrentLanguageProperties
  
  myProjectContent.AddReferencedContent(pcRegistry.Mscorlib)
  
  Dim referencedAssemblies As String() = {"System", "System.Data", "System.Drawing", "System.Xml", "System.Windows.Forms", "Microsoft.VisualBasic"}
  For Each assemblyName As String In referencedAssemblies
    Dim assemblyNameCopy As String = assemblyName
    TRACE "Loading " & assemblyNameCopy & "..."
    Dim referenceProjectContent As Dom.IProjectContent = pcRegistry.GetProjectContentForReference(assemblyName, assemblyName)
    myProjectContent.AddReferencedContent(referenceProjectContent)
    If TypeOf referenceProjectContent Is Dom.ReflectionProjectContent Then
      TryCast(referenceProjectContent, Dom.ReflectionProjectContent).InitializeReferences()
    End If
  Next
  'If IsVisualBasic Then
  myProjectContent.DefaultImports = New Dom.DefaultUsing(myProjectContent)
  myProjectContent.DefaultImports.Usings.Add("System")
  myProjectContent.DefaultImports.Usings.Add("System.Text")
  myProjectContent.DefaultImports.Usings.Add("Microsoft.VisualBasic")
  'End If
  
  ParseStep()
  
  TRACE "Parse Complete"
  If lastCompilationUnit Is Nothing Then
    trace "ERR: lcu is nothing"
  end if
  TRACE "Classes Count", lastCompilationUnit.Classes.Count
  For Each cls As Dom.IClass In lastCompilationUnit.Classes
    TRACE "Class", cls.DefaultReturnType.FullyQualifiedName
  Next
End Sub


Private Sub ParseStep()

  Dim code As String = Nothing
  '' Invoke(New MethodInvoker(Function() Do
  ''     code = textEditorControl1.Text
  '' End Function))
  code = PanelRef("myTextArea").Text
  'Msgbox(code)
  Dim textReader As TextReader = New StringReader(code)
  Dim newCompilationUnit As Dom.ICompilationUnit
  Dim supportedLanguage As NRefactory.SupportedLanguage
  'If IsVisualBasic Then
  supportedLanguage = NRefactory.SupportedLanguage.VBNet
  '' Else
  ''     supportedLanguage = NRefactory.SupportedLanguage.CSharp
  '' End If
  Using p As NRefactory.IParser = NRefactory.ParserFactory.CreateParser(supportedLanguage, textReader)
      ' we only need to parse types and method definitions, no method bodies
      ' so speed up the parser and make it more resistent to syntax
      ' errors in methods
      p.ParseMethodBodies = False
  
      p.Parse()
      newCompilationUnit = ConvertCompilationUnit(p.CompilationUnit)
  End Using
  ' Remove information from lastCompilationUnit and add information from newCompilationUnit.
  myProjectContent.UpdateCompilationUnit(lastCompilationUnit, newCompilationUnit, "myFileName.vb")
  lastCompilationUnit = newCompilationUnit
  parseInformation.SetCompilationUnit(newCompilationUnit)
End Sub

Private Function ConvertCompilationUnit(cu As NRefactory.Ast.CompilationUnit) As Dom.ICompilationUnit
  Dim converter As Dom.NRefactoryResolver.NRefactoryASTConvertVisitor
  converter = New Dom.NRefactoryResolver.NRefactoryASTConvertVisitor(myProjectContent)
  cu.AcceptVisitor(converter, Nothing)
  Return converter.Cu
End Function

''' <summary>
''' Find the expression the cursor is at.
''' Also determines the context (using statement, "new"-expression etc.) the
''' cursor is at.
''' </summary>
Private Function FindExpression(textArea As TextBox) As Dom.ExpressionResult
  Dim finder As Dom.IExpressionFinder
  'If MainForm.IsVisualBasic Then
      finder = New Dom.VBNet.VBExpressionFinder()
  'Else
  '    finder = New Dom.CSharp.CSharpExpressionFinder(mainForm.parseInformation)
  'End If
  Dim expression As Dom.ExpressionResult = finder.FindExpression(textArea.Text, textArea.SelectionStart)
  If expression.Region.IsEmpty Then
     Dim line = textArea.GetLineFromCharIndex( textArea.SelectionStart)
     Dim col  =  textArea.SelectionStart-textArea.GetFirstCharIndexFromLine(line)
     expression.Region = New Dom.DomRegion(Line + 1, Col + 1)
  End If
  Return expression
End Function

Private Function FindFullExpression(textArea As TextBox) As Dom.ExpressionResult
  Dim finder As Dom.IExpressionFinder
  'If MainForm.IsVisualBasic Then
      finder = New Dom.VBNet.VBExpressionFinder()
  'Else
  '    finder = New Dom.CSharp.CSharpExpressionFinder(mainForm.parseInformation)
  'End If
  Dim expression As Dom.ExpressionResult = finder.FindFullExpression(textArea.Text, textArea.SelectionStart-1)
  If expression.Region.IsEmpty Then
     Dim line = textArea.GetLineFromCharIndex( textArea.SelectionStart)
     Dim col  =  textArea.SelectionStart-textArea.GetFirstCharIndexFromLine(line)
     expression.Region = New Dom.DomRegion(Line + 1, Col + 1)
  End If
  Return expression
End Function

Sub cmdGetExp(e)
  dim textBox = PanelRef("myTextArea")
  '' Dim exp = FindExpression(textBox)
  '' TRACE "exp",exp.Expression
  
  GenerateCompletionData("myFileName.vb", textBox, "."c)
End Sub

Public sub GenerateCompletionData(fileName As String, textArea As TextBox, charTyped As Char) 'As Dom.ICompletionData()
  ' We can return code-completion items like this:

  'return new ICompletionData[] {
  '	new DefaultCompletionData("Text", "Description", 1)
  '};
  
  Dim resolver As New ICSharpCode.SharpDevelop.Dom.NRefactoryResolver.NRefactoryResolver(myProjectContent.Language)
  Dim exp As Dom.ExpressionResult = FindExpression(textArea)
  TRACE "expression",exp.Expression
  Dim rr As Dom.ResolveResult = resolver.Resolve(exp, parseInformation, textArea.Text)
  'Dim resultList As New List(Of ICompletionData)()
  If rr IsNot Nothing Then
    Dim completionData As collections.ArrayList = rr.GetCompletionData(myProjectContent)
    If completionData IsNot Nothing Then
      TRACE "completionData.Count=",completionData.Count
      
      showListForCompletion(completionData)
      'AddCompletionData(resultList, completionData)
    else
      TRACE "completionData IS Nothing"
    End If
    
  else
    TRACE "ResolveResult IS Nothing"
  End If
  'Return resultList.ToArray()
End sub


Sub showListForCompletion(completionData as collections.ArrayList)
  myListBox.hide
  myListBox.sorted=false
  myListBox.Items.clear
  For Each item As Object in completionData
    If typeOf item is string Then
      myListBox.Items.add(item)
    else
      myListBox.Items.add(item.name)
    end if
    
  Next
  dim pos as point = myTextArea.GetPositionFromCharIndex(myTextArea.SelectionStart)
  myListBox.top = pos.y+15
  myListBox.left = pos.x+2
  myListBox.show
  myListBox.sorted=true
  myListBox.bringToFront
End Sub

sub showTooltip()
  'Dim textArea As TextEditor.TextArea = editor.ActiveTextAreaControl.TextArea
  Dim expression As Dom.ExpressionResult = FindFullExpression(myTextArea)
  TRACE "expression",expression.Expression
  
  Dim resolver As New NRefactoryResolver(myProjectContent.Language)
  Dim rr As ResolveResult = resolver.Resolve(expression, parseInformation, myTextArea.Text)
  Dim text as string = GetText(rr)
  
  dim pos as point = myTextArea.GetPositionFromCharIndex(myTextArea.SelectionStart)
  dim lab As Label = PanelRef("myInfoLabel")
  lab.Top = pos.y+18
  lab.left = pos.x-10
  lab.Text = text
  lab.bringToFront
  lab.show
  
  '' Dim toolTipText As String = GetText(rr)
  '' If toolTipText IsNot Nothing Then
  ''     e.ShowToolTip(toolTipText)
  '' End If
End Sub



		Private Shared Function GetText(result As ResolveResult) As String
			If result Is Nothing Then
				Return Nothing
			End If
			If TypeOf result Is MixedResolveResult Then
				Return GetText(DirectCast(result, MixedResolveResult).PrimaryResult)
			End If
			Dim ambience As IAmbience = New VBNet.VBNetAmbience()' If(MainForm.IsVisualBasic, DirectCast(New VBNetAmbience(), IAmbience), New CSharpAmbience())
			ambience.ConversionFlags = ConversionFlags.StandardConversionFlags Or ConversionFlags.ShowAccessibility
			If TypeOf result Is MemberResolveResult Then
				Return GetMemberText(ambience, DirectCast(result, MemberResolveResult).ResolvedMember)
			ElseIf TypeOf result Is LocalResolveResult Then
				Dim rr As LocalResolveResult = DirectCast(result, LocalResolveResult)
				ambience.ConversionFlags = ConversionFlags.UseFullyQualifiedTypeNames Or ConversionFlags.ShowReturnType
				Dim b As New StringBuilder()
				If rr.IsParameter Then
					b.Append("parameter ")
				Else
					b.Append("local variable ")
				End If
				b.Append(ambience.Convert(rr.Field))
				Return b.ToString()
			ElseIf TypeOf result Is NamespaceResolveResult Then
				Return "namespace " + DirectCast(result, NamespaceResolveResult).Name
			ElseIf TypeOf result Is TypeResolveResult Then
				Dim c As IClass = DirectCast(result, TypeResolveResult).ResolvedClass
				If c IsNot Nothing Then
					Return GetMemberText(ambience, c)
				Else
					Return ambience.Convert(result.ResolvedType)
				End If
			ElseIf TypeOf result Is MethodGroupResolveResult Then
				Dim mrr As MethodGroupResolveResult = TryCast(result, MethodGroupResolveResult)
				Dim m As IMethod = mrr.GetMethodIfSingleOverload()
				If m IsNot Nothing Then
					Return GetMemberText(ambience, m)
				Else
					Return ("Overload of " + ambience.Convert(mrr.ContainingType) & ".") + mrr.Name
				End If
			Else
				Return Nothing
			End If
		End Function

		Private Shared Function GetMemberText(ambience As IAmbience, member As IEntity) As String
			Dim text As New StringBuilder()
			If TypeOf member Is IField Then
				text.Append(ambience.Convert(TryCast(member, IField)))
			ElseIf TypeOf member Is IProperty Then
				text.Append(ambience.Convert(TryCast(member, IProperty)))
			ElseIf TypeOf member Is IEvent Then
				text.Append(ambience.Convert(TryCast(member, IEvent)))
			ElseIf TypeOf member Is IMethod Then
				text.Append(ambience.Convert(TryCast(member, IMethod)))
			ElseIf TypeOf member Is IClass Then
				text.Append(ambience.Convert(TryCast(member, IClass)))
			Else
				text.Append("unknown member ")
				text.Append(Ctype(member,object).ToString())
			End If
			Dim documentation As String = member.Documentation
			If documentation IsNot Nothing AndAlso documentation.Length > 0 Then
				text.Append(ControlChars.Lf)
				'HACK
                'text.Append(CodeCompletionData.XmlDocumentationToText(documentation))
			End If
			Return text.ToString()
		End Function
        
        
        
sub txt_myTextArea_KeyDown(e)
  If e.keyString="OEMPERIOD" Then
  
    GenerateCompletionData("myFileName.vb", myTextArea, "."c)
  End If
  If e.keyString="SHIFT-D8" Then
    
    showTooltip()
    
    
  End If
  '' ??? wie kann man contr.ENTER bequem abgreifen, ohne daß er eine in der Textbox einfügt ???

  TRACE "txt_myTextArea_KeyDown", e.keyString
  '' if e.keyString="CTRL-RETURN" then
  ''   executeDirektFenster()
  ''   e.cancel=true
  '' end if
end sub


sub onMyTextArea_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
  'monitorAdd(e.keycode.toString)
  if e.control=true
    if e.keyCode=13 then
       skipNextKeyPress=true
      'monitorAdd("KeyDown-TREFFER")
    end if
  end if
end sub



sub onMyTextArea_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
  'monitorAdd(e.KeyChar)
  if skipNextKeyPress =true then
    skipNextKeyPress=false
    e.handled=true
    monitorAdd("onMyTextArea_KeyPress: CONTROL-ENTER")
    'executeDirektFenster()
  end if
end sub


'' 
'' sub executeDirektFenster()
''   'msgBox("OK - 3")
''   dim template as string=getCodeTemplate()
''   dim insert as string =myTextArea.text
''   dim content as string=replace(template, "[[INSERT-TEST-CODE]]", insert)
'' 
''   dim filespec as string =glob_scriptPath+"es_temp3.vb"
''   IO.Directory.CreateDirectory(glob_defaultPath)
''   zz.filePutContents(filespec, content) ', [append As System.Object = False]) As System.Object
''   
''   dim script=zz.newScriptClass("es_temp3")
''   script.test1()
'' end sub 




'' 
'' '-->
'' '--> ......................... D A T A - B L O C K 
'' function getCodeTemplate()
'' #Data myData as String
'' 
'' '#####################################################
'' 
'' 'NEU-2...
'' #Imports ScriptIDE.Core
'' #Imports ScriptIDE.ScriptHost
'' #Imports ScriptIDE.ScriptWindowHelper
'' #Reference System.Windows.Forms.dll
'' #Imports System.Windows.Forms
'' 
'' #Reference System.Drawing.dll
'' #Imports System.Drawing
'' 
'' #Reference System.Data.dll
'' #Imports System.Data
'' 
'' '' #Reference TenTec.Windows.iGridLib.iGrid.dll
'' '' #Imports Tentec.Windows.iGridLib
'' '' 
'' '' #Reference WeifenLuo.WinFormsUI.Docking.dll
'' '' 
'' 
'' 
'' 
'' '--> G L O B 
'' #Para DebugMode intern
'' #Para SilentMode true
'' 
'' 
'' 
'' sub test1
'' 
''   'msgBox("ich bin die test1 von ES_TEMP3")
''   [[INSERT-TEST-CODE]]
'' end sub
'' '###################################################
'' 
'' #EndData
''   dim RESULT as string=myData 
''   return RESULT
'' end function
'' 

'-->
'--> ........................................LIBs





'-->
'--> outMONITOR

public globMonitorRef as object


 sub clearAll()
   on error resume next
   monitorClear()
   statustext_reset()
   zz.traceClear()
   zz.printLineReset()
   err.number=0
end sub



:sub setMonitorRef()
  on error resume next
  globMonitorRef = zz.scriptClass("es_testLabor")
  trace "globMonitorRef", globMonitorRef.name
  err.number=0
end sub

: function getMonitorRef()
  on error resume next
  if globMonitorRef is nothing then
    globMonitorRef = zz.scriptClass("es_testLabor")
  end if
  return globMonitorRef
end function

: sub monitorClear()
   on error resume next
   globMonitorRef.clear()
   err.number=0
end sub


: sub monitorAdd(para1 as string, optional para2 as string="")
   on error resume next
   globMonitorRef.add(para1+": "+para2+"<--")
   err.number=0
end sub

: sub statustext(para1 as string, optional para2 as string="")
   on error resume next
   'globMonitorRef.statustext(para1+": "+para2+"<--")
   globMonitorRef.statustext(para1+"")
   err.number=0
end sub

: sub statustext_reset()
   on error resume next
   globMonitorRef.statustext_reset()
   err.number=0
end sub


: sub errorText(para as string)
   on error resume next
   globMonitorRef.errorText(para)
   err.number=0
end sub




'--> ...........................EventHelper



sub callCmdByName(e)
  on error resume next
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim funcName as string=name(1)
  trace "funcName", funcName
  dim i as integer=1
  
  ''dim timerStart = GetTime()
  ''dim deltaTime as integer
  
  dim scriptClass= Me
  Dim myType As Type = scriptClass.GetType()
  Dim myMethod As System.Reflection.MethodInfo = myType.GetMethod(funcName)
  if myMethod.toString = "" then
    dim errMes="ERR: Sub '"+funcName+"' nicht gefunden" 
    statustext(errMes)
    dumpUnknownEventHandler(funcName)
    exit sub
  end if
  
  monitorAdd("______procName: "+myMethod.toString)
  monitorAdd("AAA: "+err.number.tostring)
  monitorAdd("BBB: "+err.description)

  dim args(0)' 
  args(0)=e
  myMethod.Invoke(scriptClass, args)
  'CallByName(scriptClass, funcName, Microsoft.VisualBasic.CallType.Method)

  '' deltaTime=GetTime()-timerStart
  '' monitorAdd("--------------- DONE")
  '' monitorAdd("anz schleifen je sek: "+(i / deltaTime * 1000).toString("0000"))
  '' monitorAdd("")
end sub




sub dumpUnknownEventHandler(funcName)
  dim tpl=getTemplateUnknownEventHandler()
  tpl=replace(tpl,"||FUNC-NAME||",funcName)
  monitorClear()
  MonitorAdd(tpl)
end sub


function getTemplateUnknownEventHandler()
#Data myData as String

ERR: EventHandler für '||FUNC-NAME||' nicht gefunden:
§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

Sub ||FUNC-NAME||(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! '||FUNC-NAME||': ...in Arbeit")
  'msgBox("'||FUNC-NAME||': " + "...Coming soon ")
  '...
End Sub


§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
#EndData
  dim RESULT as string=myData 
  return RESULT
end function


