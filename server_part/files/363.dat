dim appIcon as string ="http://es.teamwiki.net/docs/icons/Folder-Downloads.png"

'--> mw_ideIndexList

#runtime EXE




'-->
'--> T E S T   1 - 3 

sub test1()
  msgBox("OK - 1")
end sub

sub test2()
  msgBox("OK - 2")
end sub

sub test3()
  msgBox("OK - 3")
end sub







'-->
'--> C O N F I G - G L O B A L 


'--> Config


#Para SilentMode true
#Para DebugMode intern

Public Const winID = "mw_ideIndexList.vb"

public glob_defaultPath as string="C:\yPara\scriptIDE\scriptPara\"
public glob_defaultGridPath as string="???"










'--> glob

Dim glob As cls_globPara

shared myScriptClass

public globCurTabFileSpec as string
public globCodeListState  As New Dictionary(Of String, String)

Public appGlob As cls_globPara
Public myGlobId as string

Public Const Auto = -2
Public twLoginuser, twLoginPass, twLoginFullname, twSessID As String

public skipNextKeyPress as boolean=false

 
  


'--> NEU: Imports...

#Include <igridEx>
#Include <popupMonitor>


#Reference System.Windows.Forms.dll
#Reference System.Data.dll
#Reference System.Drawing.dll
#Reference WeifenLuo.WinFormsUI.Docking.dll
#Reference TenTec.Windows.iGridLib.iGrid.dll

#Imports System.Windows.Forms
#Imports System.Data
#Imports System.Drawing
#Imports ScriptIDE.Core
#Imports ScriptIDE.ScriptHost
#Imports ScriptIDE.ScriptWindowHelper
#Imports Tentec.Windows.iGridLib

#imports System.IO
#imports Microsoft.VisualBasic.Strings
#imports System.Collections.Generic

#reference C:\Program Files\Microsoft Visual Studio 9.0\Common7\IDE\PublicAssemblies\EnvDTE.dll
#reference C:\Program Files\Microsoft Visual Studio 9.0\Common7\IDE\PublicAssemblies\EnvDTE80.dll
#reference C:\Program Files\Microsoft Visual Studio 9.0\Common7\IDE\PublicAssemblies\EnvDTE90.dll
#Imports EnvDTE
#Imports EnvDTE80
#Imports EnvDTE90


#reference ScintillaNet.dll
#imports ScintillaNet
public IDE As EnvDTE80.DTE2
public WithEvents dte_windowEvents As EnvDTE.WindowEvents
public WithEvents dte_textEditEvents As EnvDTE.TextEditorEvents


'--> Form...

'Dim PanelRef As ScriptedPanel
Dim WithEvents ref As Object
Dim WithEvents FormRef As Form
Dim WithEvents Timer1 As System.Windows.Forms.Timer

dim toolBarColor as string="#E4E1DC"

public shared toolBar1 As ScriptedPanel
public shared toolBar2 As ScriptedPanel
public shared statBar As ScriptedPanel
public shared containerMain As ScriptedPanel
public shared container1 As ScriptedPanel
public shared container2 As ScriptedPanel
public shared container3 As ScriptedPanel

public shared splitContainer1  As Object
public shared splitContainer2  As Object
public shared splitContainer2sub  As Object '@@-

Dim pLeft1, pRight1 As ScriptedPanel
Dim pLeft2, pRight2 As ScriptedPanel
Dim pLeft22, pRight22 As ScriptedPanel


dim skipResizeEvent as boolean = true
dim WithEvents check1 as System.Windows.Forms.CheckBox
dim WithEvents check2 as System.Windows.Forms.CheckBox
dim WithEvents cmbSelIDE as System.Windows.Forms.ComboBox





'--> iGrid

Dim WithEvents IGrid1 As IgridEx
dim igFett = New TenTec.Windows.iGridLib.iGCellStyleDesign
dim  IgDefaultCellStyle = New TenTec.Windows.iGridLib.iGCellStyle(True)




'-->codeProp
public Structure codeProp
  dim scope as string
  dim type as string
  dim procName as string
  dim para as string
  dim returnPara as string
  dim source as string
  dim lines as integer
  dim kom as string
End structure

sub resetCodeProp(x as codeProp)
  x.scope = ""
  x.type = ""
  x.procName = ""
  x.para = ""
  x.returnPara = ""
  x.source = ""
  x.lines = 0
  x.kom = ""
end sub


'--> ...
'--> W I N - A P I

Function isControl() As Boolean
  isControl = False
  If zz.getAsyncKeyState(Keys.ControlKey) Then
    isControl = True
  End If
End Function


Function isShiftControl() As Boolean
  isShiftControl = False
  If zz.getAsyncKeyState(Keys.ShiftKey) And zz.getAsyncKeyState(Keys.ControlKey) Then
    isShiftControl = True
  End If
End Function






'-->
'--> I N I - T E R M I N A T E - R E S I Z E


Sub GetFormRef()
  'If ref IsNot Nothing Then Exit Sub
  ref = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)
  formRef=ref.form
  formRef.Text = "IDE IndexList    (c) mw, es"
End Sub




Sub globAddHandler()
  'AddHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  AddHandler Timer1.Tick, AddressOf onTimerEvent
  AddHandler formRef.Resize, AddressOf Form_Resize
  'AddHandler scNet.MouseDown, AddressOf scNet_MouseDown
  setMonitorRef()
  monitorAdd("globAddHandler")
end sub


Sub globRemoveHandler()
  'trace "globRemoveHandler","try..."
  setMonitorRef()
  monitorAdd("!!! globRemoveHandler")
  'if formRef is nothing then exit sub
  RemoveHandler Timer1.Tick, AddressOf onTimerEvent
  RemoveHandler formRef.Resize, AddressOf Form_Resize
  'RemoveHandler scNet.MouseUp, AddressOf scNet_MouseUp
  'RemoveHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  'trace "globRemoveHandler","DONE"
end sub



Sub Frm_FormClosing(Sender As Object,e As System.Windows.Forms.FormClosingEventArgs) Handles FormRef.FormClosing
  glob.saveFormPos(FormRef)
  glob.saveParaFile()
  trace "formClosing" , "xxx"
  'globRemoveHandler()
End Sub


Sub onTerminate()
  'trace "!!! onTerminate", "!!! es_codeList"
  globRemoveHandler()
End Sub



sub makeJumboLabel(el)
  el.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
  ''el.Size = New System.Drawing.Size(117, 37)
  el.backColor=ColorTranslator.FromHtml("#777")
  el.AutoSize = false
  el.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
end sub




'--> ...

Sub Form_Resize(sender as Object, e as EventArgs)
  on_Resize_Controls()
End Sub



Sub on_Resize_Controls()
  if skipResizeEvent = true then exit sub
  Igrid1.cols(5).width = pLeft1.Width +2
  '' Igrid1.Height = pLeft1.Height - 66
  '' Igrid1.Height = 170
  '' ref.element("test2").top=ref.Height - 28
End Sub



'-->
'--> T I M E R



sub startTimer
  ''msgbox("startTimer")
  '' Timer1 = New System.Windows.Forms.Timer()
  '' timer1.Interval = 555
  '' timer1.Enabled = True
end sub



: Sub onTimerEvent(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
 '' on error resume next
 '' 'exit sub
 ''  static lastFileSpec as string=""
 ''  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
 ''  if lastFileSpec<>ActiveTabFilespec then
 ''    lastFileSpec = ActiveTabFilespec
 ''    globCurTabFileSpec=ActiveTabFilespec
 ''    setMonitorRef()
 ''    monitorAdd(myGlobId)
 ''    monitorAdd(appGlob.para("codeListId"))
 ''    if myGlobId <> appGlob.para("codeListId") then
 ''      monitorAdd("...VERALTET")
 ''      onTerminate()
 ''      exit sub
 ''    end if
 ''    updateCodeList()
 ''    dim name as string=My.Computer.FileSystem.GetName(lastFileSpec)
 ''    ref.element("lab_title").text=name
 ''    switchToCurrentEditor()
 ''  end if
 ''  highlightIndexList()
End Sub


'-->
'--> --------------------------------------------------------



'-->
'--> ==> A U T O S T A R T  - F O R M  1 <== 

Sub AutoStart()
  myScriptClass = me
  
  ''clearAll()
  trace "AutoStart mw_ideIndexList", "1111111111"

  appGlob=zz.ideHelper.glob
  appGlob.para("codeListId")=now.toString
  myGlobId=appGlob.para("codeListId")
  'trace "globTest", glob2.para("myMes"))

  glob = ZZ.newGlobPara()
  
  skipResizeEvent=true
  createForm1()
  skipResizeEvent = False

  startTimer()
  globAddHandler()
  readIndexState(globCodeListState)
  
  on_Resize_Controls()

  formRef.parent.parent.width=300
  splitContainer1.splitterDistance=210
  formRef.parent.parent.width=215

  formRef.show()
  
  fetchIdeList_MouseClick(nothing)
  
  '' pRight1.visible=true
  '' splitContainer1.Panel2Collapsed = false
  '' splitContainer1.splitterDistance=100
  
  'reloadCodeList("dummy")
end sub



Sub createForm1()
  on error resume next
  GetFormRef()
  glob.readFormPos(FormRef)
  dim el as object
  ref.resetControls (10,5)


  
    'toolBarColor="#BFBFBF"
  'toolBarColor="#65CE84"
  dim cMain      As ScriptedPanel
  dim  cToolBar As ScriptedPanel
  dim  cStatBar  As ScriptedPanel

  '--> containerMain
  With ref
    .resetControls ()
    .activateEvents = "|IconMouseDown|   |TextboxKeyDown|"


    containerMain  =.addPanel("containerMain"    , DockStyle.Fill)
    toolBar1     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
    statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=28)

    toolBar1.resetControls  (10,3,1)
    toolbar1.visible=true
    toolbar1.height=30
    toolbar1.height=95
    toolbar1.BackColor = ColorTranslator.FromHtml(toolBarColor)
    toolbar1.BackColor = ColorTranslator.FromHtml("#878813")
    toolbar1.BackColor = ColorTranslator.FromHtml("#53575B")

    'container1.top=112
    containerMain.resetControls  (10,5,1)
    containerMain.resetControls  (10,5,1)
    containerMain.BackColor = ColorTranslator.FromHtml("#fff")
    containerMain.BackColor = ColorTranslator.FromHtml("#ccc")


    statBar.visible=false
    statBar.resetControls  (10,5,1)
    statBar.BackColor = ColorTranslator.FromHtml("#8f8")
    statBar.BackColor = ColorTranslator.FromHtml(toolBarColor)
    statBar.height=30
  end with

  '--> toolbar, oben/unten
  '.BR  '--------------------------------------------------
  
  dim deltaX as integer=25
  with toolBar1 'statBar
    .visible=true
    .height=20
    .resetControls (2,2)
    cmbSelIDE =New System.Windows.Forms.ComboBox
    with cmbSelIDE
    ''check1.Location = New System.Drawing.Point(80, 70)
      .Location = New System.Drawing.Point(120, 0)
      .Size = New System.Drawing.Size(200,20)
      .visible=true
      toolBar1.Controls.Add(cmbSelIDE)
    end with
    .insertX = 80 :.insertY = 0
    .OffsetX = 0 :.insertY = 0
    'el=.addButton  ("cmdToggleExpandMode" , "+ / ---"    , "#8B9198" , , ,50,20): .OffsetX = .OffsetX +50:: el.foreColor=ColorTranslator.FromHtml("#fff")
    '' el=.addButton  ("cmdExpandAll" ,          "+"         , "#49AA36" , , ,50,20): .OffsetX = .OffsetX +50:: el.foreColor=ColorTranslator.FromHtml("#fff")
    '' el=.addButton  ("cmdCollapsAll" ,         "---"       , "#DB3636" , , ,50,20): .OffsetX = .OffsetX +50:: el.foreColor=ColorTranslator.FromHtml("#fff")
    '' el=.addButton  ("syncCodeEditor" ,        "s"         , "#8B9198" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("reloadCodeList" ,        "r"         , "#8B9198" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX+10:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("fetchIdeList" ,          "r2"         , "#8B9198" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX+10:: el.foreColor=ColorTranslator.FromHtml("#fff")
    
  end with


  '--> splitContainer - 1)
  with containerMain  
    
    el=.addSplitcontainer("splMain", "left", pLeft1, "right", pRight1, DockStyle.Fill)
    splitContainer1=el
    el.backColor=ColorTranslator.FromHtml("#88f")
    'el.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
    'el.Panel2Collapsed = true
    'el.splitterDistance=100

  end with 
    
     
  '--> --- Left Pane - 1
  with pLeft1
    .resetControls (5,5)
    '--> --- --- igrid-1
    IGrid1 = New IgridEx
    .Controls.Add(IGrid1)
    
    IGrid1.Dock = DockStyle.Fill
    'IGrid1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    igrid_setDefaultPara(IGrid1)
    with IGrid1
      .Header.Visible = False
      .BorderStyle = TenTec.Windows.iGridLib.iGBorderStyle.None
      igFett.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      .fileSpec=glob_defaultGridPath+"esGrid1.txt"
      '.SelectionMode = TenTec.Windows.iGridLib.iGSelectionMode.MultiExtended
      .rowMode=true
      .HotTracking = False
      '.GroupBox.Visible = True
      .GroupBox.Visible = false
      .SelCellsBackColorNoFocus = ColorTranslator.FromHtml("#4488E9")
      .HScrollBar.Visibility = TenTec.Windows.iGridLib.iGScrollBarVisibility.Hide
      .GridLines.Mode = TenTec.Windows.iGridLib.iGGridLinesMode.None
      .DefaultRow.Height = 16
      .DefaultRow.NormalCellHeight = 14

      '.AllowDrop = True
      .ReadOnly = True
      IGrid1.Cols.Add("groupe")          '0
      IGrid1.Cols.Add("lastGroupHeader") '1
      IGrid1.Cols.Add("groupeState")     '2
      IGrid1.Cols.Add("type")            '3
      IGrid1.Cols.Add("id")              '4
      IGrid1.Cols.Add("nickname")        '5
      IGrid1.Cols.Add("gruppe")          '6
      IGrid1.Cols.Add("sourceLine")      '7
      IGrid1.Cols.Add("x1")              '8
      IGrid1.Cols.Add("x2")              '9
      IGrid1.GroupBox.Visible = True
      IGrid1.Cols(0).width=0
      IGrid1.Cols(1).width=0
      IGrid1.Cols(2).width=0
      IGrid1.Cols(3).width=0
      IGrid1.Cols(4).width=0
      IGrid1.Cols(5).width=210
      IGrid1.Cols(6).width=0
      IGrid1.Cols(7).width=0
      IGrid1.Cols(8).width=0
      IGrid1.Cols(9).width=0
      '.Cols(1).CellStyle.Type = TenTec.Windows.iGridLib.iGCellType.Check
      .rows.count = 100
      '.AutoResizeCols = True
    end with




''   '--> ... listView(außer Betrieb)
''      Dim typName
''     typName = "System.Windows.Forms.ListBox, System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089, processorArchitecture=MSIL"
''     .addControl ("ctrl_lv",typName)
''     .Element("ctrl_lv").Top=5
''     .Element("ctrl_lv").Width=200 : .Element("ctrl_lv").Height=400
''     With .Element("ctrl_lv").Items
''       .add ("test")
''       .add ("test")
''       .add ("test")
''       .add ("test")
''     End With
  end with
 





 
 
  '--> --- Right Pane - 2
  with pRight1
    .resetControls (5,5)
   
   
 '--> icon
 
    .insertX = 10 :.insertY = 7
    .addIcon("appPicture", "http://es.teamwiki.net/docs/icons/insert-object.png" )

   .insertX = 160 :.insertY = 4
   ''Function addLabel(strID,  strText,   bgColor fgColor  intLeft = -1,  intTop = -1, intWidth = -1, intHeight = -1) '@@-
   ''                  el=.addLabel  ("lab01", "Auswertung   UGB - Onlinefragebogen" ,  ,"#ffffff",,,380,31) :makeJumboLabel(el)
                     el=.addLabel  ("lab01", "" ,  ,"#ffffff",,,380,31) :makeJumboLabel(el)
    .insertX = 560 : el=.addLabel  ("curRow",         "zeile" ,  ,"#ffffff" , , ,55 ,31 ) :makeJumboLabel(el)
    .insertX = 620 : el=.addLabel  ("selectionCount", "sel" ,  ,"#ffffff" , , ,88 ,31 ) :makeJumboLabel(el)
 
    'el.width=200
    'el.height=50
    
     .BR  '-------------------------------------------------- '-->inBox
    .insertX = 160 :.insertY = 40
     
    .insertX = 160 :.insertY = 65
    .addButton  ("startCodeList"        , "startCodeList"     , "#6f6")
     .addButton  ("fetchIdeList"      , "get IDE"      , "#ccc")
    .addButton  ("test3"      , "test2"      , "#ccc")
    .addButton  ("test2"      , "test3"      , "#ccc")

    .br  '--------------------------------------------------
    .insertX = 160 :.insertY = 100
    check2 = New System.Windows.Forms.CheckBox
    with check2
    ''check1.Location = New System.Drawing.Point(360, 10)
      .Location = New System.Drawing.Point(160, 100)
      .Size = New System.Drawing.Size(60, 19)
      .Text = "auto"
      .UseVisualStyleBackColor = True
      .visible=true
     pRight1.Controls.Add(check2)
    end with


'-->  textbox
.BR  '-------------------------------------------------- '-->inBox
    .insertX = 5 :.insertY = 140
    .addTextbox ("inBox" ,  -2         , "inBox"    , 0  , "#FFFF99", , , "multiline=190")
       ref.element("txt_inBox").WordWrap=false
       ref.element("txt_inBox").scrollbars=System.Windows.Forms.ScrollBars.Vertical
       
'--> debugbereich
.BR  '--------------------------------------------------
   .insertX = 11 :.insertY = 470
   .addTextbox ("outMonitor" ,  -2         , "debug...:"    , 50  , "#afa", , , "multiline=80")
       ref.element("txt_outMonitor").WordWrap=false
    .br  '--------------------------------------------------
    .addButton  ("clear"        , "clear"     , "#f66")
    .addButton  ("clearGrid"    , "clearGrid" )
    .addButton  ("login"        , "logIn" )
    .addButton  ("readData"     , "Umfrage lesen" )
    .addTextbox ("inummer"      ,  200     , "iNummer:"    , 50  , "#aaa")
    .br  '--------------------------------------------------
 


    ''trace "appPath", application.executablePath
     dim exePath as string
     exePath=application.executablePath
     if exePath.toLower.endsWith("rtftabmdi.exe") then
       formRef.Text = "--> SCRIPT: ------- codeList"
     else
       '' formRef.Text = "web-READER.vb   (c)es, mw"
     end if
     trace "AutoStart", "FERTIG"
  End With
  
  
  '--> End of autostart
  
  
End Sub



'-->
'--> IDE Connect

Sub connectIDE(winTitle as string)
  Dim ideHelper = CreateObject("mwIdeHelper.Application")
  Dim ideref as Object
  ideref = ideHelper.searchIDE(6, winTitle)
  TRACE "ideRef Type",ideref.getType().name
  If ideref Is Nothing Then Exit Sub
  
  IDE = ideref
  dte_textEditEvents = IDE.Events.TextEditorEvents
  dte_windowEvents = IDE.Events.WindowEvents
  
  dim out as string="SOLUTION: "+IDE.solution.fileName+vbnewline
  For i as integer = 1 To IDE.Solution.Projects.Count
    Dim prj As EnvDTE.Project = IDE.Solution.Projects.Item(i)
    out += prj.fileName + vbnewline
  Next
  ref("inBox").Text=out
End Sub

Function getDocumentCode(doc As EnvDTE.Document) As String
  
  
  Dim txtDoc As EnvDTE.TextDocument = CType(doc.Object("TextDocument"), TextDocument)
  
  
  Dim ed As EnvDTE.EditPoint = txtDoc.StartPoint.CreateEditPoint() ''(txtDoc.StartPoint)
  
  
  Return ed.GetText(txtDoc.EndPoint)
  
  
End Function



Sub dte_textEditEvents_LineChanged(StartPoint As TextPoint, EndPoint As TextPoint, Hint As Int32) Handles dte_textEditEvents.LineChanged
  TRACE "dte_textEditEvents_LineChanged"
End Sub
Sub dte_windowEvents_WindowActivated(GotFocus As Window, LostFocus As Window) Handles dte_windowEvents.WindowActivated
  TRACE "00000 dte_windowEvents_WindowActivated",gotFocus.caption
  TRACE "11111 dte_windowEvents_WindowActivated"
  updateCodeList()
  TRACE "22222 dte_windowEvents_WindowActivated"
End Sub


'-->
'--> E V E N T S 



Sub cmbSelIDE_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSelIDE.SelectedIndexChanged
  connectIDE(cmbSelIDE.SelectedItem)
  
End Sub

Sub fetchIdeList_MouseClick(e)
  Dim ideHelper = CreateObject("mwIdeHelper.Application")
  Dim max As Integer = ideHelper.Count
  TRACE "ideRef COUNT",max
  Dim out(max) as string
  For i As Integer = 0 To max-1
    Dim ide As EnvDTE80.DTE2 = ideHelper.getIDERef(i)
    out(i) = ide.MainWindow.Caption
  Next
  cmbSelIDE.Items.clear()
  cmbSelIDE.Items.addRange(out)
End Sub



  Function listProps(ByVal props As EnvDTE.Properties) As String
    On Error Resume Next
    Dim sb As New System.Text.StringBuilder
    For i As Integer = 1 To props.Count
      sb.Append(props.Item(i).Name + ": " + Space(40 - props.Item(i).Name.Length))
      Dim value As Object = props.Item(i).Value
      sb.Append(value.GetType.Name + Space(20 - value.GetType.Name.Length))
      If TypeOf value Is String() Then
        Dim str() As String = value
        sb.Append(Join(str, ";"))
      Else
        sb.Append(value.ToString)
      End If
      sb.AppendLine()
    Next
    Return sb.ToString()
  End Function

Sub onButtonEvent(e)
  setMonitorRef()
  clearAll
  statustext("OK")
  errorText("")
  printLine 11, "e.Sender.Name" , e.Sender.Name
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  MonitorAdd("======================== onButtonEvent")
  MonitorAdd("SenderFullName: " + e.Sender.Name)
  MonitorAdd("___MouseButton: " + e.MouseButton)
  MonitorAdd("________Action: " + action)

  callCmdByName(e)

end Sub



Sub onLabelEvent(e)
  setMonitorRef()
  clearAll
  statustext("OK")
  errorText("")
  printLine 11, "e.Sender.Name" , e.Sender.Name
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  MonitorAdd("=======================- onLabelEvent")
  MonitorAdd("SenderFullName: " + e.Sender.Name)
  MonitorAdd("___MouseButton: " + e.MouseButton)
  MonitorAdd("________Action: " + action)
  
    callCmdByName(e)

end sub



'--> !!!  AUSLAGERN
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
    dim errMes="ERR: Sub '"+funcName+"' nicht gefunden" '@@-
    statustext(errMes)
    dumpUnknownEventHandler(funcName)
    exit sub
  end if
  
  monitorAdd("______procName: "+myMethod.toString)
  ''monitorAdd("AAA: "+err.number.tostring)
  ''monitorAdd("BBB: "+err.description)

  dim args(0)' 
  args(0)=e
  myMethod.Invoke(scriptClass, args)
  'CallByName(scriptClass, funcName, Microsoft.VisualBasic.CallType.Method)
  ' deltaTime=GetTime()-timerStart
  monitorAdd("--------------- DONE")
  'monitorAdd("anz schleifen je sek: "+(i / deltaTime * 1000).toString("0000"))
  'monitorAdd("")
end sub



sub dumpUnknownEventHandler(funcName)
  dim tpl=getTemplateUnknownEventHandler()
  tpl=replace(tpl,"||FUNC-NAME||",funcName)
  MonitorAdd(tpl)
end sub


function getTemplateUnknownEventHandler()
'--> ### data-block...
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






'-->
'--> --------------------------------------------------------


'--> 
'--> C O D E - L I S T




sub reloadCodeList(e as object)
  updateCodeList()
end sub



sub syncCodeEditor(e)
  highlightIndexList()
end sub


sub updateCodeList()  'maloche
  trace "updateCodeList   (...maloche)" 
  
  dim codeEditor = IDE.ActiveDocument
  trace "updateCodeList",codeEditor.fullName
  dim codeText as string=getDocumentCode(codeEditor)
  Trace "TEXT",codeText
  dim RESULT as string = scannCodeData(codeText)
  IGrid1.visible=false
  Igrid_put(IGrid1, RESULT, vbNewLine, vbTab)
  formatCodeList(IGrid1)
  IGrid1.visible=true
  IGrid1.GroupBox.Visible = false
  ''MonitorAdd("-6- DONE........................")
end sub
  
 
 
:function scannCodeData(code as string) as string
  on error resume next
  dim temp as string = code.toLower
  dim LOWER() as string: LOWER=split(temp,vbNewLine)
  dim SOURCE() as string: SOURCE=split(code,vbNewLine)

  dim max = Lower.length
  dim i as integer
  dim codeLine as string
  dim treffer as string=""

  dim OUT(max) as string
  dim counter as integer=0
  dim temp2 as string
  printLine 3, "START", "START"
  dim isGroupHeader as boolean=false
  dim lastGroupeHeader as string=""
  dim details() as string

  monitorAdd(LOWER(0))
  '' if LOWER(0).contains("appicon ")then
  ''   monitorAdd("icon FOUND")
  ''   dim apo as string=""""
  ''   dim part() as string=split(SOURCE(0),apo)
  ''   dim img as string=part(1)
  ''   monitorAdd(img)
  ''   ref.element("pic_appPicture22").image=ZZ.GetImageCached(img)
  '' else
  ''   ref.element("pic_appPicture22").image=ZZ.GetImageCached("http://es.teamwiki.net/docs/icons/Folder-Downloads.png")
  '' end if

  for i = 0 to max-1
    treffer=""
    codeLine=LOWER(i)
    if codeline.contains("s"+"ub ") then treffer="s"
    if codeline.contains("func"+"tion ") then treffer="f"
    if codeline.contains("'-->") then treffer="k"
    if treffer<>"" then
      temp2=scannCodeData_details(codeline,SOURCE(i), treffer,i)
      if temp2<>"" then
        if temp2.startsWith("EMPTY") then
          isGroupHeader=true
          '' OUT(counter)="gg"+vbTab+lastGroupeHeader+vbTab+temp2
          '' counter=counter+1
          continue for
        end if
         
        if isGroupHeader = True then
          isGroupHeader=false
          DETAILS=split("GG"+vbTab+"xxx"+vbTab+vbTab+temp2+vbTab+vbTab+SOURCE(i), vbTab)
          lastGroupeHeader=DETAILS(4)
          DETAILS(1)=lastGroupeHeader
          DETAILS(5)=replace(DETAILS(5), "    ---", "")
          OUT(counter)=join(DETAILS,vbTab)
          counter=counter+1
          continue for
        end if

        OUT(counter)="gg"+vbTab+lastGroupeHeader+vbTab+vbTab+temp2+vbTab+vbTab+SOURCE(i)
        counter=counter+1
        '' trace "codeLine: "+i.toString, temp2
        '' printLine 11, "codeLine: "+i.toString, temp2
      end if
    end if
  next
  'printLine 3, "STOP", "STOP"
  redim preserve OUT(counter)
  dim RESULT as string=join(OUT, vbNewLine)
  return RESULT
end function


:function scannCodeData_details(codeLineLower as string, codeLineSource as string, treffer as string, index as integer)
  on error resume next
  dim nicName as string=""
  scannCodeData_details=""
  dim lineNetto as string
  dim prefix as string=""
  dim subName as string
  select case treffer
    case "k"
      dim kommentar as string=""
      lineNetto=trim(codeLineSource)
      if lineNetto.startsWith("'-->") then
        kommentar = trim(mid(lineNetto,5))
        if kommentar <>"" then 
           prefix="    --- "
        else
           treffer="EMPTY"
        end if
        subName=kommentar
      else
        exit function
      end if
    case else '...sub, function, property   '@@-
      prefix="    "
      if codeLineLower.contains("exit sub") then exit function
      if codeLineLower.contains("end sub") then exit function
      if codeLineLower.contains("exit function") then exit function
      if codeLineLower.contains("end function") then exit function
      if codeLineLower.contains("end property") then exit function
      if codeLineLower.contains("exit property") then exit function
      if codeLineLower.contains("'@@-") then exit function
      nicName=codeLineSource

      dim details() as string
      details=split(nicName+"(","(")
      subName=details(0)
      subName=replace(subName,"sub","",,,CompareMethod.Text)
      subName=replace(subName,"function","",,,CompareMethod.Text)
      subName=replace(subName,"property","",,,CompareMethod.Text)
      subName=replace(subName,"COMEXPORT","",,,CompareMethod.Text)
      subName=replace(subName,"PRIVATE","",,,CompareMethod.Text)
      subName=replace(subName,"PROTECTED","",,,CompareMethod.Text)
      subName=replace(subName,"FRIEND","",,,CompareMethod.Text)
      subName=replace(subName,"PARTIAL","",,,CompareMethod.Text)
      subName=replace(subName,"SHADOWS","",,,CompareMethod.Text)
      subName=replace(subName,"OVERLOADS","",,,CompareMethod.Text)
      subName=replace(subName,"OVERRIDABLE","",,,CompareMethod.Text)
      subName=replace(subName,"OVERRIDES","",,,CompareMethod.Text)
      subName=replace(subName,"MUSTOVERRIDE","",,,CompareMethod.Text)
      subName=replace(subName,"DEFAULT","",,,CompareMethod.Text)
      subName=replace(subName,"SHARED","",,,CompareMethod.Text)
      subName=replace(subName,"READONLY","",,,CompareMethod.Text)
      subName=replace(subName,"WRITEONLY","",,,CompareMethod.Text)

      subName=replace(subName,":","")
      subName=trim(subName)
  end select
 
  dim RESULT = Treffer+vbTab+index.toString+vbTab+prefix+subName+vbTab+codeLineSource
  return RESULT
end function

 
:sub formatCodeList(ByVal grid As TenTec.Windows.iGridLib.iGrid, optional forceState as integer=-1)
  dim forceState2 as boolean=false
  if forceState > -1 then
    if forceState=0 then forceState2=false
    if forceState=1 then forceState2=true
  end if 
  dim ir as iGRow
  Dim max = Grid.Rows.Count - 1
  dim groupe as string
  dim groupeState as boolean = false
  dim key as string
  dim bgColor as string
  bgColor="#444"
  bgColor="#45450B"
  bgColor="#30310B"
  bgColor="#444444"
  For i As Integer = 0 To max
    ir = Grid.Rows(i)
    groupe= ir.Cells(0).Value.ToString

    if groupe="GG" then   '.....Überschrift
      'ir.Cells(5).BackColor=ColorTranslator.FromHtml("#bbb")
      'ir.Cells(5).Style=  grid.style.DefaultFont 'igFett
      ir.Cells(5).BackColor=ColorTranslator.FromHtml(bgColor)
      ir.Cells(5).ForeColor=ColorTranslator.FromHtml("#eee")
      key=ir.Cells(5).value
      
      groupeState=false
      dim key2 as string=globCurTabFileSpec+"|||"+key
      if globCodeListState.ContainsKey(key2) then
        if globCodeListState.item(key2)=">" then 
          groupeState=true
          ir.Cells(2).value=">"
        end if
      end if
    else   '...normaler Inhalt
      if forceState > -1 then groupeState = forceState2 
      ir.visible=groupeState 
    end if
  Next

end sub




: sub highlightIndexList()
  on error resume next
  static lastCurLine as integer
  dim codeEditor = ZZ.getActiveRTF.RTF
  dim curLineNumber as integer=cInt(codeEditor.Lines.Current.number)
  if lastCurLine = curLineNumber then exit sub
  lastCurLine = curLineNumber

  dim i as integer
  dim item as string
  dim id as integer
  dim row as igRow
  dim max as integer
  'monitorAdd("curLineNumber: "+curLineNumber.toString)
  with iGrid1
    max=iGrid1.rows.count - 1
    
    'for i=max to 0 step-20                '...rückwärtsgang
    for i=0 to max                '...rückwärtsgang
      'monitorAdd("i: "+i.toString)
      'if i>max then exit for
      row=iGrid1.rows(i)
      item=trim(row.cells(4).value)
      if item="" then continue for
      'monitorAdd("item: "+item)
      id=cInt(item)
      'monitorAdd("id: "+id.toString)
     if id > curLineNumber then exit for
    next
    
    dim j as integer
    for j=i to 0  step-1      '...vorwärts
   'for j=i to 0 step-1               '...rückwärtsgang
      'monitorAdd("222 j: "+j.toString)
      row=iGrid1.rows(j)
      item=trim(row.cells(4).value)
      if item="" then continue for
      'monitorAdd("222 item: "+item)
      id=cInt(item)
      'monitorAdd("222 id: "+id.toString)
      if val(id) < val(curLineNumber+1) then exit for
    next
    
     '' monitorAdd("        FOUND: "+j.toString)
     '' monitorAdd("curLineNumber: "+curLineNumber.toString)
     '' monitorAdd("           id: "+id.toString)
    iGrid1.curRow=row
    
    exit sub
    
    'DENKFEHLER: inhalt der Zeile ermitteln, die gerade markiert ist
    dim sourceLine as String
    sourceLine=row.cells(6).value+vbNewLine
    dim lineContent as string
    'dim codeEditor = ZZ.getActiveRTF.RTF
    lineContent = trim(codeEditor.Lines.Current.text)
    monitorAdd(sourceLine+"<--")
    monitorAdd(lineContent+"<--")
    if lineContent <>sourceLine then
      monitorAdd("indexliste VERALTET")
    end if
  end with
end sub



'-->   
'--> N A V I G A T E  ---   COLLAPS-EXPAND



: sub navigateCodeLink(grid As iGrid, curRow As iGRow)
  on error resume next
  dim codeLineNumber as integer = cInt(curRow.cells(4).value)
  '' dim lineNumber as integer = cInt(tagDATA(4))
  monitorAdd("navigiere zu: " +codeLineNumber.toString)
  dim codeEitor = ZZ.getActiveRTF.RTF
  'dim lineContent as string = codeEitor.Lines.Current.text
  'codeEitor.Lines.Current.number=codeLineNumber
  codeEitor.goTo.line(codeLineNumber+50)
  ''zz.doEvents
  codeEitor.goTo.line(codeLineNumber-10)
  ''zz.doEvents
  codeEitor.goTo.line(codeLineNumber)
  ''zz.doEvents
  codeEitor.focus()
  err.number=0
End Sub 


: sub navigateForCodeItem(findText as string)
  on error resume next
  'monitorAdd("findText: "+findText)
  dim indexListRow  As iGRow
  indexListRow = findInCodeList(findText)
  if not indexListRow is nothing then
    navigateCodeLink(iGrid1,indexListRow)
  else
    '...such globVars ???
    '...navigate Reflector
  end if
end sub


: function findInCodeList(findText as string)  As iGRow
  on error resume next
  monitorAdd("findText: "+findText)
  dim grid as iGrid=iGrid1
  dim such as string=trim(findText.toLower)
  dim ir as iGRow
  Dim max = grid.Rows.Count - 1
  dim item as string
  dim item2 as string
  For i As Integer = 0 To max
    ir = Grid.Rows(i)
    item= trim(ir.Cells(5).Value.ToString)
    'monitorAdd(item)
    if such=item.toLower then
      monitorAdd("TREFFER: "+i.toString)
      return ir
      'return i
    end if
  Next
  return nothing
end function


'--> !!! ... validate fehlt noch

'--> ...



Sub cmdToggleExpandMode(e) 
  static lastState as boolean=true
  lastState= not lastState
  dim force as integer=0:if lastState=true then force=1
  formatCodeList(iGrid1,force)  
End Sub


Sub cmdExpandAll(e)
  formatCodeList(iGrid1,1)
End Sub



Sub cmdCollapsAll(e) 
  formatCodeList(iGrid1,0)
End Sub


 
Sub IGrid1_CellMouseDown(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseDownEventArgs) Handles IGrid1.CellMouseDown
  dim lastSel As iGRow = sender.curRow
  sender.curRow=sender.rows(e.RowIndex)
  sender.curRow.selected=true
  lastSel.selected=false

  printLine  6, "IGrid1_CellMouseDown - row", e.RowIndex
  printLine  7, "IGrid1_CellMouseDown - col", e.ColIndex
  if e.Button= System.Windows.Forms.MouseButtons.Right then
    printLine  8, "IGrid1_CellMouseDown - col", "rClick"
    monitorAdd("IGrid1_CellMouseDown - col", "rClick")
  else
    printLine  8, "IGrid1_CellMouseDown - col", "NORMAL"
    collapsExpandIgridTopic(sender,  sender.curRow)
    navigateCodeLink(sender,sender.curRow)
  end if
End Sub




sub collapsExpandIgridTopic(grid As iGrid,  selRow  As iGRow )
  dim startIndex as integer = selRow.Index
  if grid.rows(startIndex).cells(0).value<>"GG" then exit sub
  dim max as integer=grid.rows.count
  dim i as integer 
  dim newState as boolean = not grid.rows(startIndex+1).visible
  dim key as string=grid.rows(startIndex).cells(5).value
  dim marker as string
   if newState=true then
     marker=">"
     grid.rows(startIndex).cells(2).value=marker
     'grid.rows(startIndex).cells(5).Style=igFett
     grid.rows(startIndex).cells(5).Style=IgDefaultCellStyle
     grid.rows(startIndex).Cells(5).BackColor=ColorTranslator.FromHtml("#444")
     grid.rows(startIndex).Cells(5).ForeColor=ColorTranslator.FromHtml("#fff")

  else
     marker=""
     grid.rows(startIndex).cells(2).value=marker
     grid.rows(startIndex).cells(5).Style=IgDefaultCellStyle
     'grid.rows(startIndex).Cells(5).BackColor=ColorTranslator.FromHtml("#bbb")
      grid.rows(startIndex).Cells(5).BackColor=ColorTranslator.FromHtml("#444")
     grid.rows(startIndex).Cells(5).ForeColor=ColorTranslator.FromHtml("#fff")
  end if

  dim key2 as string=globCurTabFileSpec+"|||"+key
  if globCodeListState.ContainsKey(key2) then
    globCodeListState.item(key2)=marker
  else
    globCodeListState.add(key2,marker)
  end if
  for i=startIndex+1 to max-1
    if grid.rows(i).cells(0).value="GG" then exit for
    grid.rows(i).visible=newState
  next
  '' dumpDictionary(globCodeListState)   '...für testzwecke
  saveIndexState(globCodeListState)
end sub



sub saveIndexState(dict as Dictionary(Of String, String))
  '--> !!! ungültige Einträge vernichten fehlt noch
  dim key as string
  dim OUT(1000) as string 
  dim counter as integer=0
  dim trenn as string="|||"
  for each key in dict.keys
    'trace dict.item(key),key
    if dict.item(key)=">" then
      OUT(counter)=dict.item(key)+trenn+key
      counter=counter+1
    end if
  next
  redim preserve OUT(counter)

  'dim RESULT=join(OUT,vbNewLine)
  ' monitorAdd(RESULT)
  dim myPath as string =glob_defaultPath
  dim fileSpec as string=myPath+"indexListState.txt"
  Directory.CreateDirectory(myPath)
  IO.File.WriteAllLines(fileSpec, OUT)
end sub


sub readIndexState(dict as Dictionary(Of String, String))
  dim myPath as string =glob_defaultPath
  dim fileSpec as string=myPath+"indexListState.txt"
  Dim fileContent = IO.File.ReadAllText(fileSpec)
  dim lines() as string=split(fileContent,vbNewLine)
  dim fields() as string
  dim item as string
  dim key as string
  dim marker as string=">"
  dim trenn as string="|||"

  dim max as integer=lines.length
  dim i as integer
  for i =0 to max-1
    item=lines(i)+trenn+trenn+trenn
    fields=split(item,trenn)
    key=fields(1)+trenn+fields(2)
    if globCodeListState.ContainsKey(key) then
      globCodeListState.item(key)=marker
    else
      globCodeListState.add(key,marker)
    end if
  next
end sub


sub dumpDictionary(dict as Dictionary(Of String, String))
  dim key as string
  for each key in dict.keys
    trace dict.item(key),key
  next
end sub


'' 
'' 
'' '--> 
'' '--> NEU:  S C I N T I L L A  -  E D I T O R
'' 
'' 
'' 
'' 
'' sub switchToCurrentEditor()
''   setMonitorRef()
''   :RemoveHandler scNet.MouseUp, AddressOf scNet_MouseUp :err.number=0
''   :RemoveHandler scNet.KeyDown,    AddressOf scNet_KeyDown :err.number=0
''   :RemoveHandler scNet.KeyPress,   AddressOf scNet_KeyPress :err.number=0
''   scNet = ZZ.getActiveRTF.RTF
''   sc1= ZZ.getActiveRTF.RTF
''   AddHandler scNet.MouseUp,     AddressOf scNet_MouseUp :err.number=0
''   AddHandler scNet.KeyDown,     AddressOf scNet_KeyDown
''   AddHandler scNet.KeyPress,    AddressOf scNet_KeyPress
'' end sub
'' 
'' 
'' 
'' 
'' 
'' 
'' sub scNet_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
''   'monitorAdd(e.keycode.toString)
''   if e.keyCode=13 then
''     monitorAdd("KeyDown-ENTER")
''     if e.control=true
''        skipNextKeyPress=true
''        monitorAdd("KeyDown-TREFFER")
''     end if
''   end if
'' end sub
'' 
'' 
'' sub scNet_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
''   'monitorAdd(e.KeyChar)
''   if skipNextKeyPress =true then
''     skipNextKeyPress=false
''     e.handled=true
''     monitorAdd("onMyTextArea_KeyPress: CONTROL-ENTER")
''     'executeDirektFenster()
''     'geplant für shortCuts - coming soon
''   end if
'' end sub
'' 
'' 
'' 
'' 
'' '...war mal mouseDown
'' Sub scNet_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
''   'Handles scNet.MouseDown
''   'monitorAdd("scNet_MouseDown")
'' 
''   Dim wordStart, wordEnd As Integer
''   Dim line = scNet.Selection.Range.StartingLine
''   Dim selStart As Integer = line.SelectionStartPosition - line.StartPosition
''   if selStart < 1 then exit sub
''   
''   For wordStart = selStart To 0 Step -1
''     If Char.IsLetterOrDigit(line.Text.Substring(wordStart - 1, 1)) = False Then 
''       if line.Text.Substring(wordStart - 1, 1)<>"_" then Exit For
''     end if
''   Next
''   
''   For wordEnd = selStart To line.Text.Length - 2
''     If Char.IsLetterOrDigit(line.Text.Substring(wordEnd, 1)) = False Then 
''       if line.Text.Substring(wordEnd, 1)<>"_" then Exit For
''     end if
''   Next
'' 
''   Dim wordUnderCursor = line.Text.Substring(wordStart, wordEnd - wordStart)
''   'If wordUnderCursor.Length < 2 Then Exit Sub
''   'navHelpByKeyword(wordUnderCursor.ToLower)
'' 
''   static lastWord as string
''   if lastWord<>wordUnderCursor then
''     lastWord=wordUnderCursor
''     '''monitorAdd("curWord(Scintilla): "+wordUnderCursor)
''     'monitorAdd("curWord(Scintilla): "+wordUnderCursor.ToLower)
''    '''monitorAdd(isControl.toString)
''   end if 
''   if isControl() = true
''     navigateForCodeItem(wordUnderCursor)
''       ''Dim line = scNet.Selection.Range.length=0
''     'zz.doevents
'' 
''     scNet.Selection.length=0
'' 
''     dim lineInt as integer = scNet.Selection.range.startingLine.number
''     ''msgbox(lineInt)
''     highlightExecutingLine(lineInt)
'' 
''   end if
'' End Sub
'' 
'' 
'' '--> ... Lib ???
'' Sub highlightExecutingLine(ByVal line As Integer)
''   'ScintillaNet.MarkerSymbol.ShortArrow
''   'setLineMarker(line, 12, My.Resources.executing, Color.Yellow, Color.Orange, True)
''   setLineMarker(line, 12, 0, Color.Yellow, Color.Orange, True)
''   setLineMarker(line, 13, 0, ColorTranslator.FromHtml("#FFFF7B"), Color.Orange, True) 'vb: #FFEE62
''   If line < 0 Then Exit Sub
''   sc1.GoTo.Line(line)
''   ' sc1.CallTip.HighlightStart = sc1.Lines(line).SelectionStartPosition
''   ' sc1.CallTip.HighlightEnd = sc1.Lines(line).SelectionEndPosition
''   ' sc1.CallTip.HighlightTextColor = Color.Red
''   ' sc1.CallTip.Message = "Pausiert: " & line
''   ' sc1.CallTip.Show()
'' End Sub
''   
''   
'' Sub highlightExecutingLine2(ByVal line As Integer)
''   'ScintillaNet.MarkerSymbol.ShortArrow
''   'setLineMarker(line, 15, My.Resources.executing4, Color.Yellow, Color.Orange, True)
''   setLineMarker(line, 15, 0, Color.Yellow, Color.Orange, True)
''   setLineMarker(line, 13, 0, ColorTranslator.FromHtml("#AFFF7A"), Color.Orange, True) 'vb: #FFEE62
''   If line < 0 Then Exit Sub
''   sc1.GoTo.Line(line)
''   ' sc1.CallTip.HighlightStart = sc1.Lines(line).SelectionStartPosition
''   ' sc1.CallTip.HighlightEnd = sc1.Lines(line).SelectionEndPosition
''   ' sc1.CallTip.HighlightTextColor = Color.Red
''   ' sc1.CallTip.Message = "Pausiert: " & line
''   ' sc1.CallTip.Show()
'' End Sub
''   
''   
'' Sub highlightErrorLine(ByVal line As Integer)
''   setLineMarker(line, 10, ScintillaNet.MarkerSymbol.Arrows, Color.Transparent, Color.Red, True)
''   setLineMarker(line, 11, 0, ColorTranslator.FromHtml("#EAB9FA"), Color.White, True)
''   If line < 0 Then Exit Sub
''   sc1.GoTo.Line(line)
'' End Sub
''   
''   
'' Sub setLineMarker(ByVal line As Integer, ByVal markerIndex As Integer, ByVal markerIcon As Object, ByVal bgColor As Color, ByVal fgColor As Color, Optional ByVal removeOthers As Boolean = False, Optional ByVal toggleMe As Boolean = False)
''   If removeOthers Or line = -1 Then sc1.Markers.DeleteAll(markerIndex)
''   If line < 0 Then Exit Sub
''   If toggleMe = True And (sc1.Lines(line).GetMarkerMask And 16) = 16 Then
''     sc1.Lines(line).DeleteMarker(4) : Exit Sub
''   End If
''   Dim m = sc1.Lines(line).AddMarker(markerIndex)
''   If IsNumeric(markerIcon) Then
''     m.Marker.Symbol = markerIcon
''   ElseIf TypeOf markerIcon Is Image Then
''     m.Marker.Symbol = ScintillaNet.MarkerSymbol.PixMap
''     m.Marker.SetImage(markerIcon, Color.FromArgb(255, 255, 255, 255))
''   Else
''     m.Marker.Symbol = ScintillaNet.MarkerSymbol.PixMap
''     m.Marker.SetImage(Image.FromFile(markerIcon), Color.FromArgb(255, 255, 255, 255))
''   End If
'' 
''   m.Marker.ForeColor = fgColor
''   m.Marker.BackColor = bgColor
'' End Sub
'' 
'' 





'-->
'--> ----------------------------------- L I B s (auslagern) 










