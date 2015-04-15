dim appIcon as string ="http://es.teamwiki.net/docs/icons/Folder-Downloads.png"
'dim appIcon as string ="http://es.teamwiki.net/docs/icons/Folder-Downloads.png"
'dim appIcon as string ="http://es.teamwiki.net/docs/icons/Folder-Downloads.png"
'                  http://es.teamwiki.net/docs/icons/Folder-Downloads.png



'-->
'--> T E S T   1 - 3 


sub test1()

  updateCodeList()
  
  'formatCodeList2(igrid1)

exit sub

  getCodeLink()
  

  dim codeEditor = ZZ.getActiveRTF.RTF
  findListIndexFromLineNumber(codeEditor)
exit sub

  getCodeLink()
  
exit sub
  testCodeLinkFormat()
exit sub
  dim codeLink as string=join(cCodeLinkName,globCodeLinkSplitter) 
  dumpCodeLinkData(codeLink)
  
exit sub
  'getCurCodeLineFromLineNr(1912)

  'msgBox("OK - 1")
  'zz.ideHelper.exec("window.reflection","ZZ","")
  'clearAll

  '' monitorAdd("1: "+lineContent)
  '' monitorAdd("2: "+startPos.toString)
  '' monitorAdd("3: "+lineNumber.toString)
  '' monitorAdd("4: "+column.toString)
  '' monitorAdd("5: "+codeEditor.Lines.Current.number.toString)
  '' monitorAdd("6: "+codeEditor.Lines.firstVisible.toString)
  '' monitorAdd("7: "+codeEditor.Lines.visibleCount.toString)
  '' monitorAdd("8: "+codeEditor.Lines.visibleCount.toString)
  '' monitorAdd("7: "+codeEditor.Lines.visibleCount.toString)
  '' monitorAdd("7: "+codeEditor.Lines.visibleCount.toString)
  '' 


  '--> TEST: goto...
  'codeEditor.Lines.Current.number=codeLineNumber
  'codeEditor.goTo.line(codeLineNumber+50)

  'dim cLine as ScintillaNet.Line = codeEditor.lines(1822)
  'cLine
  codeEditor.goto.line(1820)
  dim startPos = codeEditor.Lines.Current.SelectionStartPosition
  codeEditor.goto.position(startPos+10)
  codeEditor.focus()


  '...test codeZeile rueckwaerts ausgeben
  dim i as integer
  dim start as integer =1850
  dim lineContent2 as string
  for i = start to 1840 step -1
    lineContent2=codeEditor.lines(i).text
    monitorAdd(i.toString+":  "+replace(lineContent2,vbNewLine,""))
  next
end sub






sub test2() '>>>
  ''msgBox("OK - 2")
  monitorAdd("test2()")
  dim mainWin as ScriptIDE.Main.frm_main
  mainWin = ZZ.IDEHelper.MainWindow
  
  dim control as control
  monitorAdd(mainWin.Controls.Count)
  for each control in mainWin.Controls
    'monitorAdd(control.Name)
  next
  mainWin.StatusStrip1.visible = false
  'mainWin.pnlTitlebar.visible = false
  Dim menu As Menustrip
  menu = mainWin.controls.find("MenuStrip1",true)(0)
  menu.visible=false
  'mainWin.MenuStrip1.visible=false
  'mainWin.toolstripContainer1.TopToolStripPanel.visible=false
  'mainWin.DockPanel1.parent.top=59
  
  monitorAdd(mainWin.DockPanel1.parent.top)
  monitorAdd(mainWin.DockPanel1.height)
  monitorAdd(mainWin.height)
  'mainWin.ToolStripContainer1.TopToolStripPanelVisible = False

end sub

sub test3()
  monitorAdd("test2()")
  dim mainWin as object
  mainWin = ZZ.IDEHelper.MainWindow
  dim control as control
  monitorAdd(mainWin.controls.count)
  for each control in mainWin.controls
    'monitorAdd(control.name)
  next
  mainWin.StatusStrip1.visible = true
  mainWin.pnlTitlebar.visible = true
  Dim menu As Menustrip
  menu = mainWin.controls.find("MenuStrip1",true)(0)
  menu.visible=true
  'mainWin.toolstripContainer1.TopToolStripPanel.visible=true

  'mainWin.ToolStripContainer1.visible = true
  'mainWin.DockPanel1.parent.top=49

end sub




sub testCounter
    static counter
  counter=counter+1
  statustext(counter.toString)
  system

end sub


sub testCodeScanner
 ' msgBox("OK - 3")
 'updateCodeList
  
  dim codeEitor = ZZ.getActiveRTF.RTF
  dim codeText as string=codeEitor.text
  dim filespec as string =""
  dim RESULT as string = codeScanner_vb(codeText, fileSpec)
  static counter
  counter=counter+1
  statustext(counter.toString)
end sub









'-->
'--> C O N F I G - G L O B A L 


'--> Config


#Para SilentMode true
#Para DebugMode intern

public lastFileSpec as string=""


public globShowHistoryTimeOut as integer = 5555
public globDumpMode as boolean=false 
Public Const winID = "es_codeList.vb"
Public winTitle =    "codeList    (c) mw, es"

Public Const winID2 = "es_codeList.history"
Public winTitle2 =   "History: ... Betatest"

Public Const winID3 = "es_codeList.bookmarks"
Public winTitle3 =   "Bookmarks: ... zur Zeit noch in Arbeit"

'public glob_defaultPath as string="C:\yPara\scriptIDE\scriptPara\"
public glob_defaultPath as string=""

public glob_defaultGridPath as string="? ? ? "

public globIndexListState as string = "indexListState.txt"
public globIndexListStateFileSpec as string = "..."

public globIndexListColors as string = "indexListColors.txt"
public globIndexListColorsFileSpec as string = "..."

public globIndexListHistory as string = "indexListHistory.txt"
public globIndexListHistoryFileSpec as string = "..."

public globIndexListBookmarks as string = "indexListBookmarks"
public globIndexListBookmarksFileSpec as string = "...wird nicht mehr verwendet"

public globWebAppName as string="scripthost"

'' 'pfad richtig verwalten
'' Dim RESULT As System.String
'' glob_defaultPath = ZZ.IDEHelper.GetSettingsFolder()+"scriptPara\"








'--> G L O B - P A R A 

  Private WithEvents fManager As iGDragDropManager''  = New iGDragDropManager
'    fManager = New iGDragDropManager



Dim glob As cls_globPara

shared myScriptClass

public globCurTabFileSpec as string
public globScannerData  As New Dictionary(Of String, Array)

public globCodeListState   As New Dictionary(Of String, String)
public globCodeListColors  As New Dictionary(Of String, String)


Public appGlob As cls_globPara
Public myGlobId as string

Public Const Auto = -2
Public twLoginuser, twLoginPass, twLoginFullname, twSessID As String

public skipNextKeyPress as boolean=false
public globSkipNextHistoryAction as boolean=false
public globTempExpandCaption as string=""
public globCurSelCaption as string=""
public globLastHighlightLine as integer
public globShowHistoryTimeStamp as integer
public globShowHistoryAutoHide as boolean =true 
public globShowHistoryState as boolean =false



'--> .... Colors

public globCodeListHighlight1(10) as string    'label
public globCodeListHighlight2(10) as string    'grid
public globCodeListHighlight3(10) As System.Drawing.Color   
public globCodeListHighlight4(10) As System.Drawing.Color  

public bgColorNickName  As System.Drawing.Color  
public foreColorNickName  As System.Drawing.Color  
public bgColorNickName2  As System.Drawing.Color  
public foreColorNickName2  As System.Drawing.Color  

public bgColorHighLight  As System.Drawing.Color  
public foreColorHighLight  As System.Drawing.Color  



'--> NEU: Imports...

#Reference siaIDEMain.dll
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

#reference ScintillaNet.dll
#imports ScintillaNet
public WithEvents scNet As ScintillaNet.Scintilla
public WithEvents sc1 As ScintillaNet.Scintilla




'--> Form...

Dim WithEvents panelRef As ScriptedPanel
Dim WithEvents panelRef2 As ScriptedPanel
Dim WithEvents panelRef3 As ScriptedPanel

'Dim WithEvents ref As Object
Dim WithEvents FormRef As Form
Dim WithEvents FormRef2 As Form
Dim WithEvents FormRef3 As Form
Dim WithEvents Timer1 As System.Windows.Forms.Timer

dim toolBarColor as string="#E4E1DC"

public shared toolBar1 As ScriptedPanel
public shared toolBar2 As ScriptedPanel
public shared toolBar3 As ScriptedPanel

public shared statBar As ScriptedPanel
public shared statBar2 As ScriptedPanel
public shared statBar3 As ScriptedPanel

public shared containerMain As ScriptedPanel
public shared containerMain2 As ScriptedPanel
public shared containerMain3 As ScriptedPanel


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
dim WithEvents check3 as System.Windows.Forms.CheckBox

dim WithEvents check11 as System.Windows.Forms.CheckBox
dim WithEvents check22 as System.Windows.Forms.CheckBox
dim WithEvents check33 as System.Windows.Forms.CheckBox

dim WithEvents check111 as System.Windows.Forms.CheckBox
dim WithEvents check222 as System.Windows.Forms.CheckBox
dim WithEvents check333 as System.Windows.Forms.CheckBox





'--> iGrid

Dim WithEvents IGrid1 As IgridEx
Dim WithEvents IGrid2 As IgridEx
Dim WithEvents IGrid3 As IgridEx



dim igFett = New TenTec.Windows.iGridLib.iGCellStyleDesign
dim  IgDefaultCellStyle = New TenTec.Windows.iGridLib.iGCellStyle(True)


'--> C O N S T   -   D A T A 

public iconList() as string



'--> ... iGridFelder

public const cGroupe          = 0
public const cBorder1         = 1
public const cBorder2         = 2
public const cLastGroupHeader = 3
public const cGroupeState     = 4
public const cTyp             = 5
public const cId              = 6
public const cNickname        = 7
public const cFlags          = 8
public const cSourceLine      = 9
public const cCodeFinder      = 10
public const cFileSpec        = 11
public const cHighLighter     = 12
public const cReserve2        = 13
public const cCodeLink        = 14


Public cColName() as string = _  '...
    {"groupe"            , _  ' 0
     "Border1        "   , _  ' 1
     "Border2        "   , _  ' 2
     "lastGroupHeader"   , _  ' 3
     "groupeState"       , _  ' 4
     "typ"               , _  ' 5
     "id"                , _  ' 6
     "nickname"          , _  ' 7
     "Flags"            , _  ' 8
     "sourceLine"        , _  ' 9
     "CodeFinder"        , _  ' 10
     "fileSpec"          , _  ' 11
     "highLighter"       , _  ' 12
     "reserve2"          , _  ' 13
     "CodeLink" }   'LAST     ' 14




'--> ... codeLink

public const globCodeLinkSplitter as string ="_|#|_"

public const ccReserve1        = 0
public const ccVersion          = 1
public const ccTimeStamp        = 2
public const ccReserve2         = 3

public const ccType             = 4
public const ccApp              = 5
public const ccFileSpec         = 6
public const ccReserve3         = 7
public const ccNicName          = 8
public const ccTopic            = 9
public const ccTopicListIndex   = 10
public const ccTopicLineNr      = 11
public const ccTopicSourceLine  = 12
public const ccDeltaLine        = 13
public const ccColPos           = 14
public const ccTargetSourceLine = 15
public const ccReserve4         = 16


Public cCodeLinkName() as string = _  '...
    {"Reserve1"          , _  ' 0
     "Version"           , _  ' 1
     "TimeStamp"         , _  ' 2
     "Reserve2"          , _  ' 3
     "Type"              , _  ' 4
     "App"               , _  ' 5
     "FileSpec"          , _  ' 6
     "Reserve3"          , _  ' 7
     "NicName"           , _  ' 8
     "Topic"             , _  ' 9
     "TopicListIndex"    , _  ' 10
     "TopicLineNr"       , _  ' 11
     "TopicSourceLine"   , _  ' 11
     "DeltaLine"         , _  ' 12
     "ColPos"            , _  ' 13
     "TopicSourceLine"   , _  ' 14
     "Reserve4" }   'LAST     ' 15

  
  '' OUT(ccReserve1)         = ""
  '' OUT(ccVersion)          = ""
  '' OUT(ccTimeStamp)        = ""
  '' OUT(ccReserve2)         = ""
  '' OUT(ccType)             = ""
  '' OUT(ccApp)              = ""
  '' OUT(ccFileSpec)         = ""
  '' OUT(ccReserve3)         = ""
  '' OUT(ccNicName)          = ""
  '' OUT(ccTopic)            = ""
  '' OUT(ccTopicListIndex)   = ""
  '' OUT(ccTopicLineNr)      = ""
  '' OUT(ccTopicSourceLine)  = ""
  '' OUT(ccDeltaLine)        = ""
  '' OUT(ccColPos)           = ""
  '' OUT(ccTargetSourceLine) = ""
  '' OUT(ccReserve4)         = ""



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


'-->
'--> W I N - A P I












Public Declare Function GetTime Lib "winmm.dll" Alias "timeGetTime" () As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short

Function isControl() As Boolean
  isControl = False
  If isKeyPressed(Keys.ControlKey) Then
    isControl = True
  End If
End Function


Function isShift() As Boolean
  isShift = False
  If isKeyPressed(Keys.ShiftKey) Then
    isShift = True
  End If
End Function


Function isShiftControl() As Boolean
  isShiftControl = False
  If isKeyPressed(Keys.ShiftKey) And isKeyPressed(Keys.ControlKey) Then
    isShiftControl = True
  End If
End Function


Function isKeyPressed(ByVal key As Keys) As Boolean
  isKeyPressed = False
  Dim stat As Short
  GetAsyncKeyState(key) 'puffer leeren
  stat = GetAsyncKeyState(key)
  'Debug.Print(key.ToString + vbTab + stat.ToString)
  If stat <> 0 Then
    isKeyPressed = True
  End If
End Function





'-->
'--> I N I - T E R M I N A T E     


function getOuterWindowRef(panelRef as object) as object
  if panelRef.form.parent.parent is Nothing then
    '...falls es kein DockPanel-Fenster ist:
    :return panelRef.form
  else
    '...DockPanel-Fenster haben noch weitere Schichten dazwischen:
    :return panelRef.form.parent.parent
  end if
End function




Sub globAddHandler()
  'AddHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  AddHandler Timer1.Tick, AddressOf onTimerEvent
  AddHandler FormRef.Resize,  AddressOf Form1_Resize
  AddHandler FormRef2.Resize, AddressOf Form2_Resize
  AddHandler FormRef3.Resize, AddressOf Form3_Resize
  'AddHandler scNet.MouseDown, AddressOf scNet_MouseDown
  setMonitorRef()
  monitorAdd("globAddHandler")
end sub


Sub globRemoveHandler()
  'trace "globRemoveHandler","try..."
  setMonitorRef()
  monitorAdd("... globRemoveHandler")
  'if formRef is nothing then exit sub
  RemoveHandler Timer1.Tick,     AddressOf onTimerEvent
  RemoveHandler FormRef.Resize,  AddressOf Form1_Resize
  RemoveHandler FormRef2.Resize, AddressOf Form2_Resize
  RemoveHandler FormRef3.Resize, AddressOf Form3_Resize
  RemoveHandler scNet.MouseUp,   AddressOf scNet_MouseUp
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
  'trace "! ! ! onTerminate", "! ! ! es_codeList"
  globRemoveHandler()
End Sub



sub makeJumboLabel(el)
  el.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
  ''el.Size = New System.Drawing.Size(117, 37)
  el.backColor=ColorTranslator.FromHtml("#777")
  el.AutoSize = false
  el.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
end sub





'-->
'--> T I M E R  -  R E S I Z E 



sub startTimer
  ''msgbox("startTimer")
  Timer1 = New System.Windows.Forms.Timer()
  timer1.Interval = 333
  timer1.Enabled = True
end sub


sub historyAutoHide()
  dim id ="ToolBar|##|tbScriptWin|##|es_codeList.history"
  dim toolWindow=ZZ.getDockPanelRef(id)
  toolWindow.hide
  'globShowHistoryTimeStamp = getTime()
  globShowHistoryState=false
end sub



: Sub onTimerEvent(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick '>>>
 on error resume next
 'exit sub
 
  dim timeStamp as integer = getTime
  if globShowHistoryAutoHide =true then
    if globShowHistoryState = true
      dim deltaTime as integer =timeStamp - globShowHistoryTimeStamp
      'monitorAdd(deltaTime.toString)
      if deltaTime > globShowHistoryTimeOut then
        historyAutoHide()
        globShowHistoryState=false
      end if
    end if
  end if
  
  ''static lastFileSpec as string=""
  static slowTimerEvent as integer=0
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  if lastFileSpec<>ActiveTabFilespec then
    lastFileSpec = ActiveTabFilespec
    globCurTabFileSpec=ActiveTabFilespec
    setMonitorRef()
 
   
    ''monitorAdd(myGlobId)
    ''monitorAdd(appGlob.para("codeListId"))
    if myGlobId <> appGlob.para("codeListId") then
      monitorAdd("...VERALTET")
      onTerminate()
      exit sub
    end if
    updateCodeList()
    dim name as string=My.Computer.FileSystem.GetName(lastFileSpec)
    panelRef.element("lab_title").text=name
    switchToCurrentEditor()
  end if
  'dim foundCaption as string
  if slowTimerEvent>0 then      'sinnvoller Bereich: 0 bis ca. 3
    highlightIndexList()
    slowTimerEvent=0
  else
    slowTimerEvent=slowTimerEvent+1
  end if
  '' end if
  ''   if globTempExpandCaption<>"" and foundCaption<>globTempExpandCaption then
  ''     '...collaps
  ''     monitorAdd("---foundCaption: "+foundCaption)
  ''     monitorAdd("---glob_Caption: "+globTempExpandCaption)
  ''     indexListCollapsTempMode(igrid1, globTempExpandCaption)
  ''     globTempExpandCaption=foundCaption
  ''   end if
End Sub




Sub Form1_Resize(sender as Object, e as EventArgs)
  on_Resize_Controls()
End Sub

Sub Form2_Resize(sender as Object, e as EventArgs)
  on_Resize_Controls()
End Sub

Sub Form3_Resize(sender as Object, e as EventArgs)
  on_Resize_Controls()
End Sub



Sub on_Resize_Controls()
  'exit sub
  if skipResizeEvent = true then skipResizeEvent =false: exit sub
  dim width =pLeft1.Width +2
  'Igrid1.cols(cNickname).width = width
  Igrid2.cols(cNickname).width = containerMain2.width-77
  Igrid3.cols(cNickname).width = containerMain3.width-44

  '' Igrid1.Height = pLeft1.Height - 66
  '' Igrid1.Height = 170
  '' panelRef.element("test2").top=ref.Height - 28
End Sub





'-->
'--> --------------------------------------- F O R M s




'-->
'--> ==> AUTOSTART   

Sub AutoStart()
  myScriptClass = me
  
  ''clearAll()
  trace "AutoStart es_codeList", "1111111111"
  glob_defaultPath = ZZ.IDEHelper.GetSettingsFolder()+"scriptPara\"
  ''msgbox(glob_defaultPath)
  
  appGlob=zz.ideHelper.glob
  appGlob.para("codeListId")=now.toString
  myGlobId=appGlob.para("codeListId")
  'trace "globTest", glob2.para("myMes"))
  glob = ZZ.newGlobPara()

  dim diz as string=getDiz()
  if diz<>"" then diz=diz+"_"
  
  globIndexListStateFileSpec     = glob_defaultPath + diz + globIndexListState
  globIndexListColorsFileSpec    = glob_defaultPath + diz + globIndexListColors
  globIndexListHistoryFileSpec   = glob_defaultPath + diz + globIndexListHistory
  globIndexListBookmarksFileSpec = globIndexListBookmarks
  
  readDictOfStringString(globCodeListState,  globIndexListStateFileSpec)
  readDictOfStringString(globCodeListColors, globIndexListColorsFileSpec)


  skipResizeEvent=true
  iniHighlightColors()
  createForm1()
  skipResizeEvent = False

  startTimer()
  on_Resize_Controls()

  'formRef.parent.parent.width=300
  'splitContainer1.splitterDistance=210
  'formRef.parent.parent.width=215
  'formRef.parent.parent.width=180

  formRef.show()
  'formRef.width=333
  'getOuterWindowRef(panelRef).width=280
  getOuterWindowRef(panelRef).width=254


  '' pRight1.visible=true
  '' splitContainer1.Panel2Collapsed = false
  '' splitContainer1.splitterDistance=100
  
  'reloadCodeList("dummy")
  
  
  createForm2()
  
  dim id ="ToolBar|##|tbScriptWin|##|es_codeList.history"
  ToggleDockWindowFromId(id,0)

  createForm3()
  cmdBookmarkRead(nothing)
  'readGridData(iGrid3)
  'formatCodeList(iGrid3)
  
  on_Resize_Controls()
  globAddHandler()
  
end sub



function getDiz() as string
  dim diz = zz.IDEHelper.DIZ()
  if diz="diz" then
    msgbox("...bitte Diktatzeichen unten rechts eingeben")
    return ""
  end if
  return diz
end function

function getDizPlus() as string
  dim diz = zz.IDEHelper.DIZ()
  if diz="diz" then
    msgbox("...bitte Diktatzeichen unten rechts eingeben")
    return ""
  end if
  return diz+"_"
end function


sub iniIGrid(grid)
  with grid
    .Cols.Add(cColName(0))
    .Cols.Add(cColName(1))
    .Cols.Add(cColName(2))
    .Cols.Add(cColName(3))
    .Cols.Add(cColName(4))
    .Cols.Add(cColName(5))
    .Cols.Add(cColName(6))
    .Cols.Add(cColName(7))
    .Cols.Add(cColName(8))
    .Cols.Add(cColName(9))
    .Cols.Add(cColName(10))
    .Cols.Add(cColName(11))
    .Cols.Add(cColName(12))
    .Cols.Add(cColName(13))
    .Cols.Add(cColName(13))

    .GroupBox.Visible = false
    .Cols(cGroupe).width          = 0
    .Cols(cBorder1).width         = 30
    .Cols(cBorder2).width         = 10
    .Cols(cLastGroupHeader).width = 0
    .Cols(cGroupeState).width     = 0
    .Cols(cTyp).width             = 0
    .Cols(cId).width              = 0
    .Cols(cNickname).width        = 210
    .Cols(cFlags).width          = 0
    .Cols(cSourceLine).width      = 0
    .Cols(cCodeFinder).width      = 0
    .Cols(cFileSpec).width        = 0
    .Cols(cHighLighter).width     = 0
    .Cols(cReserve2).width        = 0
    .Cols(cCodeLink).width        = 0
    '.Cols(1).CellStyle.Type = TenTec.Windows.iGridLib.iGCellType.Check
    .rows.count = 0
    '.AutoResizeCols = True

    '.Cols(cGroupe).CellStyle = New iGCellStyle
    '.Cols(cGroupe).CellStyle.Selectable = igBool.false
    
    
    '' '--> --- ...Selection ausblenden: tut nicht
    '' .Cols(cId).IncludeInSelect     = False
    '' '.Cols(cNickname).IncludeInSelect = true
    .Cols(cGroupe).IncludeInSelect = False
    .Cols(cBorder1).IncludeInSelect = False
    .Cols(cBorder2).IncludeInSelect = False
    '.Cols(cGroupeState).IncludeInSelect = False
    '.Cols(cTyp).IncludeInSelect = False
  end with
end sub


'-->
'--> ==> FORM 1  -  I N D E X - L I S T      



Sub GetFormRef()
  'If panelRef IsNot Nothing Then Exit Sub
  panelRef = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)
  CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = true
  formRef=panelRef.form
  formRef.Text =winTitle
End Sub


function iniHighlightColors()
on error resume next
#Data dataLines as String
''============================================================================================================================================
''labelText |||gridText ||| labelColor ||| gridColor    

  x  |||  --- ||| #43526F  ||| #43526F 
  >  ||| [ > ]||| #FF850A  ||| #FF850A 
  !  ||| [ ! ]||| #00C6AC  ||| #00C6AC 
  ?  ||| [ ? ]||| #7337FF  ||| #7337FF 
  
''============================================================================================================================================
#EndData

  dim DATA() as string=split(dataLines,vbNewLine)
  dim LINE() as string
  dim temp as string
  dim index as integer
  'dim icon,type, para as string
  dim max as integer=DATA.length
  dim i as integer
  dim counter as integer = 1
  for i =0 to max-1
    temp=DATA(i)
    if trim(temp) = "" then continue for
    if trim(temp).startsWith("'") then continue for
    LINE=split(temp, "|||")
    globCodeListHighlight1(counter) = LINE(0)  
    globCodeListHighlight2(counter) = LINE(1) 
    globCodeListHighlight3(counter) = ColorTranslator.FromHtml(trim(LINE(2))) 
    globCodeListHighlight4(counter) = ColorTranslator.FromHtml(trim(LINE(3))) 
    counter=counter+1
  next
end function


Sub createForm1()
  on error resume next
  scNet=nothing
  GetFormRef()
  glob.readFormPos(FormRef)
  dim el as object
  panelRef.resetControls (10,5)

  dim cMain      As ScriptedPanel
  dim  cToolBar As ScriptedPanel
  dim  cStatBar  As ScriptedPanel

  '--> containerMain
  With panelRef
    .resetControls ()
    .activateEvents = "|IconMouseDown|   |TextboxKeyDown|"


    containerMain  =.addPanel("containerMain"    , DockStyle.Fill)
    toolBar1     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
    statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=28)

    with toolBar1
      .resetControls  (10,3,1)
      .visible=true
      .height=30
      .height=95
      .BackColor = ColorTranslator.FromHtml(toolBarColor)
      .BackColor = ColorTranslator.FromHtml("#878813")
      .BackColor = ColorTranslator.FromHtml("#364359")
    end with
    
    with containerMain
      '.top=112
       .resetControls  (10,5,1)
      .BackColor = ColorTranslator.FromHtml("#fff")
      .BackColor = ColorTranslator.FromHtml("#ccc")
    end with

    with statBar
      .visible=false
      .resetControls  (10,5,1)
      .BackColor = ColorTranslator.FromHtml("#8f8")
      .BackColor = ColorTranslator.FromHtml(toolBarColor)
      .height=30
    end with
  end with

  '--> toolbar, oben/unten
  '.BR  '--------------------------------------------------
  
  dim deltaX as integer=25
  with toolBar1 'statBar
    .visible=true
    .height=32' 26
    .resetControls (2,2)
    check1 = New System.Windows.Forms.CheckBox
    with check1
    ''check1.Location = New System.Drawing.Point(80, 70)
      .Location = New System.Drawing.Point(80, 30)
      .Size = New System.Drawing.Size(60, 19)
      .Text = "Spalten"
      .UseVisualStyleBackColor = True
      .visible=true 
      toolBar1.Controls.Add(check1)
    end with
    .insertX = 80 :.insertY = 4
    .OffsetX = 3 :.insertY = 7
    'el=.addButton  ("cmdToggleExpandMode" , "+ / ---"    , "#8B9198" , , ,50,20): .OffsetX = .OffsetX +50:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdReload" ,               "r"        , "#43526F" , ,   ,25,20): .OffsetX = .OffsetX +30:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdHistoryDeltaIndex_2" ,  "<<<"        , "#43526F" , , ,35,20): .OffsetX = .OffsetX +30:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdHistoryDeltaIndex_1" ,   ">>>"       , "#43526F" , , ,35,20): .OffsetX = .OffsetX +30:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdExpandAll" ,             "+"         , "#43526F" , , ,25,20): .OffsetX = .OffsetX +25:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdCollapsAll" ,            "---"       , "#43526F" , , ,25,20): .OffsetX = .OffsetX +25:: el.foreColor=ColorTranslator.FromHtml("#fff")
    'el=.addButton  ("reloadCodeList" ,          "r"         , "#43526F" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    
    .OffsetX = 140: .insertY = 1
    el=.addLabel  ("cmdToggleKom_1",           "[  ] k-1" ,  ,"",,,36,14) :
    'el=.addButton ("cmdToggleKom1" ,          "[_] kom1"   , "#bbb" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el.foreColor=ColorTranslator.FromHtml("#999")
    el.tag=" k-1"
    el=.addLabel  ("cmdToggleKom_2",           "[x] k-2" ,  ,"",,,36,14) :
    'el=.addButton ("cmdToggleKom1" ,          "[_] kom1"   , "#bbb" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el.foreColor=ColorTranslator.FromHtml("#7DFA63")
    el.tag=" k-2"

    dim i as integer=0
   .OffsetX = 140 :.insertY = 16
    i=i+1: el=.addLabel  ("cmdToggleColor_1",  globCodeListHighlight1(i) ,  ,"",,,20,13) :
    el.BackColor= globCodeListHighlight3(i)
    el.ForeColor=ColorTranslator.FromHtml("#ddd")
    el.tag=i
    i=i+1: el=.addLabel  ("cmdToggleColor_2",  globCodeListHighlight1(i) ,  ,"",,,20,13) :
    el.BackColor= globCodeListHighlight3(i)
    el.ForeColor=ColorTranslator.FromHtml("#ddd")
    el.tag=i
    i=i+1: el=.addLabel  ("cmdToggleColor_3",  globCodeListHighlight1(i) ,  ,"",,,20,13) :
    el.BackColor= globCodeListHighlight3(i)
    el.ForeColor=ColorTranslator.FromHtml("#ddd")
    el.tag=i
    i=i+1: el=.addLabel  ("cmdToggleColor_4",  globCodeListHighlight1(i) ,  ,"",,,20,13) :
    el.BackColor= globCodeListHighlight3(i)
    el.ForeColor=ColorTranslator.FromHtml("#ddd")
    el.tag=i



   .insertX = 5  :.insertY = 60
    dim labelText as string ="es_codeList"
    el=.addLabel  ("title", labelText ,  ,"",,,-2,35) :
    makeJumboLabel(el)
    el.backColor=ColorTranslator.FromHtml("#6A6F17")
    el.foreColor=ColorTranslator.FromHtml("#D5E437"): 
  end with


  '--> splitContainer - 1)
  with containerMain  
     el=.addSplitcontainer("splMain", "left", pLeft1, "right", pRight1, DockStyle.Fill)
    splitContainer1=el
    el.backColor=ColorTranslator.FromHtml("#88f")
    'el.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
    el.Panel2Collapsed = true
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
      .name="iGrid1"
      '.Header.Visible = False
      .backColor= ColorTranslator.FromHtml("#68738A")
      .BorderStyle = TenTec.Windows.iGridLib.iGBorderStyle.None
      igFett.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      .fileSpec=globIndexListStateFileSpec
      '.SelectionMode = TenTec.Windows.iGridLib.iGSelectionMode.MultiExtended
      .rowMode=true
      .HotTracking = False
      '.GroupBox.Visible = True
      .GroupBox.Visible = false
      IGrid1.SelCellsBackColorNoFocus = ColorTranslator.FromHtml("#4488E9")
      IGrid1.SelCellsBackColorNoFocus = ColorTranslator.FromHtml("#f44")
      IGrid1.SelCellsBackColorNoFocus = System.Drawing.Color.yellow
      
      IGrid1.SelRowsBackColor = ColorTranslator.FromHtml("#FFF70A")
      IGrid1.SelRowsBackColorNoFocus =  ColorTranslator.FromHtml("#FFF70A")
      IGrid1.SelRowsForeColor = ColorTranslator.FromHtml("#000")
      IGrid1.SelRowsForeColorNoFocus = ColorTranslator.FromHtml("#000")
      
      IGrid1.SelRowsBackColor = ColorTranslator.FromHtml("#F28500")
      IGrid1.SelRowsBackColorNoFocus =  ColorTranslator.FromHtml("#F28500")
      IGrid1.SelRowsForeColor = ColorTranslator.FromHtml("#fff")
      IGrid1.SelRowsForeColorNoFocus = ColorTranslator.FromHtml("#fff")
      
      
      
      
      
      '.SelCellsForeColorNoFocus = System.Drawing.SystemColors.HighlightText

      .HScrollBar.Visibility = TenTec.Windows.iGridLib.iGScrollBarVisibility.Hide
      .GridLines.Mode = TenTec.Windows.iGridLib.iGGridLinesMode.None
      .DefaultRow.Height = 16
      .DefaultRow.NormalCellHeight = 14
      .allowSort=false
      '.AllowDrop = True
      .ReadOnly = True

      iniIGrid(IGrid1)      
      
    end with




''   '--> ... listView(ausser Betrieb)
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
 
  
  '--> End of FORM-1

 
End Sub



'-->
'--> ==> FORM 2  -  H I S T O R Y     


Sub GetFormRef2()
  'If panelRef2 IsNot Nothing Then Exit Sub
  panelRef2 = ZZ.CreateWindow(winID2, DWndFlags.DockWindow, 1)
  formRef2=panelRef2.form
  formRef2.Text = winTitle2
End Sub


Sub createForm2()
  on error resume next
  scNet=nothing
  GetFormRef2()
  glob.readFormPos(FormRef2)
  dim el as object
  panelRef2.resetControls (10,5)
  'toolBarColor="#65CE84"
  dim cMain      As ScriptedPanel
  dim  cToolBar As ScriptedPanel
  dim  cStatBar  As ScriptedPanel


  '--> containerMain
  With panelRef2
    .resetControls ()
    .activateEvents = "|IconMouseDown|   |TextboxKeyDown|"


    containerMain2  =.addPanel("containerMain2"    , DockStyle.Fill)
    toolBar2     =.addPanel("toolBar2"      , DockStyle.Top, intHeight:=33)
    statBar2     =.addPanel("statBar2"      , DockStyle.Bottom, intHeight:=28)
    
    with toolBar2
      .resetControls  (10,3,1)
      .visible=true
      .height=30
      .height=95
      .BackColor = ColorTranslator.FromHtml(toolBarColor)
      .BackColor = ColorTranslator.FromHtml("#878813")
      .BackColor = ColorTranslator.FromHtml("#53575B")
    end with

    with containerMain2
      '.top=112
      .resetControls  (10,5,1)
      .BackColor = ColorTranslator.FromHtml("#fff")
      .BackColor = ColorTranslator.FromHtml("#ccc")
    end with

    with statBar2
      .visible=false
      .resetControls  (10,5,1)
      .BackColor = ColorTranslator.FromHtml("#8f8")
      .BackColor = ColorTranslator.FromHtml(toolBarColor)
      .height=30
    end with
  end with

  '--> toolbar, oben/unten
  '.BR  '--------------------------------------------------
  
  dim deltaX as integer=40
  with toolBar2 'statBar
    .visible=true
    .height=20
    .resetControls (2,2)
    check11 = New System.Windows.Forms.CheckBox
    with check11
    ''check1.Location = New System.Drawing.Point(80, 70)
      .Location = New System.Drawing.Point(80, 30)
      .Size = New System.Drawing.Size(60, 19)
      .Text = "Spalten"
      .UseVisualStyleBackColor = True
      .visible=true 
      toolBar2.Controls.Add(check11)
    end with
    .insertX = 80 :.insertY = 0
    .OffsetX = 0 :.insertY = 0
    deltaX = 30
    'el=.addButton  ("cmdToggleExpandMode2" ,   "+ / ---"     , "#8B9198" , , ,50,20): .OffsetX = .OffsetX +50:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdHistoryAdd" ,        "+"           , "#43526F" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdHistoryRemove" ,     "---"         , "#43526F" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    deltaX = 40
    el=.addButton  ("cmdHistoryClear" ,      "clear"       , "#43526F" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    ''el=.addButton  ("cmdHistorySave" ,       "save"        , "#43526F" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    ''el=.addButton  ("cmdHistoryRead" ,       "read"        , "#43526F" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
  
    'el=.addButton  ("cmdHistoryClear" ,         "clear"       , "#8B9198" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    
    .insertX = 5  :.insertY = 60
    dim labelText as string ="es_codeList"
    el=.addLabel  ("title", labelText ,  ,"",,,-2,35) :
    makeJumboLabel(el)
    el.backColor=ColorTranslator.FromHtml("#6A6F17")
    el.foreColor=ColorTranslator.FromHtml("#D5E437"): 
  end with

'' 
''   '--> splitContainer - 1)
''   with containerMain2  
''     
''     el=.addSplitcontainer("splMain2", "left", pLeft2, "right", pRight2, DockStyle.Fill)
''     splitContainer1=el
''     el.backColor=ColorTranslator.FromHtml("#88f")
''     'el.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
''     el.Panel2Collapsed = true
''     'el.splitterDistance=100
'' 
''   end with 
''     
''   


''   with containerMain2  

''   '--> --- containerMain2  
  with containerMain2  
    .resetControls (5,5)
    '--> --- --- igrid-2
    IGrid2 = New IgridEx
    .Controls.Add(IGrid2)
    
    IGrid2.Dock = DockStyle.Fill
    'IGrid2.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    IGrid2.DefaultCol.CellStyle.TextTrimming = TenTec.Windows.iGridLib.iGStringTrimming.None'=IGCellStyleDesign2
    igrid_setDefaultPara(IGrid1)
    with IGrid2
      .Header.Visible = False
      .BorderStyle = TenTec.Windows.iGridLib.iGBorderStyle.None
      igFett.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      '.SelectionMode = TenTec.Windows.iGridLib.iGSelectionMode.MultiExtended
      .rowMode=true
      .HotTracking = False
      '.GroupBox.Visible = True
      .GroupBox.Visible = false
      '.SelCellsBackColorNoFocus = ColorTranslator.FromHtml("#4488E9")
      .SelCellsBackColorNoFocus = ColorTranslator.FromHtml("#0A246A")
      .SelCellsForeColorNoFocus = ColorTranslator.FromHtml("#fff")
      .HScrollBar.Visibility = TenTec.Windows.iGridLib.iGScrollBarVisibility.Hide
      .GridLines.Mode = TenTec.Windows.iGridLib.iGGridLinesMode.None
      IGrid2.SelRowsBackColor = ColorTranslator.FromHtml("#FFF70A")
      IGrid2.SelRowsBackColorNoFocus =  ColorTranslator.FromHtml("#FFF70A")
      IGrid2.SelRowsForeColor = ColorTranslator.FromHtml("#000")
      IGrid2.SelRowsForeColorNoFocus = ColorTranslator.FromHtml("#000")
      .DefaultRow.Height = 14
      .DefaultRow.NormalCellHeight = 14
      .fileSpec=globIndexListHistoryFileSpec
      .rootDir="notSet"
      .allowSort=false
      .backColor = ColorTranslator.FromHtml("#364359")
      
      IGrid2.DefaultCol.CellStyle.TextTrimming = TenTec.Windows.iGridLib.iGStringTrimming.None'=IGCellStyleDesign2
 
      '.AllowDrop = True
      .ReadOnly = True

      iniIGrid(IGrid2)      
     
    end with
   end with



''   '--> ... listView(ausser Betrieb)
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
''  end with
 
  formRef2.visible=false
  '--> End of Form-2
 
  cmdHistoryRead(nothing)
 
End Sub




'-->
'--> ==> FORM 3  -  B O O K M A R K S      




Sub GetFormRef3()
  'If panelRef2 IsNot Nothing Then Exit Sub
  panelRef3 = ZZ.CreateWindow(winID3, DWndFlags.DockWindow, 1)
  formRef3=panelRef3.form
  formRef3.Text = winTitle3
End Sub



Sub createForm3()
  on error resume next
  scNet=nothing
  GetFormRef3()
  glob.readFormPos(FormRef3)
  dim el as object
  panelRef3.resetControls (10,5)


  
  'toolBarColor="#BFBFBF"
  'toolBarColor="#65CE84"
  dim cMain      As ScriptedPanel
  dim  cToolBar As ScriptedPanel
  dim  cStatBar  As ScriptedPanel

  '--> containerMain
  With panelRef3
    .resetControls ()
    .activateEvents = "|IconMouseDown|   |TextboxKeyDown|"


    containerMain3  =.addPanel("containerMain3"    , DockStyle.Fill)
    toolBar3     =.addPanel("toolBar3"      , DockStyle.Top, intHeight:=33)
    statBar3     =.addPanel("statBar3"      , DockStyle.Bottom, intHeight:=28)

    with toolBar3
      .resetControls  (10,3,1)
      .visible=true
      ''.height=30
      .height=20
      .BackColor = ColorTranslator.FromHtml(toolBarColor)
      .BackColor = ColorTranslator.FromHtml("#878813")
      .BackColor = ColorTranslator.FromHtml("#445370")
    end with

    with containerMain3
      'container1.top=112
      .resetControls  (10,5,1)
      .resetControls  (10,5,1)
      .BackColor = ColorTranslator.FromHtml("#fff")
      .BackColor = ColorTranslator.FromHtml("#ccc")
    end with

    with statbar3
      .visible=true
      .resetControls  (10,5,1)
      .BackColor = ColorTranslator.FromHtml("#445370")
      .height=20
    end with
  end with

  '--> toolbar, oben/unten
  '.BR  '--------------------------------------------------
  
  dim deltaX as integer=40
  
    with toolBar3 'statBar
    .visible=true
    .height=20
    .resetControls (2,2)
    '.insertX = 80 :.insertY = 0
    .OffsetX = 0 :.insertY = 0
    deltaX =30
    el=.addButton  ("cmdNavBookmarks_1" ,       "1"     , "#DB3636" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdNavBookmarks_2" ,       "2"     , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdNavBookmarks_3" ,       "3"     , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdNavBookmarks_4" ,       "4"     , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdNavBookmarks_5" ,       "5"     , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdNavBookmarks_6" ,       "6"     , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdNavBookmarks_7" ,       "7"     , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    
   end with



  '--> statbar, unten
  with statbar3 'statBar
    .visible=true
    .height=20
    .resetControls (2,2)
    check111 = New System.Windows.Forms.CheckBox
    with check111
    ''check1.Location = New System.Drawing.Point(80, 70)
      .Location = New System.Drawing.Point(80, 30)
      .Size = New System.Drawing.Size(60, 19)
      .Text = "Spalten"
      .UseVisualStyleBackColor = True
      .visible=true 
      toolBar3.Controls.Add(check111)
    end with
    .insertX = 80 :.insertY = 0
    .OffsetX = 0 :.insertY = 0
    deltaX =25
    el=.addButton  ("cmdBookmarkDelItem" ,       "---"         , "#DB3636" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    deltaX=50
    el=.addButton  ("cmdBookmarkAddTitle" ,      "+ Titel"     , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdBookmarkAddItem" ,       "+ Zeile"     , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdInsertCurDoc" ,           "+ Doc"       , "#43526F" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    deltaX =25
    'el=.addButton  ("cmdBookmarkAddItem" ,       "Add"         , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    'el=.addButton  ("cmdBookmarkExpandAll" ,     "+"            , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    'el=.addButton  ("cmdBookmarkCollapsAll" ,    "---"            , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    deltaX =40
    el=.addButton  ("cmdBookmarkSave" ,          "save"        , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdBookmarkRead" ,          "read"        , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
    el=.addButton  ("cmdBookmarkClear" ,         "clear"       , "#445370" , , ,deltaX,20): .OffsetX = .OffsetX +deltaX:: el.foreColor=ColorTranslator.FromHtml("#fff")
     
    .insertX = 5  :.insertY = 60
    dim labelText as string ="es_codeList"
    el=.addLabel  ("title", labelText ,  ,"",,,-2,35) :
    makeJumboLabel(el)
    el.backColor=ColorTranslator.FromHtml("#6A6F17")
    el.foreColor=ColorTranslator.FromHtml("#D5E437"): 
  end with



  with containerMain3  
    .resetControls (5,5)


'--> --- --- igrid-3
    IGrid3 = New iGridEx
    .Controls.Add(IGrid3)
    
    IGrid3.Dock = DockStyle.Fill
    'IGrid2.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    igrid_setDefaultPara(IGrid1)
    with IGrid3
      .name="iGrid3"
      .Header.Visible = False
      '.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      ' AUSGESCHALTET: .Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      .BorderStyle = TenTec.Windows.iGridLib.iGBorderStyle.None
      igFett.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      '.SelectionMode = TenTec.Windows.iGridLib.iGSelectionMode.MultiExtended
      .rowMode=true
      .HotTracking = False
      '.GroupBox.Visible = True
      .GroupBox.Visible = false
      '.SelCellsBackColorNoFocus = ColorTranslator.FromHtml("#4488E9")
      .SelCellsBackColor = ColorTranslator.FromHtml("#FFEF58")
      .SelCellsForeColor = ColorTranslator.FromHtml("#000")
      .SelCellsBackColorNoFocus = ColorTranslator.FromHtml("#FFEF58")
      .SelCellsForeColorNoFocus = ColorTranslator.FromHtml("#000")
      .HScrollBar.Visibility = TenTec.Windows.iGridLib.iGScrollBarVisibility.Hide
      .GridLines.Mode = TenTec.Windows.iGridLib.iGGridLinesMode.None
      .DefaultRow.Height = 16
      .DefaultRow.NormalCellHeight = 14
      .fileSpec=globIndexListBookmarksFileSpec+"_1"+".txt"
      '.backColor = ColorTranslator.FromHtml("#7F8AA6")
      .backColor = ColorTranslator.FromHtml("#1B3C84")
      .rootDir="notSet"
      .allowSort=false
      .AllowDrop = True
      .ReadOnly = false

      iniIGrid(IGrid3)
      
      .rows.count = 55

      end with
    end with
    trace "!!! iGDragDropManager", "START"
    fManager = New iGDragDropManager
    fManager.Manage(IGrid3, True, True)
    fManager.AllowMove = True
    fManager.AllowInsert = True
    fManager.AllowMoveWithinOneGrid = True
    trace "!!! fManager.AllowInsert", fManager.AllowInsert

 

  '--> End of Form3
 

 
 
End Sub




'-->
'--> E V E N T S 


Sub onButtonEvent(e)
  setMonitorRef()
  MonitorAdd("= ----- ================== onButtonEvent")
  monitorClear()
  clearAll()
  statustext("----OK----")
  errorText("")
  printLine 11, "e.Sender.Name" , e.Sender.Name
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  MonitorAdd("=+=+=+================== onButtonEvent")
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



'--> --- --- --- ! ! !  AUSLAGERN


sub callCmdByName(e)
  on error resume next
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim funcName as string=name(1)
  trace "funcName", funcName
  dim i as integer=1
  ''dim timerStart = GetTime()
  ''dim deltaTime as integer
  
 
  dim scriptClass= Me
  dim scriptClassName as string=scriptClass.name
  
  
  Dim myType As Type = scriptClass.GetType()
  Dim myMethod As System.Reflection.MethodInfo = myType.GetMethod(funcName)
  if myMethod.toString = "" then
    setLastEventHandlerPara(funcName, scriptClassName, false)
    dim errMes="ERR: Sub '"+funcName+"' nicht gefunden" '@@-
    statustext(errMes)
    dumpUnknownEventHandler(funcName)
    exit sub
  end if

  setLastEventHandlerPara(funcName, scriptClassName, true)

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


sub setLastEventHandlerPara(evName as string, scriptClass as string, found as boolean)
  if evName="cmdInsertEventHandler" then exit sub
  dim sc = zz.scriptClass("es_schnellTester")
  sc.onSetLastEventHandlerPara(evName, scriptClass, found)
end sub


sub xxx_callCmdByName(e)
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



'--> 
'--> E V E N T S - 2   + + +


Sub cmdToggleColor(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  'dim index as integer = val(name(2))
  dim ig = iGrid1
  dim ir = ig.curRow
  dim index as integer = val(e.Sender.tag)

  dim key2 as string=globCurTabFileSpec+"|||"+ir.cells(cNickname).value
  ''msgBox(globCurTabFileSpec)
  ''msgBox(ir.cells(cNickname).value)
  ''msgBox(key2)


  if index=1 then                       '...clear Item
    ir.cells(cBorder1).value =""
    ir.cells(cBorder1).backColor=ColorTranslator.FromHtml("#445370")
    if globCodeListColors.ContainsKey(key2) then
      globCodeListColors.remove(key2)
      saveDictOfStringString(globCodeListColors, globIndexListColorsFileSpec)
    end if
    exit sub
  end if
  
  ir.cells(cBorder1).value =globCodeListHighlight2(index)
  ir.cells(cBorder1).backColor=globCodeListHighlight4(index)
  if globCodeListColors.ContainsKey(key2) then
    globCodeListColors.item(key2)=index.toString
  else
    globCodeListColors.add(key2, index.toString)
  end if
  saveDictOfStringString(globCodeListColors, globIndexListColorsFileSpec)
End Sub



Sub cmdToggleKom(e)
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  
  dim text as string=e.Sender.text
  if text.startsWith("[x]") then
     e.Sender.text="[  ]"+e.Sender.tag.toString
  else
     e.Sender.text="[x]"+e.Sender.tag.toString
  end if
  '...
  statustext("... 'cmdToggleKom': ...in Arbeit")
  'msgBox("'cmdToggleKom': " + "...Coming soon ")
  '...
End Sub



Sub check1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles check1.CheckedChanged
  trace "check1_CheckedChanged", check1.checked
  dim ig = iGrid1
  dim col = IGrid1.CurCell.ColIndex
  dim row = IGrid1.CurCell.RowIndex
  ig.rowmode=not sender.checked
  ig.resetAllSelections(ig)
  if ig.rowmode=true then
    ig.rows(row).selected=true
  else
    ig.cells(row,col).selected=true
  end if
  ig.focus()
End Sub


Sub check2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles check2.CheckedChanged
  trace "check22_CheckedChanged", check1.checked
End Sub


Sub check3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles check3.CheckedChanged
  trace "check333_CheckedChanged", check1.checked
End Sub



'' 
'' '-->
'' '--> H E L P E R
'' 
function ToggleDockWindowFromId(id as string, optional forceState as integer=-1)as boolean
  dim toolWindow=ZZ.getDockPanelRef(id)
  
  '' exit function
  if toolWindow is nothing then exit function
  dim curState as boolean  = toolWindow.visible
  
  ''msgBox(curState)
  dim newState as boolean= not curState
  if forceState=0 then newState=false
  if forceState=1 then curState=true
  dim id2 as string="ToolBar|##|tbScriptWin|##|es_testLabor.mainWin"
  ''monitorAdd(id)
  ''monitorAdd(id2)
  
  'newState=true
  if newState = true then
    toolWindow.show()
    toolWindow.visible =true
    toolWindow.parent.visible =true
    if id=id2 then toolWindow.parent.parent.visible =true   
  else
    toolWindow.hide()
    toolWindow.visible =false
    if id=id2 then  toolWindow.parent.parent.visible =false
  end if
  
  return newState
end function








'-->
'--> -------------------------------------- A P P






'--> 
'--> A P P - H E L P E R 


sub readDictOfStringString(dict as Dictionary(Of String, String), fileSpec as string)
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
    marker=fields(0)
    if dict.ContainsKey(key) then
      dict.item(key)=marker
    else
      dict.add(key,marker)
    end if
  next
end sub


sub saveDictOfStringString(dict as Dictionary(Of String, String), fileSpec as string)
  '!!! ungueltige Eintraege vernichten fehlt noch
  dim key as string
  dim OUT(1000) as string 
  dim counter as integer=0
  dim trenn as string="|||"
  for each key in dict.keys
    'trace dict.item(key),key
    if dict.item(key)<>"<" then
      OUT(counter)=dict.item(key)+trenn+key
      counter=counter+1
    end if
  next
  redim preserve OUT(counter)
  
  dim myPath as string =glob_defaultPath
  Directory.CreateDirectory(myPath)
  IO.File.WriteAllLines(fileSpec, OUT)
end sub




sub dumpDictionary(dict as Dictionary(Of String, String))
  dim key as string
  for each key in dict.keys
    trace dict.item(key),key
  next
end sub



'--> 
'--> S C A N N E R 



Sub cmdReload(e)
  updateCodeList(true)
End Sub




sub reloadCodeList(e as object)
  updateCodeList(true)
end sub



'' sub XXX_syncCodeEditor(e)
''   highlightIndexList()
'' end sub
'' 

sub updateCodeList(optional force as boolean =false)  'maloche
  ''MonitorAdd("updateCodeList   (...maloche)" )
  dim ActiveTabFilespec  as string= ZZ.getActiveTabFilespec()
  'dim fileSpec as string = mid(ActiveTabFilespec,6)
  dim fileSpec as string = ActiveTabFilespec

  dim codeEitor = ZZ.getActiveRTF.RTF
  dim codeText as string=codeEitor.text
  
  if codeText ="" then
    lastFileSpec =""
    '' msgbox("noch leer")
    exit sub
  end if
  
  
  
  'msgBox(codeText)

  dim key=ActiveTabFilespec
  dim RESULT as string
  dim DATA() as string
  
  if globScannerData.ContainsKey(key) then
    if force=false then
      DATA= globScannerData.item(key)
      RESULT=join(DATA,vbNewLine)
    else
      RESULT= codeScanner_Select(codeText,fileSpec)
      DATA=split(RESULT, vbNewLine)
     end if
  Else
    RESULT= codeScanner_Select(codeText,fileSpec)
    DATA=split(RESULT, vbNewLine)
    globScannerData.add(key,DATA)
  end if
  
  test_fillIgrid(IGrid1, RESULT, vbNewLine, vbTab)
  'Igrid_put(IGrid1, RESULT, vbNewLine, vbTab)

  formatCodeList(IGrid1)
  'formatCodeList2(IGrid1)

  IGrid1.visible=true
  IGrid1.GroupBox.Visible = false
  ''MonitorAdd("-6- DONE........................")
end sub
 



'--> TEST: performance(iGrid)
: Sub test_fillIgrid(ByVal Grid As iGrid, ByVal FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
  on error resume next
  Dim out() = Split(FileContent, LineDelim)
  'msgbox(FileContent+"<--")
  Grid.Rows.Clear()
  Grid.Rows.Count = out.Length
  For i As Integer = 0 To out.Length - 1
    test_splitToIGridLine(Grid.Rows(i), out(i), ColDelim)
  Next
  'msgBox(out(out.Length - 2))
  'msgBox(out(out.Length - 1))
  '' if trim(out(out.Length - 1))<>"" then
  ''   SplitToIGridLine(Grid.Rows(out.Length - 1), out(out.Length - 1), ColDelim)
  '' end if
End Sub
 



: Sub test_splitToIGridLine(ByVal line As iGRow, ByVal input As String, Optional ByVal Delimiter As String = vbTab)
  on error resume next
  Dim max as integer
  max = line.Cells.Count - 1
  'max=7
  Dim out() = Split(input, Delimiter)
  ReDim Preserve out(max)
  :For i As Integer = 0 To max
  :  line.Cells(i).Value = out(i)
  :Next
End Sub











'--> 
'--> S C A N N E R - E X T E N T I O N S  


function codeScanner_Select(code as string, fileSpec as string) as string
  dim RESULT as string 
  dim extention as string =System.IO.Path.GetExtension(fileSpec).toLower
  'msgBox(extention)
  select case extention
    case ".vb"
      RESULT = codeScanner_vb(code, fileSpec)
    case ".js", ".htm"
      RESULT = codeScanner_js(code, fileSpec)
    case else
      RESULT = codeScanner_vb(code, fileSpec)
  end select
  return RESULT
end function

'-->.
'-->==> V B   '>>>




:function codeScanner_vb(code as string, fileSpec as string) as string
  on error resume next
  monitorAdd("XXXXXXXXXXXXXXXXXXXXXXXXXXX codeScanner_vb")
  '' printLine 3, "START", "START"
  
  dim temp as string = code.toLower
  dim LOWER() as string: LOWER=split(temp,vbNewLine)
  dim SOURCE() as string: SOURCE=split(code,vbNewLine)
  dim max = LOWER.length
  dim OUT(max) as string
  dim counter as integer=0: dim i as integer
  dim codeLine, treffer, rFlag as string
  dim codeDetails() as string
  dim isGroupHeader as boolean=false
  dim lastGroupeHeader as string=""
  dim komItem as string="'"

  for i = 0 to max-1
    treffer = "": rFlag = ""
    codeLine=LOWER(i)
    if codeline.contains("sub ")          then treffer="s"         '@@-
    if codeline.contains("function")      then treffer="f"         '@@-
    if codeline.contains("property ")     then treffer="p"         '@@-
    if codeline.contains(komItem + "-->") then treffer="k"         '@@-
    'if codeline.contains("//-->")    then treffer="k2"            '@@-

    '...besser machen
    if codeline.contains(komItem + "!!!")      then treffer="!"     '@@-
    if codeline.contains(komItem + "???")      then treffer="?"     '@@-
    if codeline.contains(komItem + ">>>")      then treffer=">"     '@@-
    if treffer="" then continue for
     
    codeDetails=codeScannerDetails_vb(codeline, SOURCE(i), treffer, i, rFlag)
    if rFlag="EXIT" then continue for
    if rFlag="EMPTY" then isGroupHeader=true : continue for

    ''...highlighter: ... veraltet '???
    if treffer="_" and isGroupHeader=false then
      OUT(counter-1)=OUT(counter-1)+vbTab+"!!!KOM!!!"
    end if
   
    if isGroupHeader = True then
      isGroupHeader=false
      codeDetails(cGroupe)="GG"
      lastGroupeHeader=codeDetails(cId)
      codeDetails(cNickname)=replace(codeDetails(cNickname), "    ---", "")
    else
      codeDetails(cGroupe)="gg"
      codeDetails(cBorder1)=codeDetails(cFlags)
      'codeDetails(cNickname)=replace(codeDetails(cNickname), "    ---", "")
    end if
    codeDetails(cLastGroupHeader)=lastGroupeHeader
    codeDetails(cFileSpec)=fileSpec
    OUT(counter)=join(codeDetails,vbTab)
    counter=counter+1
  next
  redim preserve OUT(counter)
  dim RESULT as string=join(OUT, vbNewLine)
  return RESULT


  '...veraltet:
  
  '' DETAILS=split("GG"+vbTab+""+vbTab+""+vbTab+"xxx"+vbTab+vbTab+temp2+vbTab+fileSpec, vbTab)
  '' lastGroupeHeader=DETAILS(6)
  '' DETAILS(3)=lastGroupeHeader
  '' DETAILS(7)=replace(DETAILS(7), "    ---", "")

  '...else
  '' DETAILS=split("gg"+vbTab+vbTab+vbTab+lastGroupeHeader+vbTab+vbTab+temp2+vbTab+fileSpec, vbTab)
  '' DETAILS(cBorder1)=DETAILS(cFlags)

  'printLine 3, "STOP", "STOP"
end function







:function codeScannerDetails_vb(codeLineLower as string, codeLineSource as string, treffer as string, index as integer, byRef rFlag as string) as string()
  on error resume next
  dim fieldsCount as integer=cColName.length
  dim qq(fieldsCount) as string
  dim lineNetto as string
  
  dim subName, nicNamePrefix, codeFinder, komFlags as string
  dim clearCodeFinder as boolean=true
  dim komItem as string="'"

  select case treffer
    case "k"
      dim kommentar as string=""
      codeFinder=""
      lineNetto=trim(codeLineSource)
      if lineNetto.startsWith(komItem + "-->") then  
        kommentar = trim(mid(lineNetto,5))
        if kommentar <>"" then 
           nicNamePrefix="    --- "
        else
           treffer="EMPTY"
           rFlag="EMPTY"
           ''return RESULT2  ' ???
        end if
        subName=kommentar
        codeFinder=subName    ' ???
      
      'wird der noch gebraucht '???
      else  
        rFlag="EXIT"
        return qq
      end if

    case else '...sub, function, property                      '@@-
      nicNamePrefix="     "
      komFlags=""
      dim exitFunc as boolean =false
      
      if codeLineLower.contains("exit sub")        then exitFunc=true
      if codeLineLower.contains("end sub")         then exitFunc=true
      if codeLineLower.contains("exit function")   then exitFunc=true
      if codeLineLower.contains("end function")    then exitFunc=true
      if codeLineLower.contains("end property")    then exitFunc=true
      if codeLineLower.contains("exit property")   then exitFunc=true
      if codeLineLower.contains(komItem + "@@-")   then exitFunc=true
      if exitFunc=true then rFlag="EXIT" : return qq
       
      if treffer="s"  then nicNamePrefix="     "        :komFlags=""      :clearCodeFinder=false
      if treffer="k"  then nicNamePrefix="     "        :komFlags=""      :clearCodeFinder=false
      if treffer="f"  then nicNamePrefix="F   "         :komFlags=""      :clearCodeFinder=false
      if treffer="p"  then nicNamePrefix="p   "         :komFlags=""      :clearCodeFinder=false
      if treffer="!"  then nicNamePrefix="    ... "     :komFlags="   !"     
      if treffer="?"  then nicNamePrefix="    ... "     :komFlags="   ?"    
      if treffer=">"  then nicNamePrefix="    ... "     :komFlags="   >"  :clearCodeFinder=false
      if treffer="_"  then nicNamePrefix="_   "         :komFlags=""
       
      dim details() as string
      details=split(codeLineSource+"(","(")
      subName=details(0)
      subName=replace(subName,"sub","",,,CompareMethod.Text)                  
      subName=replace(subName,"function","",,,CompareMethod.Text)          '@@-
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
      codeFinder=subName
      if clearCodeFinder=true then codeFinder=""  '...steuert die HauptNavigation ueber codeLinks
  end select
  qq(cTyp)        = treffer
  qq(cId)         = index.toString
  qq(cNickname)   = nicNamePrefix+subName
  qq(cFlags)      = komFlags
  qq(cSourceLine) = codeLineSource
  qq(cCodeFinder) = codeFinder
  
  ''' RESULT = Treffer+vbTab+index.toString+vbTab+prefix+subName+vbTab+codeLineSource+vbTab+vbTab+codeFinder
  
  ''dim RESULT = Treffer+vbTab+index.toString+vbTab+prefix+subName+vbTab+komFlags+vbTab+codeLineSource+vbTab+codeFinder
  ''RESULT2=split(RESULT,vbTab)
  return qq
end function











'--> .
'--> ==> J A V A S C R I P T  '>>>




:function codeScanner_js(code as string, fileSpec as string) as string
  on error resume next
  monitorAdd("XXXXXXXXXXXXXXXXXXXXXXXXXXX codeScanner_js")
  '' printLine 3, "START", "START"
  
  dim temp as string = code.toLower
  dim LOWER() as string: LOWER=split(temp,vbNewLine)
  dim SOURCE() as string: SOURCE=split(code,vbNewLine)
  dim max = LOWER.length
  dim OUT(max) as string
  dim counter as integer=0: dim i as integer
  dim codeLine, treffer, rFlag as string
  dim codeDetails() as string
  dim isGroupHeader as boolean=false
  dim lastGroupeHeader as string=""
  dim komItem as string="//"

  for i = 0 to max-1
    treffer = "": rFlag = ""
    codeLine=LOWER(i)
    'if codeline.contains("sub ")         then treffer="s"        '@@-
    if codeline.contains("function")      then treffer="f"        '@@-
    'if codeline.contains("property ")    then treffer="p"        '@@-
    if codeline.contains(komItem + "-->") then treffer="k"        '@@-
 
    '...besser machen
    if codeline.contains(komItem + "!!!")  then treffer="!"       '@@-
    if codeline.contains(komItem + "???")  then treffer="?"       '@@-
    if codeline.contains(komItem + ">>>")  then treffer=">"       '@@-
    if treffer="" then continue for
     
    codeDetails=codeScannerDetails_js(codeline, SOURCE(i), treffer, i, rFlag)
    if rFlag="EXIT" then continue for
    if rFlag="EMPTY" then isGroupHeader=true : continue for

    ''...highlighter: ... veraltet ' ???
    if treffer="_" and isGroupHeader=false then
      OUT(counter-1)=OUT(counter-1)+vbTab+"!!!KOM!!!"
    end if
   
    if isGroupHeader = True then
      isGroupHeader=false
      codeDetails(cGroupe)="GG"
      lastGroupeHeader=codeDetails(cId)
      codeDetails(cNickname)=replace(codeDetails(cNickname), "    ---", "")
    else
      codeDetails(cGroupe)="gg"
      ' ??? codeDetails(cNickname)=replace(codeDetails(cNickname), "    ---", "")
      codeDetails(cBorder1)=codeDetails(cFlags)
      'codeDetails(cNickname)=replace(codeDetails(cNickname), "    ---", "")
    end if
    codeDetails(cLastGroupHeader)=lastGroupeHeader
    codeDetails(cFileSpec)=fileSpec
    OUT(counter)=join(codeDetails,vbTab)
    counter=counter+1
  next
  redim preserve OUT(counter)
  dim RESULT as string=join(OUT, vbNewLine)
  return RESULT
end function






:function codeScannerDetails_js(codeLineLower as string, codeLineSource as string, treffer as string, index as integer, byRef rFlag as string)as string()
  on error resume next
  dim fieldsCount as integer=cColName.length
  dim qq(fieldsCount) as string
  dim lineNetto as string
  
  dim subName, nicNamePrefix, codeFinder, komFlags as string
  dim clearCodeFinder as boolean=true
  dim komItem as string="//"

  select case treffer
    case "k"
      dim kommentar as string=""
      codeFinder=""
      lineNetto=trim(codeLineSource)
      if lineNetto.startsWith(komItem + "-->") then
        
        '... PFUI: das ist noch Murks (mid(6))  '!!! 
        kommentar = trim(mid(lineNetto,6))                   
        if kommentar <>"" then 
           nicNamePrefix="    --- "
        else
           treffer="EMPTY"
           rFlag="EMPTY"
           ''return RESULT2  ' ???
        end if
        subName=kommentar
        codeFinder=subName    ' ???
      
      'wird der noch gebraucht ' ???
      else  
        rFlag="EXIT"
        return qq
      end if

    case else '...sub, function, property                      '@@-
      nicNamePrefix="     "
      komFlags=""
      dim exitFunc as boolean =false
      
      if codeLineLower.contains("//endfunction")   then exitFunc=true  '@@-
      if codeLineLower.contains("end function")    then exitFunc=true  '@@-
      if codeLineLower.contains(komItem + "@@-")    then exitFunc=true
      if exitFunc=true then rFlag="EXIT" : return qq

       
      if treffer="k"  then nicNamePrefix="    --- "     :komFlags=""      :clearCodeFinder=false
      if treffer="f"  then nicNamePrefix="    "         :komFlags=""      :clearCodeFinder=false
      if treffer="!"  then nicNamePrefix="    ... "     :komFlags="   !"     
      if treffer="?"  then nicNamePrefix="    ... "     :komFlags="   ?"    
      if treffer=">"  then nicNamePrefix="    ... "     :komFlags="   >"  :clearCodeFinder=false
      if treffer="_"  then nicNamePrefix="_   "

      dim details() as string
      details=split(codeLineSource+"(","(")
      subName=details(0)
      subName=replace(subName,"window.","",,,CompareMethod.Text)
      subName=replace(subName,"= function","",,,CompareMethod.Text)   '@@-
      subName=replace(subName,"=function","",,,CompareMethod.Text)    '@@-
      subName=replace(subName,"function","",,,CompareMethod.Text)    '@@-

      subName=replace(subName,":","")
      subName=trim(subName)
      codeFinder=subName
      if clearCodeFinder=true then codeFinder=""  '...steuert die HauptNavigation ueber codeLinks
  end select
  
  qq(cTyp)        = treffer
  qq(cId)         = index.toString
  qq(cNickname)   = nicNamePrefix+subName
  qq(cFlags)      = komFlags
  qq(cSourceLine) = codeLineSource
  qq(cCodeFinder) = codeFinder
  
  ''' RESULT = Treffer+vbTab+index.toString+vbTab+prefix+subName+vbTab+codeLineSource+vbTab+vbTab+codeFinder
  
  ''dim RESULT = Treffer+vbTab+index.toString+vbTab+prefix+subName+vbTab+komFlags+vbTab+codeLineSource+vbTab+codeFinder
  ''RESULT2=split(RESULT,vbTab)
  return qq
end function







'--> 
'--> I N D E X - L I S T E




 
Sub IGrid1_CellMouseDown(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseDownEventArgs) Handles IGrid1.CellMouseDown
  dim lastSel As iGRow = sender.curRow
  sender.curRow=sender.rows(e.RowIndex)
  sender.curRow.selected=true
  lastSel.selected=false
  dim isCaption as boolean =false
  
  printLine  6, "IGrid1_CellMouseDown - row", e.RowIndex
  printLine  7, "IGrid1_CellMouseDown - col", e.ColIndex
  if e.Button= Windows.Forms.MouseButtons.Right then
    printLine  8, "IGrid1_CellMouseDown - col", "rClick"
    monitorAdd("IGrid1_CellMouseDown - col", "rClick")
  else
    printLine  8, "IGrid1_CellMouseDown - col", "NORMAL"
    isCaption = collapsExpandIgridTopic(sender,  sender.curRow)
    if isCaption=false then navigateCodeLink(sender,sender.curRow)
  end if
End Sub



Private Sub IGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles IGrid1.DoubleClick
  monitorAdd("IGrid1__DoubleClick", "xxx")
  'msgBox("...vorgesehen für addBookmark")
  cmdBookmarkAddItem(nothing)
End Sub




: sub navigateCodeLink(grid As iGrid, curRow As iGRow)
  on error resume next


  ' ...wird von der indexListe Verwendet  
  ' ... und rClick vom codeEditor         
  
  dim codeLinkItem() as string=getCodeLinkItem(grid,curRow)
  'msgBox(join(codeLinkItem,vbnewLine))



  '...switch Tab
  dim fileSpec as string=curRow.cells(cFileSpec).value
  ''msgBox(fileSpec)
  dim tab = ZZ.IDEHelper.NavigateFile(fileSpec)

  
  
  '!!! ...ueberarbeiten
  dim codeLineNumber as integer = cInt(curRow.cells(cId).value)
  
  '' dim lineNumber as integer = cInt(tagDATA(4))
  
  ''monitorAdd("navigiere zu: " +codeLineNumber.toString)
  dim codeEitor = ZZ.getActiveRTF.RTF
  'dim lineContent as string = codeEitor.Lines.Current.text
  'codeEitor.Lines.Current.number=codeLineNumber
  codeEitor.goTo.line(codeLineNumber+50)
  ''zz.doEvents
  codeEitor.goTo.line(codeLineNumber-10)
  ''zz.doEvents
  codeEitor.goTo.line(codeLineNumber)
  ''zz.doEvents
  highlightExecutingLine(codeLineNumber)
  
  codeEitor.focus()
  
  historyAddItem(codeLinkItem)

  err.number=0
End Sub 



: Function indexListExpandTempMode(grid As iGrid, rowIndex as integer, showMode as boolean)

  ' ...find startIndex(caption)
  dim i as integer
  dim max as integer = grid.rows.count
  dim item as string
  dim foundIndex as integer=0
  dim foundCaption as string=""
  for i=rowIndex to 0 step -1
    item = grid.rows(i).cells(cGroupe).value
    if item="GG" then
      foundIndex=i
      foundCaption=grid.rows(i).cells(cNickname).value
      exit for
    end if
  next

  ' ...loopStart
  dim bgColor = ColorTranslator.FromHtml("#E9E9FF") 
  bgColor = ColorTranslator.FromHtml("#F28500") 
  bgColor = ColorTranslator.FromHtml("#fff") 
  dim row as igRow
  if showMode = true then
    for i = foundIndex+1 to max-1
      row = grid.rows(i)
      item = row.cells(cGroupe).value
      if item="GG" then
        foundIndex=i
        exit for
      end if
      row.visible=true
      row.cells(cBorder2).backColor=ColorTranslator.FromHtml("#A9B0BD") 
      row.cells(cNickname).backColor=bgColor
      dim ir =row
        if ir.Cells(cHighLighter).Value<>"" and not nicName.startsWith("'") then
        ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#F5D7FF")
      end if
      if nicName.contains("'") then
        ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#ddd")
      end if
      if nicName.startsWith("!") then
        ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#F7ED9B")
      end if
      if nicName.startsWith("?") then
        ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#CEF59B")
      end if
      if nicName.startsWith(">") then
        ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#C965DF")
        'ir.Cells(cNickname).ForeColor=ColorTranslator.FromHtml("#fff")
      end if
      if nicName.startsWith("_") then
        ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#FCF7B7")
        'ir.Cells(cNickname).ForeColor=ColorTranslator.FromHtml("#fff")
      end if

    next
  end if
  return foundCaption
end Function




: sub indexListCollapsTempMode(grid As iGrid, lastTempExpandCaption as string)
  dim findText as string=lastTempExpandCaption
  on error resume next
  
  ''monitorAdd("##################################################")
  ''monitorAdd("indexListCollapsTempMode: "+findText)
  dim such as string=trim(findText.toLower)
  dim ir as iGRow
  dim foundIndex as integer
  Dim max = grid.Rows.Count
  dim i as integer
  dim item as string
  dim item2 as string
  
  'findCaption
  For i = 0 To max - 1
    ir = Grid.Rows(i)
    item= trim(ir.Cells(cNickname).Value)
    'monitorAdd(item)
    if such=trim(item.toLower) then
      'monitorAdd("TREFFER: "+i.toString)
        if globDumpMode=true then  monitorAdd("TREFFER(Collaps...): "+trim(item))
      foundIndex=i
      exit for
    end if
  Next

  ' ...loopStart
  'dim bgColor = ColorTranslator.FromHtml("#DDDDFF")
  'dim bgColor = ColorTranslator.FromHtml("#B7B7EA")
  dim bgColor = ColorTranslator.FromHtml("#fff")
  for i = foundIndex+1 to max-1
    ir = grid.rows(i)
    item = ir.cells(cGroupe).value
    if item="GG" then
      'foundIndex=i
      exit for
    end if
    ir.visible=false
    ir.cells(cNickname).backColor=bgColor
  next
end sub




function getCodeLinkItem(grid As iGrid,Row As iGRow)
  
  'NEU...
  checkForIndexListUpdate(grid, row.index)
  dim i as integer
  dim max as integer=grid.cols.count
  dim OUT(max)as string
  for i =0 to max-1
    OUT(i)=row.cells(i).value
  next
  return OUT
end function




'-->
'--> D E S I G N - H I G H L I G H T  




sub iniGridColors(ByVal grid As iGridEx)
  if grid.name="iGrid1" then
    'msgBox("TREFFER")
     bgColorNickName     = ColorTranslator.FromHtml("#445370")  '..caption
     foreColorNickName   = ColorTranslator.FromHtml("#fff")
     bgColorNickName2    = ColorTranslator.FromHtml("#fff")     '...details
     foreColorNickName2  = ColorTranslator.FromHtml("#000")
     bgColorHighLight    = ColorTranslator.FromHtml("#CD7000")
     foreColorHighLight  = ColorTranslator.FromHtml("#fff")
  end if
  if grid.name="iGrid3" then
     bgColorNickName = ColorTranslator.FromHtml("#142D6F")  '..caption
     foreColorNickName = ColorTranslator.FromHtml("#fff")
     bgColorNickName2 = ColorTranslator.FromHtml("#2450A4")    '...details
     foreColorNickName2 = ColorTranslator.FromHtml("#ddd")
     bgColorHighLight    = ColorTranslator.FromHtml("#DFD14D")
     foreColorHighLight  = ColorTranslator.FromHtml("#000")
  end if
end sub



: sub formatCodeList(ByVal grid As TenTec.Windows.iGridLib.iGrid, optional forceState as integer=-1) 
  on error resume next
  'exit sub
  '!!! trennen in d e f a u l t  -  s p e z i a l 
   iniGridColors(grid)
  '' dim bgColorNickName  As System.Drawing.Color  
  '' dim foreColorNickName  As System.Drawing.Color  
  '' dim bgColorNickName2  As System.Drawing.Color  
  '' dim foreColorNickName2  As System.Drawing.Color  
  '' bgColorHighLight    = ColorTranslator.FromHtml("#DFD14D")
  '' foreColorHighLight  = ColorTranslator.FromHtml("#fff")

  dim forceState2 as boolean=false
  if forceState > -1 then
    if forceState=0 then forceState2=false
    if forceState=1 then forceState2=true
  end if 
  dim ir as iGRow
  Dim max = Grid.Rows.Count - 1
  dim groupe as string
  
  dim groupeState as boolean = true  'startVerhalten ...
  dim key as string
  dim bgColor as string
  dim nicName
  
  
  
  bgColor="#444"
  bgColor="#45450B"
  bgColor="#30310B"
  bgColor="#444444"
  bgColor="#445370"
  'dim bgColor2="#ddf"
  'dim bgColor2="#D0D7E7"
  dim bgColor2="#fff"
  For i As Integer = 0 To max
    ir = Grid.Rows(i)
    ir.ContentIndent= New TenTec.Windows.iGridLib.iGIndent(1, 0, 1, 0)
    groupe= ir.Cells(cGroupe).Value.ToString
    
    ir.Cells(cBorder1).backColor = bgColorNickName   'ColorTranslator.FromHtml("#445370")
    ir.Cells(cBorder1).foreColor = foreColorNickName 'ColorTranslator.FromHtml("#fff")
    ir.Cells(cBorder2).backColor = bgColorNickName   ' ColorTranslator.FromHtml("#445370")
    
    '??? textalign... ir.Cells(cBorder2).textalign = ...
 

   key=ir.Cells(cNickname).value
   dim key2 as string=globCurTabFileSpec+"|||"+key
   if groupe="GG" then   '.....ueberschrift
     ''ir.height= 14
    
      'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#bbb")
      'ir.Cells(cNickname).Style=  grid.style.DefaultFont 'igFett
      
      ir.Cells(cNickname).BackColor= bgColorNickName     'ColorTranslator.FromHtml(bgColor)
      ir.Cells(cNickname).ForeColor=foreColorNickName   'ColorTranslator.FromHtml("#eee")
      
      
      groupeState=false
      if globCodeListState.ContainsKey(key2) then
        if globCodeListState.item(key2)=">" then 
          groupeState=true
          ir.Cells(cGroupeState).value=">"
        end if
      end if
      
   
    else   '...normaler Inhalt
       ir.height= 14
      if forceState > -1 then groupeState = forceState2 
      ir.visible=groupeState 
      'ir.Cells(cLastGroupHeader).BackColor=ColorTranslator.FromHtml(bgColor2)
      'ir.Cells(cLastGroupHeader).backColor = ColorTranslator.FromHtml("#445370")
     
      
    
      ir.Cells(cBorder2).backColor = bgColorNickName     'ColorTranslator.FromHtml("#fff")
      ir.Cells(cNickname).BackColor=bgColorNickName2     'ColorTranslator.FromHtml(bgColor2)
      ir.Cells(cNickname).ForeColor=foreColorNickName2   'ColorTranslator.FromHtml(bgColor2)
     
      ' continue for
     
      'nicName= trim(ir.Cells(cNickname).Value.ToString)
      nicName= trim(ir.Cells(cBorder1).Value.ToString)
      'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml(bgColor2)
 
      if ir.Cells(cHighLighter).Value<>"" and not nicName.startsWith("'") then
        'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#F5D7FF")
        ir.Cells(cBorder1).BackColor=ColorTranslator.FromHtml("#F5D7FF")
      end if
      if nicName.contains("'") then
        'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#ddd")
        ir.Cells(cBorder1).BackColor=ColorTranslator.FromHtml("#ddd")
         ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#eee")
         ir.Cells(cNickname).foreColor=ColorTranslator.FromHtml("#777")
      end if
      if nicName.startsWith("!") then
        'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#F7ED9B")
        ir.Cells(cBorder1).BackColor=ColorTranslator.FromHtml("#007ACC")
         ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#eee")
         ir.Cells(cNickname).foreColor=ColorTranslator.FromHtml("#777")
      end if
      if nicName.startsWith("?") then
        'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#CEF59B")
        ir.Cells(cBorder1).BackColor=ColorTranslator.FromHtml("#0B6900")
         ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#eee")
         ir.Cells(cNickname).foreColor=ColorTranslator.FromHtml("#777")
      end if
      if nicName.startsWith(">") then
         ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#C965DF")
         'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#FFFF44")
         ir.Cells(cBorder1).BackColor=ColorTranslator.FromHtml("#ACB200")
       'ir.Cells(cNickname).ForeColor=ColorTranslator.FromHtml("#fff")
      end if
      if nicName.startsWith("_") then
        'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#FCF7B7")
        ir.Cells(cBorder1).BackColor=ColorTranslator.FromHtml("#FCF7B7")
        'ir.Cells(cNickname).ForeColor=ColorTranslator.FromHtml("#fff")
      end if
      
    end if
    
    dim index as integer
    if globCodeListColors.ContainsKey(key2) then
      ''monitorAdd(globCodeListColors.item(key2))
      index=val(globCodeListColors.item(key2))
      ' msgBox(globCodeListColors.item(key2))
      ir.Cells(cBorder1).value=globCodeListHighlight2(index)
      ir.Cells(cBorder1).backColor=globCodeListHighlight4(index)
    end if
  Next

     
      grid.Rows(0).ContentIndent= New TenTec.Windows.iGridLib.iGIndent(1, 0, 1, 0)

end sub



: sub formatCodeList2(ByVal grid As TenTec.Windows.iGridLib.iGrid, optional forceState as integer=-1) 
  on error resume next
  dim forceState2 as boolean=false
  if forceState > -1 then
    if forceState=0 then forceState2=false
    if forceState=1 then forceState2=true
  end if 
  dim ir as iGRow
  Dim max = Grid.Rows.Count - 1
  dim groupe as string
  
  dim groupeState as boolean = true  'startVerhalten ...
  dim key as string
  dim bgColor as string
  dim nicName
  
  
  bgColor="#444"
  bgColor="#45450B"
  bgColor="#30310B"
  bgColor="#444444"
  bgColor="#445370"
  'dim bgColor2="#ddf"
  'dim bgColor2="#D0D7E7"
  dim bgColor2="#fff"
  For i As Integer = 0 To max
    ir = Grid.Rows(i)
    groupe= ir.Cells(cGroupe).Value.ToString
    
    '''ir.ContentIndent= New TenTec.Windows.iGridLib.iGIndent(1, 0, 1, 0)
    '''ir.Cells(cBorder1).backColor = ColorTranslator.FromHtml("#445370")
    '''ir.Cells(cBorder1).foreColor = ColorTranslator.FromHtml("#fff")
    '''ir.Cells(cBorder2).backColor = ColorTranslator.FromHtml("#445370")
    
    '??? textalign... ir.Cells(cBorder2).textalign = ...
 

    if groupe="GG" then   '.....ueberschrift
    
    ''' ir.height= 14
    '''key=ir.Cells(cNickname).value
    
      'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#bbb")
      'ir.Cells(cNickname).Style=  grid.style.DefaultFont 'igFett
      ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml(bgColor)
      ir.Cells(cNickname).ForeColor=ColorTranslator.FromHtml("#eee")
      
      
      groupeState=false
      dim key2 as string=globCurTabFileSpec+"|||"+key
      if globCodeListState.ContainsKey(key2) then
        if globCodeListState.item(key2)=">" then 
          groupeState=true
          ir.Cells(cGroupeState).value=">"
        end if
      end if
    
    
    else   '...normaler Inhalt
     continue for
    
      ir.height= 14
      if forceState > -1 then groupeState = forceState2 
      ir.visible=groupeState 
      'ir.Cells(cLastGroupHeader).BackColor=ColorTranslator.FromHtml(bgColor2)
      'ir.Cells(cLastGroupHeader).backColor = ColorTranslator.FromHtml("#445370")
     ir.Cells(cBorder2).backColor = ColorTranslator.FromHtml("#fff")
     nicName= trim(ir.Cells(cNickname).Value.ToString)
      'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml(bgColor2)
      ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml(bgColor2)
 
      if ir.Cells(cHighLighter).Value<>"" and not nicName.startsWith("'") then
        'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#F5D7FF")
        ir.Cells(cBorder1).BackColor=ColorTranslator.FromHtml("#F5D7FF")
      end if
      if nicName.contains("'") then
        'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#ddd")
        ir.Cells(cBorder1).BackColor=ColorTranslator.FromHtml("#ddd")
      end if
      if nicName.startsWith("!") then
        'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#F7ED9B")
        ir.Cells(cBorder1).BackColor=ColorTranslator.FromHtml("#007ACC")
      end if
      if nicName.startsWith("?") then
        'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#CEF59B")
        ir.Cells(cBorder1).BackColor=ColorTranslator.FromHtml("#0B6900")
      end if
      if nicName.startsWith(">") then
         'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#C965DF")
         'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#FFFF44")
         ir.Cells(cBorder1).BackColor=ColorTranslator.FromHtml("#ACB200")
       'ir.Cells(cNickname).ForeColor=ColorTranslator.FromHtml("#fff")
      end if
      if nicName.startsWith("_") then
        'ir.Cells(cNickname).BackColor=ColorTranslator.FromHtml("#FCF7B7")
        ir.Cells(cBorder1).BackColor=ColorTranslator.FromHtml("#FCF7B7")
        'ir.Cells(cNickname).ForeColor=ColorTranslator.FromHtml("#fff")
      end if
      
    end if
  Next

     
      grid.Rows(0).ContentIndent= New TenTec.Windows.iGridLib.iGIndent(1, 0, 1, 0)

end sub



'--> ...

: function highlightIndexList() 
  on error resume next
  'static lastCurLine as integer
  dim codeEditor = ZZ.getActiveRTF.RTF
  dim curLineNumber as integer=cInt(codeEditor.Lines.Current.number)
  
  
  '???  . . . alles OK, vereinfachter updateCheck  ???
  if globLastHighlightLine = curLineNumber then exit function                             '... E X I T 
  globLastHighlightLine = curLineNumber

  

  '' !!! vereinfachtes Verfahren bei... 
  '' - on Click,   
  '' - cursor hoch / tief
  dim grid  As iGrid = iGrid1
  dim max as integer = grid.rows.count
  dim ir as igRow=grid.curRow
  dim rowIndex1 as integer=ir.index
  dim rowIndex2 as integer=rowIndex1+1
  if rowIndex2>max-1 Then rowIndex2=rowIndex1
  
  dim lineNr1=cInt(ir.cells(cId).value)
  dim lineNr2=cInt(grid.rows(rowIndex2).cells(cId).value)

  if lineNr1-1 < curLineNumber and curLineNumber < lineNr2+0 then
    '...alles OK, vereinfachter updateCheck ???
    '' monitorAdd("LINE: ============================= OK")
    exit function
  else
    '' monitorAdd("...")
    '' monitorAdd("LINE-1: "+lineNr1.toString)
    '' monitorAdd("LINE-2: "+curLineNumber.toString)
    '' monitorAdd("LINE-3: "+lineNr2.toString)
    '' monitorAdd("LINE: >>>>>>>>>>>>>>>>>>>> U P D A T E ")
  end if
  
  
  '' monitorAdd("#####(highlight): "+globTempExpandCaption)

  dim i as integer
  dim item as string
  dim id as integer
  dim row as igRow
  dim foundIndex as integer=-1



  '--> SUCHE-1:  vorwaerts
  if  curLineNumber > lineNr1 then
      if globDumpMode=true then  monitorAdd("SUCHE-1:  vorwaerts...")
    max=rowIndex1+5  
    if max>grid.rows.count then max = grid.rows.count
    with grid
      for i=rowIndex1-1 to max          
        row=iGrid1.rows(i)
        item=trim(row.cells(cId).value)
        if item="" then continue for
        id=cInt(item)
       if id > curLineNumber then 
         foundIndex=(i-1)
         exit for
       end if
      next
    end with
  end if
  

  '--> SUCHE-2:  rueckwaerts
  if  curLineNumber < lineNr2 then
      if globDumpMode=true then  monitorAdd("SUCHE-2:  rueckwaerts...")
    max=rowIndex2
    dim start as integer=rowIndex2-5
    if start<0 then start=0
    with grid
      'max=iGrid1.rows.count - 1
      for i=max to start  step-1        
        row=iGrid1.rows(i)
        item=trim(row.cells(cId).value)
        if item="" then continue for
        'monitorAdd("item: "+item)
        id=cInt(item)
        'monitorAdd("id: "+id.toString)
       if id < curLineNumber then 
         foundIndex=(i)
         exit for
       end if
      next
    end with
  end if
  

  
  
  '--> SUCHE-3: falls noch nichts gefunden wurde, dann alles durchsuchen
  if foundIndex = -1
      if globDumpMode=true then  monitorAdd("SUCHE-3:  ALLES DURCHSUCHEN...")
    with grid
      max = grid.rows.count
      'max=iGrid1.rows.count - 1
      
      for i=0 to max-1  step 20              '...vorwaerts - step 20
        'monitorAdd("i: "+i.toString)
        'if i>max then exit for
        row=iGrid1.rows(i)
        item=trim(row.cells(cId).value)
        if item="" then continue for
        'monitorAdd("item: "+item)
        id=cInt(item)
        'monitorAdd("id: "+id.toString)
       if id > curLineNumber then exit for
      next
      
      dim j as integer
      for j=i to 0  step-1                '...rueckwaertsgang - step -1
        'monitorAdd("222 j: "+j.toString)
        row=iGrid1.rows(j)
        item=trim(row.cells(cId).value)
        if item="" then continue for
        'monitorAdd("222 item: "+item)
        id=cInt(item)
        'monitorAdd("222 id: "+id.toString)
        if val(id) < val(curLineNumber+1) then exit for
      next
      foundIndex = row.index
    end with 
  end if

  '--> SUCHE-X: checkForIndexListUpdate

  with grid
    '...FOUND:
    
    '...CHECK FOR UPDATE
    ' ??? checkForIndexListUpdate(grid, foundIndex)
    checkForIndexListUpdate(grid, foundIndex)
    
    '...SELECT
    iGrid1.curRow=iGrid1.rows(foundIndex)

    ''dim foundCaption as string=""

    '### 
    '...EXPAND HIDDEN ITEMS
    dim newTempExpandCaption as string=""
    dim expandCaption as boolean =false
    if iGrid1.curRow.visible = false then expandCaption=true
    ''monitorAdd("iGrid1.curRow.visible = false")

    newTempExpandCaption = indexListExpandTempMode(grid, iGrid1.curRow.index, expandCaption)
    ' globTempExpandCaption = newTempExpandCaption

  
  
    if globTempExpandCaption<>"" and globTempExpandCaption <> newTempExpandCaption then
      indexListCollapsTempMode(grid, globTempExpandCaption)
    end if
    
    if newTempExpandCaption <> "" and expandCaption =true then
      globTempExpandCaption=newTempExpandCaption
    end if
    

  end with

end function




: sub checkForIndexListUpdate(grid As iGrid, rowIndex as integer)
  dim maxRows=grid.rows.count
  if rowIndex<maxRows-2 then rowIndex=rowIndex+1                 ' Eintrag n+1 ist besser geeignet fuer updateCheck

  dim row as igRow = grid.rows(rowIndex)
  dim oldLineNumber = trim(row.cells(cId).value)
  dim oldSourceLine  = row.cells(cSourceLine).value        
  'dim oldSourceLine  = row.cells(cGruppe).value
  dim curCodeLine as string = getCurCodeLineFromLineNr(cInt(oldLineNumber))
  'dim foundIndex as integer=row.index

  if oldSourceLine+vbNewLine <> curCodeLine then
    monitorAdd("UPDATE ********************(checkForIndexListUpdate)")
    '' monitorAdd("-old- "+oldSourceLine)
    '' monitorAdd("-new- "+curCodeLine)
    '' monitorAdd("UPDATE ********************(checkForIndexListUpdate)")
    
    

    'NEU...
    dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
    dim key=ActiveTabFilespec
    if globScannerData.ContainsKey(key) then
      globScannerData.remove(key)
    end if
    
    updateCodeList() 
    '??? highlightIndexList hier aufrufen ???  
  end if
  
end sub









: function getCurCodeLineFromLineNr(lineNumber as integer)as string
  dim codeEitor = ZZ.getActiveRTF.RTF
  'dim lineContent as string = codeEitor.Lines.Current.text
  'codeEitor.Lines.Current.number=codeLineNumber
  'codeEitor.goTo.line(codeLineNumber+50)

  dim cLine as ScintillaNet.Line = codeEitor.lines(lineNumber)
  return(cLine.text)
end function




'--> 
'--> C O L L A P S - E X P A N D 








function collapsExpandIgridTopic(grid As iGridEx,  selRow  As iGRow )
  on error resume next
  'msgBox(grid.name)
  '' dim bgColorNickName  As System.Drawing.Color  
  '' dim foreColorNickName  As System.Drawing.Color  
  '' dim bgColorNickName2  As System.Drawing.Color  
  '' dim foreColorNickName2  As System.Drawing.Color  
  
  iniGridColors(grid)
  dim startIndex as integer = selRow.Index
  if grid.rows(startIndex).cells(cGroupe).value<>"GG" then return false
  dim bgColor as string = "#445370"
  dim max as integer=grid.rows.count
  dim i as integer 
  dim newState as boolean = not grid.rows(startIndex+1).visible
  dim key as string=grid.rows(startIndex).cells(cNickname).value
  dim marker as string
   if newState=true then
     marker=">"
     grid.rows(startIndex).cells(cGroupeState).value=marker
     'grid.rows(startIndex).cells(cNickname).Style=igFett
     ' ? ? ? grid.rows(startIndex).cells(cNickname).Style=IgDefaultCellStyle
     
     'grid.rows(startIndex).Cells(cNickname).BackColor=ColorTranslator.FromHtml("#CD7000") 'bgColorNickName
     'grid.rows(startIndex).Cells(cNickname).ForeColor=ColorTranslator.FromHtml("#fff") 'foreColorNickName
      
     grid.rows(startIndex).Cells(cNickname).BackColor=bgColorNickName  'bgColorNickName     'ColorTranslator.FromHtml(bgColor)
     grid.rows(startIndex).Cells(cNickname).ForeColor=foreColorNickName   'ColorTranslator.FromHtml("#eee")
     '' grid.rows(startIndex).Cells(cBorder1).BackColor=bgColorHighLight  'bgColorNickName     'ColorTranslator.FromHtml(bgColor)
     '' grid.rows(startIndex).Cells(cBorder1).ForeColor=foreColorHighLight   'ColorTranslator.FromHtml("#eee")
     '' grid.rows(startIndex).Cells(cBorder2).BackColor=bgColorHighLight  'bgColorNickName     'ColorTranslator.FromHtml(bgColor)
     '' grid.rows(startIndex).Cells(cBorder2).ForeColor=foreColorHighLight   'ColorTranslator.FromHtml("#eee")

  else
     marker=""
     grid.rows(startIndex).cells(cGroupeState).value=marker
     ' ? ? ? grid.rows(startIndex).cells(cNickname).Style=IgDefaultCellStyle
     'grid.rows(startIndex).Cells(cNickname).BackColor=ColorTranslator.FromHtml("#bbb")
     grid.rows(startIndex).Cells(cNickname).BackColor=bgColorNickName
     grid.rows(startIndex).Cells(cNickname).ForeColor=foreColorNickName
     '' grid.rows(startIndex).Cells(cBorder1).BackColor=bgColorNickName  'bgColorNickName     'ColorTranslator.FromHtml(bgColor)
     '' grid.rows(startIndex).Cells(cBorder1).ForeColor=foreColorNickName   'ColorTranslator.FromHtml("#eee")
     '' grid.rows(startIndex).Cells(cBorder2).BackColor=bgColorNickName  'bgColorNickName     'ColorTranslator.FromHtml(bgColor)
     '' grid.rows(startIndex).Cells(cBorder2).ForeColor=foreColorNickName   'ColorTranslator.FromHtml("#eee")
  end if

  dim key2 as string=globCurTabFileSpec+"|||"+key
  if globCodeListState.ContainsKey(key2) then
    globCodeListState.item(key2)=marker
  else
    globCodeListState.add(key2,marker)
  end if
  dim newBorder=ColorTranslator.FromHtml("#F28500")
  if newState=true then newBorder=ColorTranslator.FromHtml("#fff")
  for i=startIndex+1 to max-1
    if grid.rows(i).cells(cGroupe).value="GG" then exit for
    grid.rows(i).visible=newState
    'grid.rows(i).Cells(cBorder2).BackColor=bgColorNickName2 'newBorder
    grid.rows(i).Cells(cBorder2).BackColor=bgColorNickName 'newBorder
    grid.rows(i).Cells(cNickname).BackColor=bgColorNickName2
    grid.rows(i).Cells(cNickname).ForeColor=foreColorNickName2
    'grid.Rows(i).ContentIndent= New TenTec.Windows.iGridLib.iGIndent(1, 0, 1, 0)
    grid.Rows(i).ContentIndent= New TenTec.Windows.iGridLib.iGIndent(0, 0, 0, 0)
    'grid.Rows(i).height= 14

  next
  '' dumpDictionary(globCodeListState)   '...fuer testzwecke
  saveDictOfStringString(globCodeListState, globIndexListStateFileSpec)
  return true
end function




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





'--> 
'--> E D I T O R - T A B S 







:sub scNet_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
  'monitorAdd(e.keycode.toString)
  if globShowHistoryState =true then historyAutoHide()

  if e.keyCode=13 then
    'monitorAdd("KeyDown-ENTER")
    if e.control=true
       skipNextKeyPress=true
       monitorAdd("KeyDown(contr.ENTER")
    end if
  end if
end sub






:sub scNet_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
  'monitorAdd(e.KeyChar)
  if skipNextKeyPress =true then
    skipNextKeyPress=false
    e.handled=true
    monitorAdd("onMyTextArea_KeyPress: CONTROL-ENTER")
    'executeDirektFenster()
    'geplant fuer shortCuts - coming soon
  end if
end sub




'...war mal mouseDown
Sub scNet_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
'-->ERROR...
exit sub


  '' if isControl() = true
  ''   zz.ideHelper.exec("Window.Reflection","ideHelper","ideHelper")
  '' end if
  ''       

  'Handles scNet.MouseDown
  'monitorAdd("scNet_MouseDown")
  
  'msgBox(globShowHistoryAutoHide.toString)
  if globShowHistoryState =true then historyAutoHide()

  Dim wordStart, wordEnd As Integer
  Dim line = scNet.Selection.Range.StartingLine
  Dim selStart As Integer = line.SelectionStartPosition - line.StartPosition
  if selStart < 1 then exit sub
  
  For wordStart = selStart To 0 Step -1
    If Char.IsLetterOrDigit(line.Text.Substring(wordStart - 1, 1)) = False Then 
      if line.Text.Substring(wordStart - 1, 1)<>"_" then Exit For
    end if
  Next
  
  For wordEnd = selStart To line.Text.Length - 2
    If Char.IsLetterOrDigit(line.Text.Substring(wordEnd, 1)) = False Then 
      if line.Text.Substring(wordEnd, 1)<>"_" then Exit For
    end if
  Next

  Dim wordUnderCursor = line.Text.Substring(wordStart, wordEnd - wordStart)
  'If wordUnderCursor.Length < 2 Then Exit Sub
  'navHelpByKeyword(wordUnderCursor.ToLower)

  static lastWord as string
  if lastWord<>wordUnderCursor then
    lastWord=wordUnderCursor
    '''monitorAdd("curWord(Scintilla): "+wordUnderCursor)
    'monitorAdd("curWord(Scintilla): "+wordUnderCursor.ToLower)
   '''monitorAdd(isControl.toString)
  end if 
  if isControl() = true
    navigateForCodeItem(wordUnderCursor)
  end if
  if isShift() = true
    msgBox("isShift")
  end if
End Sub



'--> ...


: sub navigateForCodeItem(findText as string)
  on error resume next
  'monitorAdd("findText: "+findText)
  dim indexListRow  As iGRow
  indexListRow = findInCodeList(findText)
  if indexListRow is nothing then
    '...navigate Reflector
    zz.ideHelper.exec("Window.Reflection","ZZ","")

    '??? ...such globVars
  else   
    navigateCodeLink(iGrid1,indexListRow)
    ''Dim line = scNet.Selection.Range.length=0
    '' 'zz.doevents
    '' scNet.Selection.length=0
    '' dim lineInt as integer = scNet.Selection.range.startingLine.number
    '' highlightExecutingLine(lineInt)
  end if
end sub






sub codeEditorNavigate(fileSpec as string, lineNumber as integer)
  on error resume next

  '... der wechselt den tab
  dim tab = ZZ.IDEHelper.NavigateFile(fileSpec)
  
  
  dim codeEitor = ZZ.getActiveRTF.RTF
  codeEitor.goTo.line(lineNumber+50)
  codeEitor.goTo.line(lineNumber-10)
  codeEitor.goTo.line(lineNumber)
  highlightExecutingLine(lineNumber)
  codeEitor.focus()
  err.number=0


end sub





'--> 
'--> N A V I G A T E




sub switchToCurrentEditor()
  setMonitorRef()
  :RemoveHandler scNet.MouseUp, AddressOf scNet_MouseUp :err.number=0
  :RemoveHandler scNet.KeyDown,    AddressOf scNet_KeyDown :err.number=0
  :RemoveHandler scNet.KeyPress,   AddressOf scNet_KeyPress :err.number=0
  scNet = ZZ.getActiveRTF.RTF
  sc1= ZZ.getActiveRTF.RTF
  AddHandler scNet.MouseUp,     AddressOf scNet_MouseUp :err.number=0
  AddHandler scNet.KeyDown,     AddressOf scNet_KeyDown
  AddHandler scNet.KeyPress,    AddressOf scNet_KeyPress
end sub


'--> ...


sub navigateCodeLink3(codeLinkString as string)  '... das ist der neue
  on error resume next
  dim codeLinkDATA() as string = split(codeLinkString, globCodeLinkSplitter)
  
  '' OUT(ccReserve1)         = ""
  '' OUT(ccVersion)          = ""
  '' OUT(ccTimeStamp)        = ""
  '' OUT(ccReserve2)         = ""
  '' OUT(ccNicName)          = ""
  '' OUT(ccType)             = ""
  '' OUT(ccApp)              = ""


  '' OUT(ccFileSpec)         = ""
  '' OUT(ccReserve3)         = ""
  '' OUT(ccTopic)            = ""
  '' OUT(ccTopicListIndex)   = ""
  '' OUT(ccTopicLineNr)      = ""
  '' OUT(ccTopicSourceLine)  = ""
  '' OUT(ccDeltaLine)        = ""
  '' OUT(ccColPos)           = ""
  '' OUT(ccTargetSourceLine) = ""
  '' OUT(ccReserve4)         = ""

  dim fileSpec as string          = codeLinkDATA(ccFileSpec)
  dim targetLineNumber as integer = cInt(codeLinkDATA(ccTargetSourceLine)) '... ein vorlaeufiger wert

  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()

  '...falls tabwechsel noetig ist
  if not ActiveTabFilespec.contains(fileSpec) then
     dim tab = ZZ.IDEHelper.NavigateFile(fileSpec)
     updateCodeList()
     dim name as string=My.Computer.FileSystem.GetName(ActiveTabFilespec)
     panelRef.element("lab_title").text=name
     switchToCurrentEditor()
  end if
  
  '...eintrag in codeList suchen
  dim topic=codeLinkDATA(ccTopic)
  dim row As iGRow = findInCodeList(topic)

  if row is Nothing then
  '--> [_] targetLineNummer fehlt noch
    
    monitorAdd("Topic '"+topic+"' ... nicht gefunden")
    ' msgBox("Topic '"+topic+"' ... nicht gefunden")
    
    codeEditorNavigate(fileSpec, targetLineNumber)  '... kann reichlich daneben liegen
    '...loeschen vorschlagen
    '...oder mittels suchfunktion versuchen ???
    exit sub
    'alternative suchmethode ergaenzen
  end if

  dim topicLineNumber as integer  = cInt(trim(row.cells(cId).value))
  'dim topicLineNumber as integer  = cInt(codeLinkDATA(ccTopicLineNr))
  dim deltaLine as integer        = cInt(codeLinkDATA(ccDeltaLine))
  targetLineNumber                = topicLineNumber+deltaLine  '...jetzt relativ zu topic
  
  ''' monitorAdd(topicLineNumber.toString)
  ''' monitorAdd(deltaLine.toString)
  ''' monitorAdd(targetLineNumber.toString)
  
  '--> [_] historyAddItem(codeLinkItem)
  codeEditorNavigate(fileSpec, targetLineNumber)
End Sub 



'--> ...





: function findInCodeList(findText as string)  As iGRow
  '!!! ... validate fehlt noch
  on error resume next
    if globDumpMode=true then  monitorAdd("findText: "+findText)
  dim grid as iGrid=iGrid1
  dim such as string=trim(findText.toLower)
  dim ir as iGRow
  Dim max = grid.Rows.Count - 1
  dim item as string
  dim item2 as string
  For i As Integer = 0 To max
    ir = Grid.Rows(i)
    item= trim(ir.Cells(cCodeFinder).Value.ToString)
    'monitorAdd(item)
    if such=item.toLower then
        if globDumpMode=true then  monitorAdd("TREFFER(findInCodeList): "+i.toString)
      return ir
      'return i
    end if
  Next
  return nothing
end function




function findListIndexFromLineNumber(codeEditor As ScintillaNet.Scintilla, optional lineNumberSuch as integer=-1)as integer
  on error resume next
  'dim codeEditor = ZZ.getActiveRTF.RTF
  if lineNumberSuch<0 then
    lineNumberSuch=codeEditor.Lines.Current.number
  end if
  if globDumpMode=true then  monitorAdd("findText: "+lineNumberSuch.toString)
  lineNumberSuch=lineNumberSuch+1                  ' WICHTIG !!!
 
  dim grid as iGrid=iGrid1
  dim ir as iGRow
  Dim max = grid.Rows.Count - 1
  dim i as integer
  dim item as string
  dim lineNumber as integer
  dim item2 as string
  For i  = max-1  To 0 step -1
    ir = Grid.Rows(i)
    item= trim(ir.Cells(cId).Value.ToString)
    lineNumber=cInt(item)
    if lineNumber < lineNumberSuch and trim(ir.Cells(cCodeFinder).Value) <>""then
        if globDumpMode=true then  monitorAdd("TREFFER(findListIndexFromLineNumber): "+i.toString)
        if globDumpMode=true then  monitorAdd("TREFFER: "+ir.Cells(cCodeFinder).Value)
      return i
      'return i
    end if
  Next
  return -1
end function





'--> 
'--> C O D E - L I N K 




sub testCodeLinkFormat()
  dim max as integer = cCodeLinkName.length
  dim OUT(max) as string
  OUT(ccReserve1)         = "0"
  OUT(ccVersion)          = "1"
  OUT(ccTimeStamp)        = "22"
  OUT(ccReserve2)         = "333"
  OUT(ccNicName)          = "4444"
  OUT(ccType)             = "55555"
  OUT(ccApp)              = "666666"
  OUT(ccFileSpec)         = "7777777"
  OUT(ccReserve3)         = "88888888"
  OUT(ccTopic)            = "999999999" 
  OUT(ccTopicListIndex)   = "0000000000" 
  OUT(ccTopicLineNr)      = "11111111111"
  OUT(ccTopicSourceLine)  = "222222222222"
  OUT(ccDeltaLine)        = "3333333333333"
  OUT(ccColPos)           = "44444444444444"
  OUT(ccTargetSourceLine) = "555555555555555"
  OUT(ccReserve4)         = "6666666666666666"

  dim codeLink as string = join(OUT,globCodeLinkSplitter)
  if globDumpMode=true then  monitorAdd("NEU(cl): testCodeLinkFormat")
  if globDumpMode=true then  dumpCodeLinkData(codeLink)
  
 
  '' OUT(ccReserve1)         = ""
  '' OUT(ccVersion)          = ""
  '' OUT(ccTimeStamp)        = ""
  '' OUT(ccReserve2)         = ""
  '' OUT(ccNicName)          = ""
  '' OUT(ccType)             = ""
  '' OUT(ccApp)              = ""
  '' OUT(ccFileSpec)         = ""
  '' OUT(ccReserve3)         = ""
  '' OUT(ccTopic)            = ""
  '' OUT(ccTopicListIndex)   = ""
  '' OUT(ccTopicLineNr)      = ""
  '' OUT(ccTopicSourceLine)  = ""
  '' OUT(ccDeltaLine)        = ""
  '' OUT(ccColPos)           = ""
  '' OUT(ccTargetSourceLine) = ""
  '' OUT(ccReserve4)         = ""


end sub



function dumpCodeLinkData(codeLink as string)
  dim i as integer
  dim max as integer=cCodeLinkName.length
  dim FIELDS() as string = split(codeLink,globCodeLinkSplitter)
  dim item as string
  dim caption as string
  dim OUT(max)as string
  for i=0 to max-1
    caption=esStringFormat(cCodeLinkName(i),20,"_")
    item=FIELDS(i)
    OUT(i)=caption+": "+item
  next
  dim RESULT=join(OUT,vbCrLf)
  monitorAdd(RESULT)
  return RESULT
end function



function esStringFormat(str as string, len as integer, optional fill as string = " ") as string
  'pruefen auf dbNull geht hier nicht
  if String.isNullOrEmpty(str) Then str=""
  return strings.right(StrDup(math.max(0,len-str.Length), fill) + str,len)

  '' 
  '' 
  '' Dim str = "abc"
  '' TRACE    Space(20-len(str)) + str
  '' 
  '' 
  '' 
  '' Dim str = "abc"
  '' TRACE    StrDup(20-len(str), "_") + str
  '' 
  '' 
  '' TRACE RSet(str, 20)
  '' 
  '' 
  '' Dim str = "abc"
  '' TRACE    StrDup(math.max(0,20-len(str)), "_") + str
  '' 
  '' 
  '' Dim str = "abc"
  '' TRACE    right(StrDup(math.max(0,20-len(str)), "_") + str,20)
  '' 

end function


'--> --------------------------------------------






























function getCodeLink()
  on error resume next
  '--> ...evt. mehrere Eintraege ??? ...oder auf naechste sub/func (checkType)
  '... cCodeFinder auswerten ???


  '--> ... Eintrag in codeList suchen
  dim codeEditor = ZZ.getActiveRTF.RTF
  dim lineNumber as integer=codeEditor.Lines.Current.number
  dim codeListIndex as integer = findListIndexFromLineNumber(codeEditor, lineNumber)
  
  '--> ...vorher validate ???
  
  
  '--> ...deltaLines ermitteln
  
  dim grid as iGrid=iGrid1
  dim ir as iGRow = grid.rows(codeListIndex)
  dim topic as string             = ir.Cells(cCodeFinder).Value
  dim topicLineNumber as integer  = cInt(trim(ir.Cells(cId).Value.ToString))
  dim deltaLine as integer        = lineNumber - topicLineNumber
  dim topicSourceLine as string   = ir.Cells(cSourceLine).Value
  'dim fileSpec as string          = ir.Cells(cFileSpec).Value
  dim lineContent as string = codeEditor.Lines.Current.text
  dim startPos as integer = codeEditor.Lines.Current.SelectionStartPosition
  dim column as integer=codeEditor.GetColumn(startPos)
  
  dim startPos2 as integer = codeEditor.Selection.Start
  dim endPos as integer  = codeEditor.Selection.End
  dim linePos as integer=codeEditor.CaretInfo.Position()

  dim ActiveTabFilespec    = ZZ.getActiveTabFilespec()
  'dim fileSpec as string   = mid(ActiveTabFilespec,6)
  dim fileSpec as string   = ActiveTabFilespec
  
  '' monitorAdd("1: "+lineContent)
  '' monitorAdd("2: "+startPos.toString)
  '' monitorAdd("3: "+lineNumber.toString)
  '' monitorAdd("4: "+column.toString)


  '--> Ausgabe sauber als codeLink organisieren

  dim prefixNicname="["+deltaLine.toString+"] "
  if deltaLine=0 then prefixNicname=""

  dim max as integer = cCodeLinkName.length
  dim OUT(max) as string
  OUT(ccReserve1)         = "xxx-1"
  OUT(ccVersion)          = "cLink_2.0"
  OUT(ccTimeStamp)        = now.tostring
  OUT(ccReserve2)         = "xxx-2"
  OUT(ccType)             = "CODE-LINK-SCRIPT-IDE"
  OUT(ccApp)              = "SCRIPT-IDE-4"
  OUT(ccFileSpec)         = fileSpec
  OUT(ccReserve3)         = "(3)-------------------------------------"
  OUT(ccNicName)          = prefixNicname+trim(topic)   ' was soll genommen werden
  OUT(ccTopic)            = topic
  OUT(ccTopicListIndex)   = codeListIndex.toString
  OUT(ccTopicLineNr)      = topicLineNumber.toString
  OUT(ccTopicSourceLine)  = topicSourceLine
  OUT(ccDeltaLine)        = deltaLine.toString
  OUT(ccColPos)           = column.toString
  OUT(ccTargetSourceLine) = mid(lineContent, 1, lineContent.length - 2)
  OUT(ccReserve4)         = "xxx-4"

  '...Kommentare geeignet ausgeben
  dim nicName2 as string = trim(OUT(ccTargetSourceLine))
  'msgbox(nicName2)
  monitorAdd("nicName2",nicName2)
  if nicName2.startsWith("'") then
    monitorAdd("TREFFER",nicName2)
    OUT(ccNicName)=nicName2
  end if

  dim codeLink as string = join(OUT,globCodeLinkSplitter)
  if globDumpMode=true then   monitorAdd("NEU(cl): getCodeLink()")
  if globDumpMode=true then  dumpCodeLinkData(codeLink)
    
  return codeLink
end function



'--> ----------------------------------------------



'--> 
'--> H I S T O R Y 

Sub IGrid2_CellMouseDown(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseDownEventArgs) Handles IGrid2.CellMouseDown   
  dim lastSel As iGRow = sender.curRow
  sender.curRow=sender.rows(e.RowIndex)
  sender.curRow.selected=true
  lastSel.selected=false
  printLine  6, "IGrid2_CellMouseDown - row", e.RowIndex
  printLine  7, "IGrid2_CellMouseDown - col", e.ColIndex

  if e.Button= Windows.Forms.MouseButtons.Right then
    printLine  8, "IGrid2_CellMouseDown - col", "rClick"
    monitorAdd("IGrid2_CellMouseDown - col", "rClick")
    historyAddItemBeforeSelected()   
  else
    printLine  8, "IGrid2_CellMouseDown - col", "NORMAL"
    historyNavigate(sender,sender.curRow, false)
 end if
End Sub



Sub cmdHistoryAdd(e)
  dim codeLinkItem=getCodeLink()
  gridAddCodeLinkItem(iGrid2, codeLinkItem, -2) '... ans ende anfuegen
End Sub


sub historyAddItemBeforeSelected()
  dim codeLinkItem=getCodeLink()
  gridAddCodeLinkItem(iGrid2, codeLinkItem, -1) '... ans aktueller Position anfuegen
end sub




sub historyNavigate(grid As iGrid,  selRow  As iGRow, optional addHistory as boolean = false)
  globSkipNextHistoryAction=true
  '--> '!!! NEU: bookmarks gehoert aus dem namen raus
   navigateFromGrid(grid,selRow,false)
   
  globSkipNextHistoryAction=false
end sub


'--> ...


sub gridAddCodeLinkItem(grid as iGridEx, codeLinkItem as string, optional rowIndex as integer=-1)
  
  'dim codeLink=getCodeLink()
  dim codeLinkDATA() as string = split(codeLinkItem,globCodeLinkSplitter )

  dim nicName as string          = codeLinkDATA(ccNicName)
  dim fileSpec as string          = codeLinkDATA(ccFileSpec)
  dim topicLineNumber as integer = cInt(codeLinkDATA(ccTopicLineNr))
  dim deltaLineNumber as integer = cInt(codeLinkDATA(ccDeltaLine))
  dim lineNumber as string    =    (topicLineNumber+deltaLineNumber).toString
  dim codeLinkString as string    = codeLinkItem
  '' dim xxx as string    = codeLinkDATA()

  '' OUT(ccReserve1)         = ""
  '' OUT(ccVersion)          = ""
  '' OUT(ccTimeStamp)        = ""
  '' OUT(ccReserve2)         = ""
  '' OUT(ccNicName)          = ""
  '' OUT(ccType)             = ""
  '' OUT(ccApp)              = ""
  '' OUT(ccFileSpec)         = ""
  '' OUT(ccReserve3)         = ""
  '' OUT(ccTopic)            = ""
  '' OUT(ccTopicListIndex)   = ""
  '' OUT(ccTopicLineNr)      = ""
  '' OUT(ccTopicSourceLine)  = ""
  '' OUT(ccDeltaLine)        = ""
  '' OUT(ccColPos)           = ""
  '' OUT(ccTargetSourceLine) = ""
  '' OUT(ccReserve4)         = ""

  ' dim grid as iGrid=iGrid3
  if rowIndex = -1  then              'EinfuegePosition = selected
    rowIndex = grid.curRow.index
  end if
  if rowIndex = -2  then              'EinfuegePosition = ans Ende
    rowIndex = grid.rows.count
  end if
  
  dim ir as iGRow
  
  ir = grid.rows.insert(rowIndex)     ' neue Zeile einfuegen
  'ir = grid.rows.add()
  
  with ir
  .cells(cNickname).value  = nicName
  .cells(cId).value        = lineNumber 
  .cells(cFileSpec).value  = fileSpec   
  .cells(cCodeLink).value  = codeLinkString   
  '.cells(cNickname).backColor = ColorTranslator.FromHtml("#7A8599")
  '.cells(cNickname).foreColor = ColorTranslator.FromHtml("#fff")
  .cells(cNickname).backColor = ColorTranslator.FromHtml("#364359")
  .cells(cNickname).foreColor = ColorTranslator.FromHtml("#ccc")
  .ContentIndent= New TenTec.Windows.iGridLib.iGIndent(1, 0, 1, 0)
  end with
  
  
  ir.ensureVisible()
  grid.curRow=ir

  
  '' ... hier oder nicht hier ??? saveGridData(grid)


end sub

'--> UMSTELLUNG...'???


sub historyAddItem(codeLinkItem() as string, optional rowIndex as integer=-1)
  '... hier gehen bisher die Befehle rein
  if globSkipNextHistoryAction=true then exit sub

'####################
  
  dim grid as iGrid=iGrid2
  'dim curRow as igRow=grid.curRow
  '' dim codeLinkItem() as string=getCodeLinkItem2(grid,curRow)
  '' historyAddItem(codeLinkItem)

  dim codeLinkItem2=getCodeLink()
  gridAddCodeLinkItem(grid, codeLinkItem2, -2) '... ans ende anfuegen
  
  '...save
  cmdHistorySave()

exit sub
'######################

'' 
''   dim grid as iGrid=iGrid2
''   dim ir as iGRow 
''   if rowIndex < 0 then
''     ir = grid.rows.add()
''   else
''     ir = grid.rows.insert(rowIndex)
''   end if
''   
''   InsertCodeLinkItem(ir,codeLinkItem)
''   ir.ensureVisible()
''   grid.curRow=ir
''   
''   '...save
''   cmdHistorySave()
end sub


sub InsertCodeLinkItem(Row As iGRow,codeLinkItem() as string)
  dim i as integer
  dim max as integer=codeLinkItem.length
  for i =0 to max-1
    row.cells(i).value=codeLinkItem(i)
  next
  row.cells(cNickname).backColor = ColorTranslator.FromHtml("#364359")
  row.cells(cNickname).foreColor = ColorTranslator.FromHtml("#ccc")
  row.ContentIndent= New TenTec.Windows.iGridLib.iGIndent(1, 0, 1, 0)
end sub






'-->...

Sub cmdHistoryRemove(e)  ' ...
  dim grid as iGrid=iGrid2
  dim ir as igRow=grid.curRow
  dim curRowIndex as integer=ir.index
  'msgBox(curRowIndex)
  grid.rows.removeAt(curRowIndex)
  if curRowIndex<grid.rows.count then
    grid.rows(curRowIndex).selected=true
  else
    grid.rows(curRowIndex-1).selected=true
  end if
  'grid.visible=false
  'grid.visible=true
  historyNavigate(grid, grid.curRow, false)
End Sub





Sub cmdHistorySave(optional e as object = nothing) 

''ERROR...
exit sub


  '??? nur temporaer speichern 
  on error resume next
  Dim data As String
  dim grid as iGrid=iGrid2
  
  Dim max as integer = Grid.Rows.Count - 1
  dim i as integer
  dim start as integer=0
  if max > 40 then start=max-40
  Dim out(max) As String
  dim counter as integer =0
  For i = start To max
    out(counter) = JoinIGridLine(Grid.Rows(i), vbTab)
    counter = counter+1
  Next
  redim preserve out(counter)
  data = Join(out, vbNewLine)
  msgBox
  glob.para("historyData")=data
  glob.saveParaFile()
End Sub




Sub cmdHistoryRead(e)  ' ...
  dim grid as iGrid=iGrid2
  Dim data As String
  data=glob.para("historyData")
  IGrid_put(grid, data)
  dim i as integer
  dim max as integer=grid.rows.count
  dim row as igRow
  for i=0 to max-2
    row = grid.rows(i)
    row.cells(cNickname).backColor = ColorTranslator.FromHtml("#364359")
    row.cells(cNickname).foreColor = ColorTranslator.FromHtml("#ccc")
    row.ContentIndent= New TenTec.Windows.iGridLib.iGIndent(1, 0, 1, 0)

  next
  row.selected=true
  grid.rows.removeAt(max-1)
  
End Sub



Sub cmdHistoryClear(e)  ' ...
  dim grid as iGrid=iGrid2
  grid.rows.clear
End Sub


'-->...

Sub cmdHistoryDeltaIndex(e)  ' ...
  '... der kommt von der FORM-1
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim deltaIndex as integer = val(name(2))
  monitorAdd("deltaIndex", deltaIndex.toString)
  if deltaIndex=2 then deltaIndex=-1
   historyNavigateDeltaIndex(deltaIndex)
End Sub


sub historyNavigateDeltaIndex(deltaIndex as integer)
  '... ROT - GRueN
  
  '--> ...autoShowHistoryForm
  dim id ="ToolBar|##|tbScriptWin|##|es_codeList.history"
  dim toolWindow=ZZ.getDockPanelRef(id)
  
  dim winState as boolean=toolWindow.visible
  if winState = false then
    toolWindow.show
    globShowHistoryTimeStamp = getTime()
    globShowHistoryState=true
  end if
  globShowHistoryTimeStamp = getTime()
  
  dim ig as iGridEx = iGrid2
  dim index as integer=iGrid2.curRow.index
  'monitorAdd("index",index.toString)
  dim newIndex as integer=index+deltaIndex
  if newIndex < 0 then newIndex=0
  'if newIndex>(ig.rows.count - 1) then newIndex =ig.rows.count -1
  'monitorAdd("newIndex",newIndex.toString)
  'monitorAdd("ig.rows.count",ig.rows.count.toString)
  if newIndex>(ig.rows.count - 1) then 
    highlightIndexList() 
    exit sub
  end if
  'monitorAdd("newIndex",newIndex)
  ig.Rows(newIndex).selected=true '=ig.rows(newIndex)
  historyNavigate(ig, ig.curRow, false)
  highlightIndexList()
end sub







Sub cmdHistory(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! 'cmdHistory': ...in Arbeit")
  'msgBox("'cmdHistory': " + "...Coming soon ")
  '...
End Sub



Sub cmdHistoryNavigate(e)  ' ...
  '' Dim name() = Split(e.Sender.Name+"_", "_")
  '' dim action=name(1)
  '' dim index as integer = val(name(2))
  dim curRow=iGrid2.curRow
  historyNavigate(iGrid2, curRow,false)
  
End Sub







'--> 
'--> B O O K M A R K S  



Sub cmdNavBookmarks(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  dim stringIndex=name(2)
  '...
  statustext("!!! 'cmdNavBookmarks': ...in Arbeit")
  ''msgBox("'cmdNavBookmarks': " + "...Coming soon ")
  '...

  dim grid as iGridEx=IGrid3
  '' .fileSpec=globIndexListBookmarksFileSpec+"1"+".txt"
  grid.fileSpec=globIndexListBookmarksFileSpec+"_"+stringIndex+".txt"
  
  if index=7 then   '' ... für meine bisherige Liste
    grid.fileSpec=globIndexListBookmarksFileSpec+".txt"
  end if
  
  cmdBookmarkRead(nothing)
  
End Sub



Private Sub IGrid3_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles IGrid3.DoubleClick
  monitorAdd("IGrid3_DoubleClick", "xxx")
  ''msgBox("...vorgesehen für edit")
  IGrid3.ReadOnly = false
  'IGrid3.
End Sub


Sub IGrid3_CellMouseDown(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseDownEventArgs) Handles IGrid3.CellMouseDown   
'Sub IGrid3_CellMouseUp(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseUpEventArgs) Handles IGrid3.CellMouseUp   '
  dim lastSel As iGRow = sender.curRow
  sender.curRow=sender.rows(e.RowIndex)
  sender.curRow.selected=true
  lastSel.selected=false
  'IGrid3.ReadOnly = true
  
  printLine  6, "IGrid3_CellMouseDown - row", e.RowIndex
  printLine  7, "IGrid3_CellMouseDown - col", e.ColIndex

  if e.Button= Windows.Forms.MouseButtons.Right then
    printLine  8, "IGrid1_CellMouseDown - col", "rClick"
    monitorAdd("IGrid1_CellMouseDown - col", "rClick")
    
    ''veraltet:
    ''insertCodeLink()
    
    'NEU
    ' ??? bookmarkInsertCodeLink()
  else
    printLine  8, "IGrid1_CellMouseDown - col", "NORMAL"
    
    dim isCaption as boolean = collapsExpandIgridTopic(sender,  sender.curRow)
    if isCaption=false then 
      'navigateCodeLink(sender,sender.curRow)
      navigateFromGrid(sender,sender.curRow, true)
    end if
    '' collapsExpandIgridTopic(sender,  sender.curRow)
    '' navigateCodeLink(sender,sender.curRow)
 end if

  '' dim lastSel As iGRow = sender.curRow
  '' sender.curRow=sender.rows(e.RowIndex)
  '' sender.curRow.selected=true
  '' lastSel.selected=false
  '' dim isCaption as boolean =false
  '' 
  '' printLine  6, "IGrid1_CellMouseDown - row", e.RowIndex
  '' printLine  7, "IGrid1_CellMouseDown - col", e.ColIndex
  '' if e.Button= Windows.Forms.MouseButtons.Right then
  ''   printLine  8, "IGrid1_CellMouseDown - col", "rClick"
  ''   monitorAdd("IGrid1_CellMouseDown - col", "rClick")
  '' else
  ''   printLine  8, "IGrid1_CellMouseDown - col", "NORMAL"
  ''   isCaption = collapsExpandIgridTopic(sender,  sender.curRow)
  ''   if isCaption=false then navigateCodeLink(sender,sender.curRow)
  '' end if
  '' 



End Sub




''sub bookmarksNavigate(grid As iGrid,  selRow  As iGRow, optional addHistory as boolean=false ) 
sub navigateFromGrid(grid As iGrid,  selRow  As iGRow, optional addHistory as boolean=false )
  'printLine  8, "historyNavigate - col", "...."
  
  ' ??? globSkipNextHistoryAction=true
  
  dim codeLinkString as string = trim(selRow.cells(cCodeLink).value)
  if codeLinkString<>"" then
  
  if globDumpMode=true then  monitorAdd("NEU(cl): "+codeLinkString)
  if globDumpMode=true then  dumpCodeLinkData(codeLinkString)
  

    '...navigate(neu)
    navigateCodeLink3(codeLinkString)
    globSkipNextHistoryAction=false
    dim DATA() as string = split(codeLinkString,globCodeLinkSplitter)
    if addHistory = true then historyAddItem(DATA)  '!!! der verwendet den nicht
    globSkipNextHistoryAction=false

    exit sub
  end if
  
  ''msgBox("veraltet: navigateCodeLink2")
  
  monitorAdd("veraltet: navigateCodeLink2")
  '' 'collapsExpandIgridTopic(grid, selRow)
  '' navigateCodeLink2(grid, selRow)
  '' globSkipNextHistoryAction=false
end sub


'--> ...





Sub cmdBookmarkAddItem(e)  ' ...
  bookmarkInsertCodeLink()
End Sub



sub bookmarkInsertCodeLink
  '!!! insertBefore fehlt noch
  
  dim codeLink=getCodeLink()
  dim codeLinkDATA() as string = split(codeLink,globCodeLinkSplitter )

  dim nicName as string          = "   "+codeLinkDATA(ccNicName)
  dim fileSpec as string          = codeLinkDATA(ccFileSpec)
  dim topicLineNumber as integer = cInt(codeLinkDATA(ccTopicLineNr))
  dim deltaLineNumber as integer = cInt(codeLinkDATA(ccDeltaLine))
  dim lineNumber as string    =    (topicLineNumber+deltaLineNumber).toString
  dim codeLinkString as string    = codeLink
  '' dim xxx as string    = codeLinkDATA()

  '' OUT(ccReserve1)         = ""
  '' OUT(ccVersion)          = ""
  '' OUT(ccTimeStamp)        = ""
  '' OUT(ccReserve2)         = ""
  '' OUT(ccNicName)          = ""
  '' OUT(ccType)             = ""
  '' OUT(ccApp)              = ""
  '' OUT(ccFileSpec)         = ""
  '' OUT(ccReserve3)         = ""
  '' OUT(ccTopic)            = ""
  '' OUT(ccTopicListIndex)   = ""
  '' OUT(ccTopicLineNr)      = ""
  '' OUT(ccTopicSourceLine)  = ""
  '' OUT(ccDeltaLine)        = ""
  '' OUT(ccColPos)           = ""
  '' OUT(ccTargetSourceLine) = ""
  '' OUT(ccReserve4)         = ""

  dim grid as iGrid=iGrid3
  dim rowIndex as integer=grid.curRow.index
  dim ir as iGRow
  
  ir = grid.rows.insert(rowIndex+1)
  ir.selected=true
  'ir = grid.rows.add()
  
  with ir
  .cells(cNickname).value  = nicName
  .cells(cId).value        = lineNumber 
  .cells(cFileSpec).value  = fileSpec   
  .cells(cCodeLink).value  = codeLinkString   
  '.cells(cNickname).backColor = ColorTranslator.FromHtml("#7A8599")
  '.cells(cNickname).foreColor = ColorTranslator.FromHtml("#fff")
  .cells(cNickname).backColor = ColorTranslator.FromHtml("#C5CBE5")
  .cells(cNickname).foreColor = ColorTranslator.FromHtml("#000")
  .ContentIndent= New TenTec.Windows.iGridLib.iGIndent(1, 0, 1, 0)
  end with
  'saveGridData(grid)
  cmdBookmarkSave(nothing) 
end sub


'--> ...




Sub cmdBookmarkAddTitle(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("...'cmdBookmarkAddtitle': ...in Arbeit")
  'msgBox("'cmdBookmarkAddtitle': " + "...Coming soon ")
  '...
    
  dim codeLink=getCodeLink()
  dim codeLinkDATA() as string = split(codeLink,globCodeLinkSplitter )

  dim nicName as string          = "aaaaaaaaaaa" 'codeLinkDATA(ccNicName)
  dim fileSpec as string          = codeLinkDATA(ccFileSpec)
  dim topicLineNumber as integer = cInt(codeLinkDATA(ccTopicLineNr))
  dim deltaLineNumber as integer = cInt(codeLinkDATA(ccDeltaLine))
  dim lineNumber as string    =    (topicLineNumber+deltaLineNumber).toString
  dim codeLinkString as string    = codeLink
  '' dim xxx as string    = codeLinkDATA()

  '' OUT(ccReserve1)         = ""
  '' OUT(ccVersion)          = ""
  '' OUT(ccTimeStamp)        = ""
  '' OUT(ccReserve2)         = ""
  '' OUT(ccNicName)          = ""
  '' OUT(ccType)             = ""
  '' OUT(ccApp)              = ""
  '' OUT(ccFileSpec)         = ""
  '' OUT(ccReserve3)         = ""
  '' OUT(ccTopic)            = ""
  '' OUT(ccTopicListIndex)   = ""
  '' OUT(ccTopicLineNr)      = ""
  '' OUT(ccTopicSourceLine)  = ""
  '' OUT(ccDeltaLine)        = ""
  '' OUT(ccColPos)           = ""
  '' OUT(ccTargetSourceLine) = ""
  '' OUT(ccReserve4)         = ""

  dim grid as iGrid=iGrid3
  dim rowIndex as integer=grid.curRow.index
  dim ir as iGRow
  
  ir = grid.rows.insert(rowIndex)
  'ir = grid.rows.add()
  
  with ir
  .cells(cNickname).value  = nicName
  .cells(cGroupe).value  = "GG"
  '' .cells(cId).value        = lineNumber 
  '' .cells(cFileSpec).value  = fileSpec   
  '' .cells(cCodeLink).value  = codeLinkString   
  '' '.cells(cNickname).backColor = ColorTranslator.FromHtml("#7A8599")
  '' '.cells(cNickname).foreColor = ColorTranslator.FromHtml("#fff")
  .cells(cNickname).backColor = ColorTranslator.FromHtml("#C5CBE5")
  .cells(cNickname).foreColor = ColorTranslator.FromHtml("#000")
  .ContentIndent= New TenTec.Windows.iGridLib.iGIndent(1, 0, 1, 0)
  end with
  'saveGridData(grid)
  
  'cmdBookmarkSave(nothing) 

End Sub


Sub cmdBookmarkDelItem(e)  ' ...
  dim grid as iGrid=iGrid3
  dim ir as igRow=grid.curRow
  dim curRowIndex as integer=ir.index
  'msgBox(curRowIndex)
  grid.rows.removeAt(curRowIndex)
  if curRowIndex<grid.rows.count then
    grid.rows(curRowIndex).selected=true
  else
    grid.rows(curRowIndex-1).selected=true
  end if
  grid.visible=false
  grid.visible=true
  saveGridData(grid)
End Sub





Sub cmdBookmarkClear(e)  ' ...
  dim grid as iGrid=iGrid3
  grid.rows.clear
End Sub

'--> 
'--> R E A D - S A V E 


Sub cmdBookmarkSave(e)  ' ...
  dim grid as iGridEx=iGrid3
  'saveGridData(grid)
  
  dim fileName=getDizPlus+grid.fileSpec  
  Dim content As String
  IGrid_get (grid, content)
  Dim myPath as string = glob_defaultPath+"\backupBookmarks\"
  Directory.CreateDirectory(myPath)
  
  Dim myDate As Date = Now()
  dim fileName2=myDate.ToString("yyyy-MM-dd__HH_mm__")+fileName
  printLine 2, filename2, filename2

  ZZ.filePutContents(myPath+fileName2, Content)
  saveWebFile(Content,fileName)

  'IO.File.WriteAllLines(fileSpec, lines)
  'IO.File.WriteAllText(fileSpec, fileName)

End Sub


sub saveWebFile(strContent as string, strFileName as string, optional strAppName as string="") 'speichern
  if strAppName="" then strAppName=globWebAppName
  printLine 11, "saveWebFile", strFileName
  dim errMes as string
  '' Igrid_get(IGrid1, strContent, vbNewLine, "|")
  errMes=ZZ.TwSaveFile(strAppName, strFileName, strContent)
  if Not errMes.startsWith("OK:") then msgBox(errMes)
  'msgBox(errMes)
  'ref.element("txt_outMonitor").text=errMes
end sub



Sub cmdBookmarkRead(e)  ' ...
  trace "OK","OK"
  dim grid as iGridEx=IGrid3
  dim fileName=getDizPlus+grid.fileSpec

  ''  msgbox("fileName: "+fileName)
  
  Dim Content as string =""
  Content = readWebFile(FileName)
  '' Content = ZZ.fileGetContents("C:\yPara\scriptIDE\codeBookmarks.txt")
  '' msgBox(content)

  iniIGrid(grid)
  Igrid_put (grid, Content)

  dim i as integer
  dim max as integer=grid.rows.count
  dim row as igRow
  for i=0 to max-1
    row = grid.rows(i)
     '' .cells(cNickname).backColor = ColorTranslator.FromHtml("#7A8599")
     '' .cells(cNickname).foreColor = ColorTranslator.FromHtml("#fff")
    row.cells(cNickname).backColor = ColorTranslator.FromHtml("#7F8AA6")
    row.cells(cNickname).foreColor = ColorTranslator.FromHtml("#fff")
    row.ContentIndent= New TenTec.Windows.iGridLib.iGIndent(1, 0, 1, 0)
  next

  formatCodeList(iGrid3)
  on_Resize_Controls()

'' 
''   '#######################################
'' 
'' '' readGridData(iGrid3)
''   '' glob_defaultPath
''   
''   '#######################################
''   dim grid=IGrid3
''   dim fileName=getDizPlus+grid.fileSpec  
'' 
''   dim fileSpec= glob_defaultPath+fileName
''   msgbox(fileSpec)
''   Dim fileContent = IO.File.ReadAllText(fileSpec)
''   'msgbox(fileContent+"<--")
''   IGrid_put(grid, fileContent)
End Sub


Function readWebFile(strFileName as string, optional strAppName as string="") 'lesen
  printLine 11, "readWebFile", strFileName
  if strAppName="" then strAppName=globWebAppName
  dim strContent as string
  strContent=ZZ.TwReadFile(strAppName, strFileName) 
  return strContent
end Function



'--> ...


'' sub readSnippetFile()
''   trace "OK","OK"
''   dim FileName as string = "es-code-snippets.txt"
'' 
''   Dim Content as string =""
''   Content = readWebFile(FileName)
''   '' Content = ZZ.fileGetContents("C:\yPara\scriptIDE\codeBookmarks.txt")
''   '' msgBox(content)
'' 
''   Igrid_put (IGrid1, Content)
'' End Sub
'' 
'' sub saveSnippetFile()
''   Dim Content As String
''   Igrid_get (IGrid1, Content)
'' 
''   Dim myDate As Date = Now()
''   dim path as string="C:\yPara\scriptIDE\codeBookmarks\"
''   dim fileName=myDate.ToString("yyyy-MM-dd__HH_mm__")+"codeBookmarks.txt"
''   printLine 2, filename, filename
''   ZZ.filePutContents(path+fileName, Content)
'' 
''   saveWebFile("", Content)
'' end sub
'' 






sub saveGridData(grid as IgridEx)
  dim fileSpec=grid.fileSpec  
  Dim data As String
  IGrid_get (grid, data)
  'msgBox(data)
  'dim subFolder=IO.Path.GetFilenameWithoutExtension(fileName)+"\"
  Dim myPath as string = IO.Path.GetDirectoryName(fileSpec)
  ''msgBox(myPath)
  Directory.CreateDirectory(myPath)
  'IO.File.WriteAllLines(fileSpec, lines)
  IO.File.WriteAllText(fileSpec, data)
end sub

Function readGridData(grid as IgridEx)
  dim fileSpec=grid.fileSpec  
  Dim fileContent = IO.File.ReadAllText(fileSpec)
  'msgbox(fileContent+"<--")
  IGrid_put(grid, fileContent)
  dim i as integer
  dim max as integer=grid.rows.count
  dim row as igRow
  for i=0 to max-1
    row = grid.rows(i)
     '' .cells(cNickname).backColor = ColorTranslator.FromHtml("#7A8599")
     '' .cells(cNickname).foreColor = ColorTranslator.FromHtml("#fff")
    row.cells(cNickname).backColor = ColorTranslator.FromHtml("#7F8AA6")
    row.cells(cNickname).foreColor = ColorTranslator.FromHtml("#fff")
    row.ContentIndent= New TenTec.Windows.iGridLib.iGIndent(1, 0, 1, 0)
  next

end Function







'-->
'--> --------------------------- L I B s (auslagern) 






'--> 
'--> S C I N T I L L A 


'--> ... auslagern als Lib '???
Sub highlightExecutingLine(ByVal line As Integer)
  'ScintillaNet.MarkerSymbol.ShortArrow
  'setLineMarker(line, 12, My.Resources.executing, Color.Yellow, Color.Orange, True)

  'setLineMarker(line, 12, 0, Color.Yellow, Color.Orange, True)
  'setLineMarker(line, 13, 0, ColorTranslator.FromHtml("#FFFF7B"), Color.Orange, True) 'vb: #FFEE62
  
  setLineMarker(line, 12, 0, ColorTranslator.FromHtml("#bbbbbb"), Color.Orange, True)
  setLineMarker(line, 13, 0, ColorTranslator.FromHtml("#ccccff"), Color.Orange, True) 'vb: #FFEE62
  
  If line < 0 Then Exit Sub
  sc1.GoTo.Line(line)
  ' sc1.CallTip.HighlightStart = sc1.Lines(line).SelectionStartPosition
  ' sc1.CallTip.HighlightEnd = sc1.Lines(line).SelectionEndPosition
  ' sc1.CallTip.HighlightTextColor = Color.Red
  ' sc1.CallTip.Message = "Pausiert: " & line
  ' sc1.CallTip.Show()
End Sub


  
Sub highlightExecutingLine2(ByVal line As Integer)
  'ScintillaNet.MarkerSymbol.ShortArrow
  'setLineMarker(line, 15, My.Resources.executing4, Color.Yellow, Color.Orange, True)
  setLineMarker(line, 15, 0, Color.Yellow, Color.Orange, True)
  setLineMarker(line, 13, 0, ColorTranslator.FromHtml("#AFFF7A"), Color.Orange, True) 'vb: #FFEE62
  If line < 0 Then Exit Sub
  sc1.GoTo.Line(line)
  ' sc1.CallTip.HighlightStart = sc1.Lines(line).SelectionStartPosition
  ' sc1.CallTip.HighlightEnd = sc1.Lines(line).SelectionEndPosition
  ' sc1.CallTip.HighlightTextColor = Color.Red
  ' sc1.CallTip.Message = "Pausiert: " & line
  ' sc1.CallTip.Show()
End Sub
  
  
Sub highlightErrorLine(ByVal line As Integer)
  setLineMarker(line, 10, ScintillaNet.MarkerSymbol.Arrows, Color.Transparent, Color.Red, True)
  setLineMarker(line, 11, 0, ColorTranslator.FromHtml("#EAB9FA"), Color.White, True)
  If line < 0 Then Exit Sub
  sc1.GoTo.Line(line)
End Sub
  
  
Sub setLineMarker(ByVal line As Integer, ByVal markerIndex As Integer, ByVal markerIcon As Object, ByVal bgColor As Color, ByVal fgColor As Color, Optional ByVal removeOthers As Boolean = False, Optional ByVal toggleMe As Boolean = False)
  If removeOthers Or line = -1 Then sc1.Markers.DeleteAll(markerIndex)
  If line < 0 Then Exit Sub
  If toggleMe = True And (sc1.Lines(line).GetMarkerMask And 16) = 16 Then
    sc1.Lines(line).DeleteMarker(4) : Exit Sub
  End If
  Dim m = sc1.Lines(line).AddMarker(markerIndex)
  If IsNumeric(markerIcon) Then
    m.Marker.Symbol = markerIcon
  ElseIf TypeOf markerIcon Is Image Then
    m.Marker.Symbol = ScintillaNet.MarkerSymbol.PixMap
    m.Marker.SetImage(markerIcon, Color.FromArgb(255, 255, 255, 255))
  Else
    m.Marker.Symbol = ScintillaNet.MarkerSymbol.PixMap
    m.Marker.SetImage(Image.FromFile(markerIcon), Color.FromArgb(255, 255, 255, 255))
  End If

  m.Marker.ForeColor = fgColor
  m.Marker.BackColor = bgColor
End Sub




'-->
'-->  I G R I D - E X 

Class IgridEx

Inherits Tentec.Windows.Igridlib.IGrid
Public Event updateUserFeedback1(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String)
Public Event updateUserFeedback2(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String)

'sub IGrid1_updateUserFeedback1(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String) Handles iGrid1.updateUserFeedback1 '@@-
'sub IGrid1_updateUserFeedback2(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String) Handles iGrid1.updateUserFeedback2 '@@-

private g_rootDir as string
: Property rootDir () As String
:   Get 
:     Return g_rootDir
:   End Get
:   Set (ByVal value As String)
:     g_rootDir=value
:   End Set
: End Property


private g_fileSpec as string
: Property fileSpec () As String
:   Get 
:     Return g_fileSpec
:   End Get
:   Set (ByVal value As String)
:     g_fileSpec=value
:   End Set
: End Property

  '' Property ItemText(ByVal idx) As String
  ''   Get
  ''     Return Me.Items(idx)
  ''   End Get
  ''   Set(ByVal value As String)
  ''     Me.Items(idx) = value
  ''   End Set
  '' End Property


private g_name as string
: Property name () As String
:   Get 
:     Return g_name
:   End Get
:   Set (ByVal value As String)
:     g_name=value
:   End Set
: End Property

private g_allowSort as boolean=false
: Property allowSort ()as boolean
:   Get 
:     Return g_allowSort
:   End Get
:   Set (ByVal value as boolean)
:     g_allowSort=value
:   End Set
: End Property

private g_deltaIndex as integer=-1
: Property deltaIndex ()as integer
:   Get 
:     Return g_deltaIndex
:   End Get
:   Set (ByVal value as integer)
:     g_deltaIndex=value
:   End Set
: End Property


Sub myIgrid_ColHdrClick(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGColHdrClickEventArgs) Handles myBase.ColHdrClick
  '... Default: nicht sortieren
  e.DoDefault = g_allowSort
End Sub


: Sub myIgrid_CellMouseDown(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseDownEventArgs) Handles myBase.CellMouseDown
  printLine  6, "IGrid1_CellMouseDown - row", e.RowIndex
  printLine  7, "IGrid1_CellMouseDown - col", e.ColIndex
  if e.Button= Windows.Forms.MouseButtons.Right then
    printLine  8, "IGrid1_CellMouseDown - col", "rClick"
  else
    printLine  8, "IGrid1_CellMouseDown - col", "NORMAL"
  end if
End Sub


: Sub myIgrid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles myBase.KeyDown
  dim keyCode=e.keycode
  printLine  4, "myIgrid_KeyDown", e.keycode 
  '' 'test: fehler
  '' dim y as object
  '' y.kjdfkdjskfj
    
  '' msgBox ("aaaaaaaaaaaaaaa")
  printLine  5, "myIgrid_KeyDown", e.keycode 
  if e.control=true then
    'trace "CONTROL",e.keycode 
    'myScriptClass.monitorAdd("CONTROL: "+keyCode.toString)
    if keycode=88 then onClipboardCutSelected(me)   'X
    if keycode=67 then onClipboardCopySelected(me)  'C
    if keycode=86 then onClipboardPasteBeforeSelectedLine(me) 'V
    e.handled =true
  end if
End Sub


: sub onClipboardCutSelected(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  myScriptClass.monitorAdd("CONTROL: onClipboardCut()")
  onClipboardCopySelected(iGrid,true)
end sub


: sub onClipboardCopySelected(ByVal iGrid As TenTec.Windows.iGridLib.iGrid, optional remove as boolean=false)
  myScriptClass.monitorAdd("CONTROL: onClipboardCopy()")
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim anz as integer
  dim max =iGrid.Rows.count
  dim OUT(max) as string
  dim DEL(max) as integer
  dim counter as integer=0
  For Each row In iGrid.Rows
    'Debug.Print(row.Selected)
    if row.Selected=true then 
      'myScriptClass.monitorAdd(row.index.toString)
      OUT(counter)=myScriptClass.JoinIGridLine(row)
      DEL(counter)=row.index
      counter=counter+1
      anz = anz +1
    end if
  Next
  redim preserve OUT(counter-1)
  redim preserve DEL(counter-1)
  if remove=true then
    dim i as integer
    for i=DEL.length-1 to 0 step -1
      iGrid.Rows.RemoveAt(DEL(i))
    next
  end if
  dim RESULT as string=join(OUT, vbNewLine)
  zz.setClipboardText(RESULT)

  myScriptClass.monitorAdd(RESULT)
  myScriptClass.monitorAdd("==========================================")
  myScriptClass.monitorAdd("Anz. selected: "+anz.toString)
  myScriptClass.monitorAdd("")
end sub

: sub onClipboardPasteBeforeSelectedLine(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  myScriptClass.monitorAdd("CONTROL: onClipboardPasteBeforeSelectedLine()")
  dim selIndex as integer = -1
  
  '...die erste selektierte zeile ermitteln
  Dim row As TenTec.Windows.iGridLib.iGRow
  For Each row In iGrid.Rows
    if row.Selected=true then
      if selIndex=-1 then selIndex=row.index
      row.Selected=false
    end if
  Next

  '... noch nicht entschieden
  if selIndex=-1 then
    '... es ist nichts Selektiert
    '... in Zeile 1 einfuegen / ... oder ans ende ? ? ?
    '??? ist grid ueberhaupt initialisiert
  end if
  
  dim newData as string =zz.getClipboardText()
  myScriptClass.monitorAdd(newData)
  dim DATA=split(newData,vbNewLine)
  dim i as integer=0
  dim skip as boolean =true
  for i=DATA.length-1 to 0 step -1
    dim item=DATA(i)
    iGrid.Rows.Insert(selIndex)
    myScriptClass.SplitToIGridLine(iGrid.Rows(selIndex), item)
    '... alle ausser der ersten einfuegung
    if skip = true  then 
      skip = false
    else
      iGrid.Rows(selIndex+1).selected=true
    end if
  next
  '...den ersten auch
  iGrid.Rows(selIndex).selected=true
  '' myScriptClass.monitorAdd(RESULT)
  '' myScriptClass.monitorAdd("==========================================")
  '' myScriptClass.monitorAdd("Anz. selected: "+anz.toString)
  '' myScriptClass.monitorAdd("")
end sub



: Private Sub IgridEx_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.SelectionChanged
  trace "IGrid1_SelectionChanged..."
  checkSelectionInRowMode(me)
  trace "---------------------------------" 
End Sub

: Private Sub IgridEx_CurRowChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.CurRowChanged
  dim curRow as integer
  curRow=me.curRow.index
  trace "curRow", curRow
  RaiseEvent updateUserFeedback2(me, curRow.ToString)

End Sub

:Sub checkSelectionInRowMode(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim anz as integer
  For Each row In iGrid.Rows
    'Debug.Print(row.Selected)
    if row.Selected=true then anz = anz +1
  Next
  RaiseEvent updateUserFeedback1(me, anz.ToString)

End Sub

:Sub resetAllSelections(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim anz as integer
  For Each row In iGrid.Rows
    if row.Selected=true then row.Selected=false
  Next
End Sub

End Class

















'-->
'--> I g r i d H e l p e r 2 

  
: Function JoinIGridLine(ByVal line As iGRow, Optional ByVal Delimiter As String = vbTab) As String
  on error resume next
  Dim max = line.Cells.Count - 1
  Dim out(max) As String
  :For i As Integer = 0 To max
    :out(i) = line.Cells(i).Value.ToString
  :Next
  :Return Join(out, Delimiter)
End Function

: Sub SplitToIGridLine(ByVal line As iGRow, ByVal input As String, Optional ByVal Delimiter As String = vbTab)
  on error resume next
  Dim max = line.Cells.Count - 1
  Dim out() = Split(input, Delimiter)
  ReDim Preserve out(max)
  :For i As Integer = 0 To max
  :  line.Cells(i).Value = out(i)
  :Next
End Sub

: Sub Igrid_get(ByVal Grid As iGrid, ByRef FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
  on error resume next
  Dim max = Grid.Rows.Count - 1
  Dim out(max) As String
  For i As Integer = 0 To max
    out(i) = JoinIGridLine(Grid.Rows(i), ColDelim)
  Next
  FileContent = Join(out, LineDelim)
End Sub
  
  
: Sub Igrid_put(ByVal Grid As iGrid, ByVal FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
  on error resume next
  Dim out() = Split(FileContent, LineDelim)
  'msgbox(FileContent+"<--")
  Grid.Rows.Clear()
  Grid.Rows.Count = out.Length
  For i As Integer = 0 To out.Length - 1
    SplitToIGridLine(Grid.Rows(i), out(i), ColDelim)
  Next
  'msgBox(out(out.Length - 2))
  'msgBox(out(out.Length - 1))
  '' if trim(out(out.Length - 1))<>"" then
  ''   SplitToIGridLine(Grid.Rows(out.Length - 1), out(out.Length - 1), ColDelim)
  '' end if
End Sub

: Sub Igrid_put2(ByVal Grid As iGrid, ByVal FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
  on error resume next
  Dim out() = Split(FileContent, LineDelim)
  Grid.Cols.Clear
  Grid.Rows.Count = out.Length
  dim header()=Split(out(0), ColDelim)
  for j as integer=0 to header.length-1
    grid.Cols.Add(header(j),66)
  next
  For i As Integer = 0 To out.Length - 1
    SplitToIGridLine(Grid.Rows(i), out(i), ColDelim)
  Next
End Sub


: sub igrid_setDefaultPara(iGrid)
  on error resume next
  with iGrid
    .DefaultCol.CellStyle.TextTrimming = TenTec.Windows.iGridLib.iGStringTrimming.None'=IGCellStyleDesign2
    .DefaultCol.CellStyle.EmptyStringAs = iGEmptyStringAs.EmptyString
    .DefaultCol.CellStyle.EmptyStringAs = TenTec.Windows.iGridLib.iGEmptyStringAs.EmptyString 'zack!
    .DefaultCol.ColHdrStyle.TextTrimming = TenTec.Windows.iGridLib.iGStringTrimming.None
    .SelCellsBackColorNoFocus = System.Drawing.SystemColors.Highlight
    .SelCellsForeColorNoFocus = System.Drawing.SystemColors.HighlightText
  end with
end sub




'-->
'--> D R A G - D R O P  - MANAGER


'' Imports System
'' Imports System.Drawing
#Imports System.Drawing.Drawing2D
'' Imports System.Windows.Forms
'' Imports TenTec.Windows.iGridLib
'' 


Class iGDragDropManager
Public Event DragDropDone(ByVal MovedGridRow As iGRow, ByVal sourceGrid As iGrid, ByVal destGrid As iGrid, ByVal fDstIndex As Integer)

public globDragDorp_sourceGrid as iGrid
public globDragDorp_destGrid as iGrid

'#Region "Public Methods"

:  Public Sub New()
  
    MyBase.New()
  
  End Sub



  Public Sub Manage(ByVal grid As iGrid, ByVal canDragFrom As Boolean, ByVal canDropTo As Boolean)

    ' Implement start of dragging.
    If canDragFrom Then
      trace "!!! Manage"
      trace grid.height.toString
      trace "!!! Manage"
       'AddHandler grid.StartDragCell, AddressOf grid_StartDragCell
      AddHandler grid.StartDragCell, AddressOf grid_StartDragCell
    End If

    ' Implement dropping.
    If canDropTo Then
      grid.AllowDrop = True
      AddHandler grid.DragDrop, AddressOf grid_DragDrop
      AddHandler grid.DragOver, AddressOf grid_DragOver
      AddHandler grid.DragLeave, AddressOf grid_DragLeave
    End If

    ' Implement destination indication.
    If canDropTo Then
      AddHandler grid.Paint, AddressOf grid_Paint
    End If

  End Sub

'' #End Region
'' 
'' #Region "Fields"
 
  ''' <summary>
  ''' Stores the index of the row which to insert
  ''' the rows being dragged before.
  ''' </summary>
  Private fDstIndex As Integer

  ''' <summary>
  ''' The grid where to insert the rows being dragged.
  ''' </summary>
  Private fDstGrid As iGrid

  ''' <summary>
  ''' Index of the row used to start dragging.
  ''' </summary>
  Private fSrcIndex As Integer

  ''' <summary>
  ''' The grid which cells are being dragged.
  ''' </summary>
  Private fSrcGrid As iGrid

  ''' <summary>
  ''' Indicates whether rows can be dropped between other rows,
  ''' or can only be added to the end.
  ''' </summary>
  Private fAllowInsert As Boolean = True

  ''' <summary>
  ''' Indicates whether rows can only be copied.
  ''' </summary>
  Private fAllowCopy As Boolean = True

  ''' <summary>
  ''' Indicates whether rows can only be moved.
  ''' </summary>
  Private fAllowMove As Boolean = True

  ''' <summary>
  ''' Indicates whether rows can be moved within a grid.
  ''' </summary>
  Private fAllowMoveWithinOneGrid As Boolean

'' #End Region
'' 
'' #Region "Private Methods"
 
  ''' <summary>
  ''' Starts dragging.
  ''' </summary>
  Private Sub grid_StartDragCell(ByVal sender As Object, ByVal e As iGStartDragCellEventArgs)
    On Error Resume Next
    trace "::: grid_StartDragCell"
    Dim myGrid As iGrid = CType(sender, iGrid)

    globDragDorp_sourceGrid = myGrid
    trace  "m_sourceGrid: " + globDragDorp_sourceGrid.Name 

    ' Start dragging.
    ResetDst()
    fSrcGrid = myGrid
    fSrcIndex = e.RowIndex
    Dim objDO As New DataObject
    Dim ir = myGrid.SelectedRows(0)
    Dim id As String = ir.Key.ToString
    Dim titel As String = ir.Cells("titel").Value

    ''objDO.SetData(DataFormats.Text, "\\projektLink-" + id + " [ " + titel + " ]")
    
    Dim rtfInsert As String ="??? noch nicht implementiert"
    '--> '!!! fehlt noch: Dim rtfInsert As String = creatertfProjektLink(id, titel)
    '--> '!!! fehlt noch: insertCodeLink = creatertfProjektLink(id, titel)

    ''rtfInsert = "{\rtf1\ansi\ansicpg1252\deff0\deflang1031{\fonttbl{\f0\fnil\fcharset0 Courier New;}}" + _
    ''"{\colortbl ;\red0\green0\blue0;\red211\green211\blue211;}" + _
    ''"\uc1\pard\f0\fs18\\\\projektLink-||ID||  \cf1\highlight2\ul  ||TITEL||  }\\\\projektLink-||ID||  \cf1\highlight2\ul ||TITEL|| \par \cf0\highlight0\ulnone\par"

    ''rtfInsert = Replace(rtfInsert, "||TITEL||", titel)
    ''rtfInsert = Replace(rtfInsert, "||ID||", id)

    objDO.SetData(DataFormats.Rtf, rtfInsert)
    'objDO.SetData(DataFormats.Text, id + " | " + titel)
    objDO.SetData(DataFormats.Text, "  codeDefLine: ...coming soon (space: wie ermitteln ???"+vbNewLine)
    objDO.SetData(myGrid.SelectedRows)

    myGrid.DoDragDrop(objDO, DragDropEffects.Move Or DragDropEffects.Copy)
    'myGrid.DoDragDrop(myGrid.SelectedRows, DragDropEffects.Move Or DragDropEffects.Copy)

  End Sub
   ''' <summary>
  ''' Checks whether a grid can accept the data being dragged.
  ''' </summary>
  Private Sub grid_DragOver(ByVal sender As Object, ByVal e As DragEventArgs)

    ''e.Effect = DragDropEffects.Copy
    ''Debug.Print("grid_DragOver: " + Now.ToString)
    ''Exit Sub
    
    'trace "grid_DragOver"
    
    
    globDragDorp_destGrid = sender
    If e.Data.GetDataPresent(GetType(iGSelectedRowsCollection)) AndAlso Not fSrcGrid Is Nothing Then
      ' Stop

      ResetDst()

      Dim myGrid As iGrid= CType(sender, iGrid)
      If Not myGrid Is Nothing AndAlso (Not fSrcGrid Is myGrid OrElse fAllowMoveWithinOneGrid) Then

        Dim myDstIndex As Integer
        Dim myMousePos As Point = New Point(e.X, e.Y)
        myMousePos = myGrid.PointToClient(myMousePos)

        ' Scroll the grid automatically if required.
        Dim myDeltaScroll As Integer = myGrid.DefaultRow.Height \ 2
        Dim myCellsAreaBounds As Rectangle = myGrid.CellsAreaBounds
        If myMousePos.Y - myCellsAreaBounds.Top <= myDeltaScroll Then
          myGrid.VScrollBar.Value -= myGrid.VScrollBar.SmallChange
        ElseIf myCellsAreaBounds.Bottom - myMousePos.Y <= myDeltaScroll Then
          myGrid.VScrollBar.Value += myGrid.VScrollBar.SmallChange
        End If


        If fAllowInsert Then
          ' Get the cell under the mouse pointer if any.
          myDstIndex = GetDstIndexFromPoint(myGrid, myMousePos.X, myMousePos.Y)
        Else
          myDstIndex = myGrid.Rows.Count
        End If

        If myDstIndex >= 0 AndAlso (Not fSrcGrid Is myGrid OrElse fSrcGrid.SelectedRows.Count > 1 OrElse (myDstIndex <> fSrcIndex AndAlso myDstIndex <> fSrcIndex + 1)) Then

          ' Enable the drop operation.
          fDstGrid = myGrid
          fDstIndex = myDstIndex
          If fAllowMove AndAlso fAllowCopy Then
            If Control.ModifierKeys = Keys.Control Then
              e.Effect = DragDropEffects.Copy
            Else
              e.Effect = DragDropEffects.Move
            End If
          ElseIf fAllowCopy Then
            e.Effect = DragDropEffects.Copy
          ElseIf fAllowMove Then
            e.Effect = DragDropEffects.Move
          End If

          ' Redraw the destination grid.
          InvalidateDst()

          Return
        Else
 
        End If

      End If

      End If

    ' Disable the drop operation.
    '' ??? e.Effect = DragDropEffects.None
    'e.Effect = DragDropEffects.Copy
    e.Effect = DragDropEffects.Move Or DragDropEffects.Copy
    '' Debug.Print("DragDropEffects")
  End Sub

  ''' <summary>
  ''' Gets the destination row index from the specified mouse 
  ''' coordinates.
  ''' </summary>
  Private Function GetDstIndexFromPoint(ByVal grid As iGrid, ByVal x As Integer, ByVal y As Integer) As Integer

    ' Get the row under at the point.
    Dim myRowHeight, myRowY As Integer
    Dim myRow As iGRow = grid.Rows.FromY(y, myRowY, myRowHeight)

    ' If the mouse is above a row, return its or the previous row's index.
    If Not myRow Is Nothing Then
      If y < myRowY + myRowHeight / 2 Then
        Return myRow.Index
      Else
        Return myRow.Index + 1
      End If
    End If

    
    If grid.CellsAreaBounds.Contains(x, y) Then
      Return grid.Rows.Count
    End If

    Return -1

  End Function

  ''' <summary>
  ''' Drop the rows being dragged.
  ''' </summary>
  Private Sub grid_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
    On Error Resume Next
    trace "::: grid_DragDrop" 

    'Debug.Print(fSrcGrid.CurRow.Index)
    trace "m_sourceGrid-2: " + globDragDorp_sourceGrid.Name
    'Debug.Print(globDragDorp_sourceGrid.Tag.ToString)
    trace globDragDorp_destGrid.Name 
    'Debug.Print(globDragDorp_destGrid.Tag.ToString)

    If fDstIndex >= 0 AndAlso Not fDstGrid Is Nothing AndAlso Not fSrcGrid Is Nothing Then
      'Stop

      Dim myTag As String = "Just added"
      Dim MovedRow As iGRow = Nothing






      If Not fSrcGrid Is fDstGrid Then
        fDstGrid.PerformAction(iGActions.DeselectAllRows)
      End If

  

      ' Add rows and copy cell values.
      trace fSrcGrid.CurRow.Index 
      trace fSrcGrid.SelectedCells.Count 
      trace "fDstGrid.Rows.Insert(fDstIndex)", fDstIndex 
      
      '!!! es
      If fSrcGrid.SelectedCells.Count = 0 Then
         '''fSrcGrid.Rows(fSrcGrid.CurRow.Index).Selected = True
         fSrcGrid.Rows(fSrcIndex).Selected = True
       End If
      trace fSrcGrid.SelectedCells.Count 

       trace "!!! fSrcGrid", fSrcGrid.SelectedCells.Count
        trace "!!! fDstGrid"
      ''For myIndex As Integer = fSrcGrid.SelectedCells.Count - 1 To 0 Step -1
      For myIndex As Integer = fSrcGrid.SelectedCells.Count - 0 To 0 Step -1
        Dim myRow As iGRow = fDstGrid.Rows.Insert(fDstIndex)
        
        trace "+++ fSrcGrid", myRow.index
        trace "+++ fDstGrid", fSrcIndex
        
        If fSrcGrid Is fDstGrid Then
          myRow.Tag = myTag
          
          trace "!!! If fSrcGrid Is fDstGrid Then"
          trace "!!! If fSrcGrid Is fDstGrid Then",myTag
        Else
          myRow.Selected = True
        End If

        MovedRow = myRow
        myRow.BackColor = Color.White

        'Dim mySrcRowIndex As Integer = fSrcGrid.SelectedCells(myIndex).RowIndex
        Dim mySrcRowIndex As Integer = fSrcIndex+1+myIndex 'fSrcGrid.Cells(fSrcIndex+myIndex).RowIndex
        For myColIndex As Integer = 0 To fDstGrid.Cols.Count - 1
          myRow.Cells(myColIndex).DropDownControl = fSrcGrid.Cells(mySrcRowIndex, myColIndex).DropDownControl
          myRow.Cells(myColIndex).Style = fSrcGrid.Cells(mySrcRowIndex, myColIndex).Style
          myRow.Cells(myColIndex).Value = fSrcGrid.Cells(mySrcRowIndex, myColIndex).Value
          myRow.Cells(myColIndex).BackColor = fSrcGrid.Cells(mySrcRowIndex, myColIndex).BackColor

        Next
      Next

      ' !!! SPEZIAL:
      If fSrcGrid.Cols.Count <> fDstGrid.Cols.Count Then
        '...sonderDragDrop !!!
        trace "DragDropDone- ...sonderDragDrop !!!"
        RaiseEvent DragDropDone(Nothing, fSrcGrid, fDstGrid, fDstIndex)
        Exit Sub
      End If

      ' Ensure visible dropped rows.
      If Not fSrcGrid Is fDstGrid Then
        MovedRow.EnsureVisible()
      End If

      ' Delete source rows if it is needed.
      'If e.Effect = DragDropEffects.Move Then
      
      fSrcGrid.Rows.RemoveAt(fSrcIndex+1)
        For Each selCol As iGCell In fSrcGrid.SelectedCells
          fSrcGrid.Rows.RemoveAt(selCol.RowIndex)
        Next
      'End If

      ' If rows were moved within a grid, select them.
      If e.Effect = DragDropEffects.Move Then
        If fSrcGrid Is fDstGrid Then
          For Each myCell As iGCell In fSrcGrid.Cells
            If myCell.Row.Tag Is myTag Then
              myCell.Row.Tag = Nothing
              myCell.Selected = True
              Exit For
            End If
          Next
        End If

      End If

      fDstGrid.Focus()
      trace "DragDropDone- NORMAL"
      RaiseEvent DragDropDone(MovedRow, fSrcGrid, fDstGrid, fDstIndex)
      fSrcGrid.refresh
      'ResetDst()
      'fSrcGrid = Nothing
      Exit Sub
    Else
       trace "DragDropDone- Nothing"
       RaiseEvent DragDropDone(Nothing, globDragDorp_sourceGrid, globDragDorp_sourceGrid, 0)

    End If

  End Sub

  ''' <summary>
  ''' Resets internal fields when dragging is cancelled.
  ''' </summary>
  Private Sub grid_DragLeave(ByVal sender As Object, ByVal e As EventArgs)


    '' ??? wer stop dragDrop ???

    ResetDst()

  End Sub

  ''' <summary>
  ''' Resets the fields stoting the destination.
  ''' </summary>
  Private Sub ResetDst()
    '' Exit Sub ' ???

    InvalidateDst()

    fDstIndex = -1
    fDstGrid = Nothing

  End Sub

  ''' <summary>
  ''' Causes the current destination position to redrawn.
  ''' </summary>
  Private Sub InvalidateDst()

    If fDstIndex >= 0 AndAlso Not fDstGrid Is Nothing Then
      If fDstIndex < fDstGrid.Rows.Count Then
        fDstGrid.Invalidate(New Rectangle(0, fDstGrid.Rows(fDstIndex).Y - 2, fDstGrid.Width, 3))
      ElseIf fDstIndex > 0 Then
        fDstGrid.Invalidate(New Rectangle(0, fDstGrid.Rows(fDstIndex - 1).Y + fDstGrid.Rows(fDstIndex - 1).Height - 2, fDstGrid.Width, 3))
      Else
        fDstGrid.Invalidate(New Rectangle(0, fDstGrid.CellsAreaBounds.Y, fDstGrid.Width, 1))
      End If

    End If

  End Sub

  ''' <summary>
  ''' Draws the destination.
  ''' </summary>
  Private Sub grid_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)
    On Error Resume Next

    If Not fAllowInsert Then
      Return
    End If

    If Not fAllowCopy AndAlso Not fAllowMove Then
      Return
    End If

    If fDstIndex < 0 OrElse Not fDstGrid Is sender Then
      Return
    End If

    Dim myCellsAreaBounds As Rectangle = fDstGrid.CellsAreaBounds

    Dim myY As Integer
    If fDstIndex < fDstGrid.Rows.Count Then
      myY = fDstGrid.Rows(fDstIndex).Y - 2
    ElseIf fDstIndex > 0 Then
      myY = fDstGrid.Rows(fDstIndex - 1).Y + fDstGrid.Rows(fDstIndex - 1).Height - 2
    Else
      myY = myCellsAreaBounds.Y - 2
    End If

    Dim myOldClip As Region = e.Graphics.Clip
    e.Graphics.SetClip(myCellsAreaBounds, CombineMode.Intersect)
    ''Try
    e.Graphics.FillRectangle(Brushes.SandyBrown, New Rectangle(0, myY, fDstGrid.Width, 3))
    ''Finally
    e.Graphics.Clip = myOldClip
    ''End Try

  End Sub

'' #End Region
'' 
'' #Region "Public Poperties"

  ''' <summary>
  ''' Gets or sets a value indicating whether rows can be dropped 
  ''' between other rows, or can only be added to the end.
  ''' </summary>
  Public Property AllowInsert() As Boolean
    Get
      Return fAllowInsert
    End Get
    Set(ByVal Value As Boolean)
      fAllowInsert = Value
    End Set
  End Property

  ''' <summary>
  ''' Gets or sets a value indicating whether rows can be copied.
  ''' </summary>
  Public Property AllowCopy() As Boolean
    Get
      Return fAllowCopy
    End Get
    Set(ByVal Value As Boolean)
      fAllowCopy = Value
    End Set
  End Property

  ''' <summary>
  ''' Gets or sets a value indicating whether rows can be moved.
  ''' </summary>
  Public Property AllowMove() As Boolean
    Get
      Return fAllowMove
    End Get
    Set(ByVal Value As Boolean)
      fAllowMove = Value
    End Set
  End Property

  ''' <summary>
  ''' Gets or sets a value indicating whether rows can be moved within a grid.
  ''' </summary>
  Public Property AllowMoveWithinOneGrid() As Boolean
    Get
      Return fAllowMoveWithinOneGrid
    End Get
    Set(ByVal Value As Boolean)
      fAllowMoveWithinOneGrid = Value
    End Set
  End Property

'' #End Region

End Class



'-->
'--> o u t M O N I T O R 

public globMonitorRef as object

sub clearAll()
  setMonitorRef()
  on error resume next
  monitorClear()
  statustext_reset()
  zz.traceClear()
  zz.printLineReset()
  err.number=0
end sub



sub setMonitorRef()
  'on error resume next
  globMonitorRef = zz.scriptClass("es_testLabor")
  'trace "globMonitorRef", globMonitorRef.name
end sub

: function getMonitorRef()
  on error resume next
  if globMonitorRef is nothing then
    globMonitorRef = zz.scriptClass("es_testLabor")
  end if
  return globMonitorRef
end function

sub monitorClear()
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



'-->
'-->E V E N T - T E M P L A T E


sub dumpUnknownEventHandler(funcName)
  dim tpl=getTemplateUnknownEventHandler()
  tpl=replace(tpl,"||FUNC-NAME||",funcName)
  MonitorAdd(tpl)
end sub


function getTemplateUnknownEventHandler()
'--> ### data-block...
#Data myData as String

ERR: EventHandler fuer '||FUNC-NAME||' nicht gefunden:
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
'--> ---------------------------------------- O L D 




'-->   
'-->   N A V I G A T E - O L D 



function xxx_collapsExpandHiddenTopics(grid As iGrid,  rowIndex as integer )
  '' on error resume next
  '' dim newTempExpandCaption as string=""
  '' 
  '' 'static lastTempExpandCaption as string=""
  '' if iGrid1.curRow.visible = true  then return newTempExpandCaption          '... E X I T 
  '' 
  '' if iGrid1.curRow.visible = false then
  ''   monitorAdd("iGrid1.curRow.visible = false")
  ''   '... der setzt lastTempExpandItem gleichzeitig zurueck
  ''   newTempExpandCaption = indexListExpandTempMode(grid, rowIndex, globTempExpandCaption)
  ''   ' globTempExpandCaption = newTempExpandCaption
  ''    return newTempExpandCaption                                                                             '... E X I T 
  '' end if
  '' 
  '' '' if lastTempExpandCaption <> "" then
  '' ''   '...collapsTempMode
  '' ''   indexListCollapsTempMode(grid, lastTempExpandCaption)
  '' '' end if
  '' '' 
  '' return newTempExpandCaption
end function



function xxx_getCodeLinkItem2(grid As iGrid,Row As iGRow)
  'checkForIndexListUpdate(grid, row.index)
  dim i as integer
  dim max as integer=grid.cols.count
  dim OUT(max)as string
  for i =0 to max-1
    OUT(i)=row.cells(i).value
  next
  return OUT
end function





sub xxx__navigateCodeLink2(grid As iGrid, curRow As iGRow)
  '' on error resume next
  '' dim codeLinkItem() as string=getCodeLinkItem2(grid,curRow)
  '' 'msgBox(join(codeLinkItem,vbnewLine))
  '' historyAddItem(codeLinkItem)
  '' 
  '' if trim(curRow.cells(cId).value) = "" then exit sub
  '' 
  '' dim fileSpec as string=curRow.cells(cFileSpec).value
  '' dim codeLineNumber as integer = cInt(curRow.cells(cId).value)
  '' 
  '' ''msgBox(fileSpec)
  '' dim tab = ZZ.IDEHelper.NavigateFile(fileSpec)
  '' 
  '' 
  '' dim codeEitor = ZZ.getActiveRTF.RTF
  '' codeEitor.goTo.line(codeLineNumber+50)
  '' codeEitor.goTo.line(codeLineNumber-10)
  '' codeEitor.goTo.line(codeLineNumber)
  '' highlightExecutingLine(codeLineNumber)
  '' codeEitor.focus()
  '' err.number=0
End Sub 






sub xxx__insertCodeLink

  '' '--> wird demnaechst ausgemustert !!!
  '' 
  '' dim codeLinkDATA = getCodeLinkOld()
  '' 'dim RESULT as string=join(codeLinkDATA,"<³³>")
  '' dim grid as iGrid=iGrid3
  '' dim rowIndex as integer=grid.curRow.index
  '' dim ir as iGRow
  '' 
  '' ir = grid.rows.insert(rowIndex)
  '' 'ir = grid.rows.add()
  '' 
  '' with ir
  '' .cells(cNickname).value  = codeLinkDATA(3)   'lineContent
  '' .cells(cId).value        = codeLinkDATA(4)        'lineNumber.toString
  '' .cells(cFileSpec).value  = codeLinkDATA(1)        'lineNumber.toString
  '' '.cells(cNickname).backColor = ColorTranslator.FromHtml("#7A8599")
  '' '.cells(cNickname).foreColor = ColorTranslator.FromHtml("#fff")
  '' .cells(cNickname).backColor = ColorTranslator.FromHtml("#364359")
  '' .cells(cNickname).foreColor = ColorTranslator.FromHtml("#ccc")
  '' .ContentIndent= New TenTec.Windows.iGridLib.iGIndent(1, 0, 1, 0)
  '' end with
  '' saveGridData(grid)
end sub




function xxx_getCodeLinkOld() as string()
  '' setMonitorRef()
  '' dim codeEitor = ZZ.getActiveRTF.RTF
  '' dim lineContent as string = trim(codeEitor.Lines.Current.text)
  '' lineContent=replace(lineContent,vbNewLine,"")
  '' lineContent=replace(lineContent,"sub","",,,1)
  '' lineContent=replace(lineContent,"function","",,,1)
  '' lineContent=replace(lineContent,"private","",,,1)
  '' lineContent=replace(lineContent,"public","",,,1)
  '' lineContent=trim(lineContent)
  '' 'lineContent=replace(lineContent," ","_")
  '' dim lineNumber as integer = codeEitor.Lines.Current.number
  '' 
  '' dim activeTab         = ZZ.getActiveTab()
  '' dim activeTabType     = ZZ.getActiveTabType()
  '' dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  '' 
  '' dim tagDATA(10) as string
  '' tagDATA(0)="index"
  '' tagDATA(1)=mid(ActiveTabFilespec,6)
  '' tagDATA(2)="subFunc"
  '' tagDATA(3)=lineContent
  '' tagDATA(4)=lineNumber.toString
  '' tagDATA(5)="reserve1"
  '' 'add (newTag)
  '' 
  '' ''msgbox(lineContent)
  '' monitorAdd("§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§")
  '' monitorAdd(lineContent)
  '' monitorAdd("§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§")
  '' '' Dim name() = Split(e.Sender.Name+"_", "_")
  '' '' dim index as string=name(2)
  '' '' dim id ="lab_navLab_"+index+"<--"
  '' '' monitorAdd("id: "+id)
  '' '' dim label=ref.element("lab_navLab_"+index)
  '' '' label.text=lineContent
  '' '' label.tag=newTag
  '' '' monitorAdd(e.sender.text)
  '' '' 
  '' 
  '' return tagDATA
end function











'-->
'--> M U E L L E I M E R  




'' sub XXX_saveIndexState(dict as Dictionary(Of String, String))
''   '!!! ungueltige Eintraege vernichten fehlt noch
''   dim key as string
''   dim OUT(1000) as string 
''   dim counter as integer=0
''   dim trenn as string="|||"
''   for each key in dict.keys
''     'trace dict.item(key),key
''     if dict.item(key)=">" then
''       OUT(counter)=dict.item(key)+trenn+key
''       counter=counter+1
''     end if
''   next
''   redim preserve OUT(counter)
'' 
''   'dim RESULT=join(OUT,vbNewLine)
''   ' monitorAdd(RESULT)
''   dim myPath as string =glob_defaultPath
''   dim fileSpec as string=globIndexListStateFileSpec
''   Directory.CreateDirectory(myPath)
''   IO.File.WriteAllLines(fileSpec, OUT)
'' end sub
'' 

'' sub XXX_readIndexState(dict as Dictionary(Of String, String))
''   dim myPath as string = glob_defaultPath
''   dim fileSpec as string = globIndexListStateFileSpec
''   Dim fileContent = IO.File.ReadAllText(fileSpec)
''   dim lines() as string=split(fileContent,vbNewLine)
''   dim fields() as string
''   dim item as string
''   dim key as string
''   dim marker as string=">"
''   dim trenn as string="|||"
'' 
''   dim max as integer=lines.length
''   dim i as integer
''   for i =0 to max-1
''     item=lines(i)+trenn+trenn+trenn
''     fields=split(item,trenn)
''     key=fields(1)+trenn+fields(2)
''     if globCodeListState.ContainsKey(key) then
''       globCodeListState.item(key)=marker
''     else
''       globCodeListState.add(key,marker)
''     end if
''   next
'' end sub
'' 


'-->  --> LogIn

sub login_mouseClick(e) 
  logIn()
End Sub


sub logIn()
  dim userName as string
  dim passWord as string
  dim RESULT as boolean
  userName="es"
  passWord="sandus"
  
  RESULT = doLogin( userName, passWord)
  trace "logIn", RESULT
end sub  


Function doLogin(ByVal userName As String, ByVal passWord As String) As Boolean
  Dim postData As String = "u=" + userName + "&p=" + passWord
  ''Dim RESULT = zz.ideHelper.TwAjax.postUrl("http://teamwiki.net/php/vb/app_login2.php?", postData)
  Dim RESULT = zz.http_post("http://teamwiki.net/php/vb/app_login2.php?", postData)
  Dim lines() = Split(RESULT, vbNewLine)
  ReDim Preserve lines(4)
   If lines(0) = "Login OK" Then
    twLoginuser = userName : twLoginPass = passWord : twLoginFullname = lines(3)
    '' glob.para("twLoginData") = glob.RC4StringEncrypt(userName + ":" + passWord, kData)
    twSessID = lines(2)
    Return True
  Else
    MsgBox("Ungueltige Login-Daten!", MsgBoxStyle.Exclamation)
    Return False
  End If
End Function



'--> ...
sub readData_mouseClick(e) 
  dim url as string
  dim cookies as string
  dim RESULT as string
  url="http://secure.teamwiki.net/q/profileservice.php"
  url=url+"?m=mydocs_getfile"
  url=url+"&type=out"
  url=url+"&u=ugb"
  url=url+"&f=shared_docs/umfrage_ergebnisse/"
  url=url+"&n=results_001.txt"
  cookies="twnetSID=" + twSessID
  
  RESULT=ZZ.http_get(url, cookies)
  panelRef.element("txt_outMonitor").text=RESULT
End Sub





sub test1_mouseClick(e)
  printLine 11, "xxx", "12345"
  dim fileSpec as string =   panelRef.element("txt_fileSpec").text
  readCodeFile(fileSpec)

exit sub
  ' ....webFile lesen
  
  '' '...testfehler
  '' dim y as object
  '' y.kjdfkdjskfj

  ''msgBox ("xxx-test1")
  'dim strAppName: strAppName="scriptIDE_test_es"
  dim strAppName: strAppName="projektmanager"
  dim strFileName: strFileName="gridData.grid"
  dim strContent as string
  strContent=ZZ.TwReadFile(strAppName, strFileName) 
  panelRef.element("txt_outMonitor").text=strContent
  IGrid1.cols.Count = 25

  Igrid_put(IGrid1, strContent, vbNewLine, "|")
  IGrid1.rows.count=20  'provisorisch
end sub


sub readCodeFile(filespec)
  printLine 11, "xxx", "12345"
  ' dim fileSpec as string =   panelRef.element("txt_fileSpec").text
  dim Result as string = zz.fileGetContents(filespec)
  panelRef.element("txt_inBox").text = Result
end sub





sub test2_mouseClick(e) 'webSpeichern
  'dim strAppName: strAppName="scriptIDE_test_es"
  dim strAppName: strAppName="projektmanager"
  dim strFileName: strFileName="gridData.grid"
  dim strContent as string: strContent=""
  dim errMes as string
  Igrid_get(IGrid1, strContent, vbNewLine, "|")
  errMes=ZZ.TwSaveFile(strAppName, strFileName, strContent)
  panelRef.element("txt_outMonitor").text=errMes
end sub


sub test3_mouseClick(e) 'Testdaten anzeigen
  'dim strAppName: strAppName="scriptIDE_test_es"
  dim strAppName: strAppName="projektmanager"
  dim strFileName: strFileName="gridData.grid"
  dim strContent as string: strContent=""
  dim errMes as string
  strContent="1111111111||1252173651038-es||DEFAULT-ITEM||145|131|324|22||||0px  solid #fff|||#DDD|#000|| [_] ||||||"
  strContent=strContent+vbNewLine+"1111111112||1252173651038-es||DEFAULT-ITEM||145|131|324|22||||0px  solid #fff|||#DDD|#000|| [_] ||||||"
  strContent=strContent+vbNewLine+"1111111113||1252173651038-es||DEFAULT-ITEM||145|131|324|22||||0px  solid #fff|||#DDD|#000|| [_] ||||||"
  ''IgridHelper.Igrid_get(IGrid1, strContent, vbNewLine, "|")
  ''errMes=ZZ.TwSaveFile(strAppName, strFileName, strContent)
  Igrid_put(IGrid1, strContent, vbNewLine, "|")
end sub






  '--> --- Right Pane - 2
  '' with pRight1
  ''   .resetControls (5,5)
  ''  
  ''  
''  '--> icon
''  
''     .insertX = 10 :.insertY = 7
''     .addIcon("appPicture", "http://es.teamwiki.net/docs/icons/insert-object.png" )
'' 
''    .insertX = 160 :.insertY = 4
''    ''Function addLabel(strID,  strText,   bgColor fgColor  intLeft = -1,  intTop = -1, intWidth = -1, intHeight = -1) '@@-
''    ''                  el=.addLabel  ("lab01", "Auswertung   UGB - Onlinefragebogen" ,  ,"#ffffff",,,380,31) :makeJumboLabel(el)
''                      el=.addLabel  ("lab01", "" ,  ,"#ffffff",,,380,31) :makeJumboLabel(el)
''     .insertX = 560 : el=.addLabel  ("curRow",         "zeile" ,  ,"#ffffff" , , ,55 ,31 ) :makeJumboLabel(el)
''     .insertX = 620 : el=.addLabel  ("selectionCount", "sel" ,  ,"#ffffff" , , ,88 ,31 ) :makeJumboLabel(el)
''  
''     'el.width=200
''     'el.height=50
''     
''      .BR  '-------------------------------------------------- '-->inBox
''     .insertX = 160 :.insertY = 40
''      el=.addLabel  ("label10", "Datei(lesen)" ,    ,"#555" )
'' 
''     .addTextbox ("fileSpec"   ,  -2         , "..."  , 0  , "#aaa")
''        'panelRef.element("txt_fileSpec").Text="C:\yPara\scriptIDE\scriptClass\es_codeList.vb"
''        panelRef.element("txt_fileSpec").Text="C:\yPara\scriptIDE\scriptClass\ugb_emailMischer.vb"
''      .BR  '-------------------------------------------------- '-->inBox
''     .insertX = 160 :.insertY = 65
''     .addButton  ("startCodeList"        , "startCodeList"     , "#6f6")
''      .addButton  ("test1"      , "Lesen"      , "#ccc")
''     .addButton  ("test3"      , "test2"      , "#ccc")
''     .addButton  ("test2"      , "test3"      , "#ccc")
'' 
''     .br  '--------------------------------------------------
''     .insertX = 160 :.insertY = 100
''     check2 = New System.Windows.Forms.CheckBox
''     with check2
''     ''check1.Location = New System.Drawing.Point(360, 10)
''       .Location = New System.Drawing.Point(160, 100)
''       .Size = New System.Drawing.Size(60, 19)
''       .Text = "auto"
''       .UseVisualStyleBackColor = True
''       .visible=true
''      pRight1.Controls.Add(check2)
''     end with
'' 
'' 
'' '-->  textbox
'' .BR  '-------------------------------------------------- '-->inBox
''     .insertX = 5 :.insertY = 140
''     .addTextbox ("inBox" ,  -2         , "inBox"    , 0  , "#FFFF99", , , "multiline=190")
''        panelRef.element("txt_inBox").WordWrap=false
''        panelRef.element("txt_inBox").scrollbars=System.Windows.Forms.ScrollBars.Vertical
''        
'' '--> debugbereich
'' .BR  '--------------------------------------------------
''    .insertX = 11 :.insertY = 570
''    .addTextbox ("outMonitor" ,  -2         , "debug...:"    , 50  , "#afa", , , "multiline=80")
''        panelRef.element("txt_outMonitor").WordWrap=false
''     .br  '--------------------------------------------------
''     .addButton  ("clear"        , "clear"     , "#f66")
''     .addButton  ("clearGrid"    , "clearGrid" )
''     .addButton  ("login"        , "logIn" )
''     .addButton  ("readData"     , "Umfrage lesen" )
''     .addTextbox ("inummer"      ,  200     , "iNummer:"    , 50  , "#aaa")
''     .br  '--------------------------------------------------
''  
'' 
'' 
''     ''trace "appPath", application.executablePath
''      dim exePath as string
''      exePath=application.executablePath
''      if exePath.toLower.endsWith("rtftabmdi.exe") then
''        formRef.Text = "--> SCRIPT: ------- codeList"
''      else
''        '' formRef.Text = "web-READER.vb   (c)es, mw"
''      end if
''      trace "AutoStart", "FERTIG"
''   End With
''   
''


