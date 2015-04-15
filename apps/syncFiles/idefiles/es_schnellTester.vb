dim appIcon as string ="http://es.teamwiki.net/docs/icons/package_settings.png"
'dim appIcon as string ="http://es.teamwiki.net/docs/icon24/emblem-noread.png"
'dim appIcon as string ="http://es.teamwiki.net/docs/icon24/emblem-noread.png"
'dim appIcon as string ="http://es.teamwiki.net/docs/icon24/emblem-noread.png"
'dim appIcon as string ="http://es.teamwiki.net/docs/icon24/emblem-noread.png"

'' es_schnellTester


'-->
'--> C O N F I G - G L O B A L 


'--> Config


#Para DebugMode intern
#Para SilentMode true

shared winId as string ="es_schnellTester.testWin"



'--> glob

Dim WithEvents Timer1 As System.Windows.Forms.Timer

public globCurTabFileSpec as string
public myGlobId as string

'NEU...
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

  

'' #reference Microsoft.visualbasic.dll
'' #imports microsoft.visualbasic

#reference TenTec.Windows.iGridLib.iGrid.dll
'' #imports Tentec.Windows.Igridlib



public WithEvents FormRef As Form
public WithEvents ref As object'scriptedPanel
public withEvents myWin2 as object
dim myImg as object
public WithEvents IGrid1 As Igrid

public shared toolBar As object'scriptedPanel
public shared statBar As object'scriptedPanel
public shared container1 As object'scriptedPanel
public toolBarColor as string

public iconList() as string

''dim withEvents scin As ScintillaNet.Scintilla


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


Sub GetFormRef()
  'If ref IsNot Nothing Then Exit Sub
  'ref = ZZ.IDEHelper.CreateDockWindow(winID, 4)
  ref = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)'  : err.number=0
  formRef = ref.form
  formRef.text="Schnell-Tester"
  CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = true
End Sub



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
  'AddHandler formRef.Resize, AddressOf Form_Resize
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
  'RemoveHandler formRef.Resize, AddressOf Form_Resize
  'RemoveHandler scNet.MouseDown, AddressOf scNet_MouseDown
  'RemoveHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  'trace "globRemoveHandler","DONE"
end sub




Sub Frm_FormClosing(Sender As Object,e As System.Windows.Forms.FormClosingEventArgs) Handles FormRef.FormClosing
  '' glob.saveFormPos(FormRef)
  '' glob.saveParaFile()
  trace "formClosing" , "xxx"
  'globRemoveHandler()
End Sub


Sub onTerminate()
  'trace "!!! onTerminate", "!!! es_codeList"
  globRemoveHandler()
  'If VLC1.playlist.isPlaying Then VLC1.playlist.stop()
End Sub



sub makeJumboLabel(el)
    el.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    ''el.Size = New System.Drawing.Size(117, 37)
    el.backColor=ColorTranslator.FromHtml("#777")
    el.AutoSize = false
    el.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
end sub




'-->
'--> R E S I Z E 

Sub Form_Resize(sender as Object, e as EventArgs)
  on_Resize_Controls()
End Sub



Sub on_Resize_Controls()
  '' if skipResizeEvent = true then exit sub
  '' Igrid1.cols(5).width = pLeft1.Width +2
  '' Igrid1.Height = pLeft1.Height - 66
  '' Igrid1.Height = 170
  '' ref.element("test2").top=ref.Height - 28
End Sub



'-->
'--> T I M E R



sub startTimer
  ''msgbox("startTimer")
  Timer1 = New System.Windows.Forms.Timer()
  timer1.Interval = 555
  timer1.Enabled = True
end sub



 : Sub onTimerEvent(ByVal sender As Object, ByVal e As System.EventArgs)
 on error resume next
 'exit sub
  static lastFileSpec as string=""
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  if lastFileSpec<>ActiveTabFilespec then
    dim codeEitor = ZZ.getActiveRTF.RTF
    
    setMonitorRef()
    
    lastFileSpec = ActiveTabFilespec
    globCurTabFileSpec=ActiveTabFilespec
    '' monitorAdd(myGlobId)
    ''### monitorAdd(appGlob.para("codeListId"))
    '' if myGlobId <> appGlob.para("codeListId") then
    ''   monitorAdd("...VERALTET")
    ''   onTerminate()
    ''   exit sub
    '' end if
    dim name as string=My.Computer.FileSystem.GetName(lastFileSpec)
    ref.element("lab_cmdTitleLab").text=name
   
    refreshAppIcon(codeEitor)
  end if  
End Sub


:sub refreshAppIcon(codeEitor)
  on error resume next
  'dim codeEitor = ZZ.getActiveRTF.RTF
  dim codeText as string=codeEitor.text
  dim icon as pictureBox
  dim temp as string = codeText.toLower
  dim LOWER() as string: LOWER=split(temp,vbNewLine)
  dim SOURCE() as string: SOURCE=split(codeText,vbNewLine)
  icon=ref.element("pic_cmdScriptAppIcon")
  monitorAdd(LOWER(0))
  monitorAdd(LOWER(1))
  monitorAdd(LOWER(2))
  if LOWER(0).contains("appicon ")then
    dim apo as string=""""
    dim part() as string=split(SOURCE(0),apo)
    dim img as string=part(1)
    monitorAdd("-->"+img+"<--")
    icon.image=ZZ.GetImageCached(img)
    icon.visible=true
  else
    icon.visible=false
  end if
end sub




'''#####################################################


'-->
'--> T E S T 

sub test1()
  onSetLastEventHandlerPara("cmdInsertEventHandler", "bbbbbbbbb", true)
  'msgBox("OK - 1")
  monitorAdd("111111")
  monitorAdd("111111")
  monitorAdd("111111")
  monitorAdd("111111")
  static counter as integer
  counter=counter+1
  statustext(counter.toString)
  'monitorAdd("es_iconBar.mainWin","...ich bin NACHRICHT")
  'zz.traceClear()
  'zz.printLineReset()
  'listDocumentPane()

exit sub
  
  'msgBox("OK - 1")
  monitorAdd("schnellTester-2222222","...ich bin NACHRICHT")
  zz.traceClear()
  zz.printLineReset()
  'listDocumentPane()
end sub


#Reference siaIDEMain.dll

sub test2()
  ''msgBox("OK - 2")
  monitorAdd("sub test2()")
  dim mainWin as ScriptIDE.Main.frm_main
  mainWin = ZZ.IDEHelper.MainWindow
  dim control as control
  monitorAdd(mainWin.controls.count)
  for each control in mainWin.controls
    monitorAdd(control.name)
  next
  mainWin.StatusStrip1.visible = false
  mainWin.pnlTitlebar.visible = false
  Dim menu As Menustrip
'  menu = mainWin.Controls.Find("Menustrip1",True)
  
  
  'mainWin.MenuStrip1.visible=false
  'mainWin.toolstripContainer1.TopToolStripPanel.visible=false
  'mainWin.DockPanel1.parent.top=59
 
  monitorAdd(mainWin.DockPanel1.parent.top)
  monitorAdd(mainWin.DockPanel1.height)
  monitorAdd(mainWin.height)
  'mainWin.ToolStripContainer1.TopToolStripPanelVisible = False

end sub

sub test22()
  static counter
  counter=counter+1
  statustext(counter.toString)
  '' dim Workbench.instance = ZZ.IDEHelper.MainWindow
  '' dim OUT(333) as string
  '' dim counter as integer=0
  '' For Each win As Object In Workbench.Instance.DockPanel1.Contents
  '' 
  '' 
  ''     Dim lvi = lvOpenedWins.Items.Add(win.Text)
  ''     lvi.subitems.add(win.GetPersistString())
  ''     lvi.subitems.add(win.DockState)
  ''     lvi.subitems.add(win.GetType.Name)
  ''     lvi.tag = win
  ''   Next
  '' monitorAdd("3333333333333333333333")
  '' 'listDocumentPane(7, true)
  '' 'sendMesTest2()
end sub





sub test3()
  monitorAdd("333333")
  monitorAdd("333333")
  monitorAdd("333333")
  monitorAdd("333333")
 ' msgBox("OK - 3")
  exit sub
  
  
  listDocumentPane(7, false)
  listDocumentPane(8, false)
  listDocumentPane(9, false)
  listDocumentPane(10, false)

  'sendMesTest3()
end sub




'-->
'--> --------------------------------------------------------



'-->
'--> ==> A U T O S T A R T <==

: Sub AutoStart()
  on error resume next
  'myWin2 = ZZ.scriptClass("es_schnellTester.testWin")
  'myWin2.setParent(me)
  'zz.traceClear()
  'zz.printLineReset()
  
  dim id as string
  id="ToolBar|##|tbScriptWin|##|es_direktfenster.mainwin"             : ToggleDockWindowFromId(id, 0)
  id="ToolBar|##|tbScriptWin|##|es_testLabor.mainWin"                 : ToggleDockWindowFromId(id, 0)
  id="ToolBar|##|tbScriptWin|##|es_iconBar_B.mainWin"                 : ToggleDockWindowFromId(id, 0)
  id="ToolBar|##|tbScriptWin|##|es_internsuche.mainwin"               : ToggleDockWindowFromId(id, 0)
  id="ToolBar|##|tbScriptWin|##|es_snippetManager.mainWin"            : ToggleDockWindowFromId(id, 0)
  
  GetFormRef()
  with ref
    .resetControls ()
    '.activateEvents = "|IconMouseDown|   |TextboxKeyDown|"

    ''toolBar     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=75)
    container1  =.addPanel("container1"   , DockStyle.Fill)
    toolBar     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
    'statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
    statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=28)
    toolbar.visible=true
    'container1.top=112
    toolBar.resetControls  (10,5,1)
    toolbar.height=32
    ''toolbar.BackColor = ColorTranslator.FromHtml(toolBarColor)
    toolbar.BackColor = ColorTranslator.FromHtml("#fff")
 
    statBar.resetControls  (10,5,1)
    statBar.BackColor = ColorTranslator.FromHtml(toolBarColor)
    statBar.visible=false

    container1.resetControls  (10,5,1)
    container1.BackColor = ColorTranslator.FromHtml("#eef")
  end with

  'With ref
  with container1
     dim el as object
    .resetControls (4,6)
    .backColor=ColorTranslator.FromHtml("#43526F") '("#3E6EC3")' 
    '.backColor=ColorTranslator.FromHtml("#68738A") '("#3E6EC3")' 

    '-->icon
    .OffsetX = 4 ::.insertY = 23
    .OffsetX = 4 ::.insertY = 4
    el=.addIcon("cmdScriptAppIcon", "http://es.teamwiki.net/docs/icons/128email.png" )
    'el=.addIcon("cmdScriptAppIcon", "http://icons3.iconfinder.netdna-cdn.com/data/icons/humano2/72x72/apps/gnome-monitor.png" )
    'el.visible=false
    el.autosize=false
    el.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
    el.height=65
    el.width=65
    
   '-->test / toolbar
   .OffsetX = 0 ::.insertY = 75
    '.addButton(id,txt,bgColor,Left,Top,Width,Height,iconUrl,flags,handler)
     'el=.addButton ("aufZu"       , "+ / ---"      , "#FFD350" , , ,44,20): .OffsetX = .OffsetX +44:
     el=container1.addButton("cmdCallByName"    , "callByName"   , "#364359" , , ,77,20): .OffsetX = .OffsetX +77: el.foreColor=ColorTranslator.FromHtml("#fff") 
     el=container1.addButton("cmdTest_1"    , "test_1"       , "#364359" , , ,44,20): .OffsetX = .OffsetX +44: el.foreColor=ColorTranslator.FromHtml("#fff")
     el=container1.addButton("cmdTest_2"    , "test_2"       , "#364359" , , ,44,20): .OffsetX = .OffsetX +44: el.foreColor=ColorTranslator.FromHtml("#fff") 
     el=container1.addButton("cmdTest_3"    , "test_3"       , "#364359" , , ,44,20): .OffsetX = .OffsetX +44: el.foreColor=ColorTranslator.FromHtml("#fff") 
     el=container1.addButton("openTabs"     , "allTabs"      , "#364359" , , ,60,20): .OffsetX = .OffsetX +75: el.foreColor=ColorTranslator.FromHtml("#fff") 

   .OffsetX = 77 ::.insertY = 0' 20 '70
     el=container1.addButton ("cmdHistory"     , "History"     , "#364359" , , ,60,20): .OffsetX = .OffsetX +60: el.foreColor=ColorTranslator.FromHtml("#fff")
     el=container1.addButton ("cmdSearch"      , "Suchen"      , "#364359" , , ,60,20): .OffsetX = .OffsetX +60: el.foreColor=ColorTranslator.FromHtml("#fff")
     el=container1.addButton ("cmdBookmarks"   , "Bookmarks"   , "#364359" , , ,70,20): .OffsetX = .OffsetX +70: el.foreColor=ColorTranslator.FromHtml("#fff") 

   '--> ... LABELS
   '.insertX = 75  :.insertY = 20 '28
   .insertX = 80 :.insertY = 17 '28
    dim labelText as string ="es_codeList"
    el=.addLabel  ("cmdTitleLab", labelText ,  ,"",,,333,32) :
    makeJumboLabel(el)
    el.backColor=ColorTranslator.FromHtml("#43526F")
    el.foreColor=ColorTranslator.FromHtml("#ccc"): 
    el.TextAlign = System.Drawing.ContentAlignment.MiddleLeft

   .insertX = 80 :.insertY = 38 '28
    labelText ="eventHandler"
    el=.addLabel  ("cmdInsertEventHandler", labelText ,  ,"",,,333,35) :
    'makeJumboLabel(el)
    el.backColor=ColorTranslator.FromHtml("#43526F")
    el.foreColor=ColorTranslator.FromHtml("#999"): 
    el.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    el.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))





    .insertX = 11 : .insertY = 22
    .OffsetX = 5 ::.insertY = 22
    ' .BR (22)'--------------------------------------------------
    '.addLabel(ID, Text, bgColor, fgColor               ,Left,Top,Width,Height
    .insertX = 75 :.insertY = 114
    'el=.addLabel ("test444a"     , "appIcon ?"      ,"#696E73", ,5   ,24 ,42   ,16): el.textAlign=System.Drawing.ContentAlignment.MiddleCenter: el.foreColor=ColorTranslator.FromHtml("#bbb")
    ''el=.addLabel ("test444"     , "fileName"      ,"#84888C", ,49  ,24 ,-2   ,16):  el.textAlign=System.Drawing.ContentAlignment.MiddleCenter: el.foreColor=ColorTranslator.FromHtml("#fff") 
    'el=.addLabel ("test444"     , "fileName"       ,"#84888C", ,49  ,24 ,130  ,16):  el.textAlign=System.Drawing.ContentAlignment.MiddleCenter: el.foreColor=ColorTranslator.FromHtml("#fff") 
    .insertX = 6 :.insertY = 144
     myImg=.addIcon("appPicture", "http://es.teamwiki.net/docs/icons/package_settings.png" )



    .BR(23) '--------------------------------------------------
    .OffsetX = 76 ::.insertY = 146
    el=.addButton ("profile1"     , "10"    , "#BEC8C8", , ,40,20):el.autosize=false:el.height=20
    .BR(20) '--------------------------------------------------
    el=.addButton ("profile2"     , "100"   , "#BEC8C8", , ,40,20):el.autosize=false:el.height=20
    .BR(20) '--------------------------------------------------
    el=.addButton ("profile3"     , "1000"  , "#BEC8C8", , ,40,20):el.autosize=false:el.height=20
    .BR(20) '--------------------------------------------------

    .OffsetX = 115:.insertY = 146:
    el = .addTextbox ("speed1" ,  60  , "  0.000 " , 0,"#bbb" , , , )
    'ref.element    ("txt_speed1").text=code
    el = ref.element("txt_speed1")
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.borderStyle=0: el.foreColor=ColorTranslator.FromHtml("#fff"): el.backColor=ColorTranslator.FromHtml("#84888C")
    .BR(20) '--------------------------------------------------
    el = .addTextbox    ("speed2" ,  60  , "  0.000 " , 0,"#bbb" , , , )
    el = ref.element("txt_speed2")
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.borderStyle=0: el.foreColor=ColorTranslator.FromHtml("#fff"): el.backColor=ColorTranslator.FromHtml("#84888C")
    .BR(20) '--------------------------------------------------
    el = .addTextbox    ("speed3" ,  60  , "  0.000 " , 0,"#bbb" , , , )
    el = ref.element("txt_speed3")
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.borderStyle=0: el.foreColor=ColorTranslator.FromHtml("#fff"): el.backColor=ColorTranslator.FromHtml("#84888C")
    .BR(20) '--------------------------------------------------
    
    


    '' .BR(23) '--------------------------------------------------
    '' .OffsetX = 85 ::.insertY = 44
    '' el=.addButton ("test1"     , "test-1"  , "#BEC8C8"):el.autosize=false:el.height=20
    '' el=.BR(20) '--------------------------------------------------
    '' el=.addButton ("test2"     , "test-2"  , "#BEC8C8"):el.autosize=false:el.height=20
    '' .BR(20) '--------------------------------------------------
    '' el=.addButton ("test3"     , "test-3"  , "#BEC8C8"):el.autosize=false:el.height=20
    '' .BR (26)'--------------------------------------------------
   
   .OffsetX = 4: :.insertY = 211
    ' el=.addButton ("test4"     , ">"  , "#BEC8C8"):el.autosize=false:el.height=20: el.width=22
    el=.addLabel ("xxx66"     , "Func:"  , , "#BEC8C8"):el.autosize=false:el.height=20: el.width=35
   .OffsetX = 35:.insertY = 211
   '.OffsetX = 4: :.insertY = 211
   
    el = .addTextbox ("funcName" ,  -2         , "  Code"+vbNewLine+""    , 0,"#bbb" , , , )
    'ref.element("txt_codeTextarea").WordWrap=false
    'ref.element("txt_codeTextarea").text=code
    ref.element("txt_funcName").Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    ref.element("txt_funcName").backColor=ColorTranslator.FromHtml("#84888C")
    ref.element("txt_funcName").foreColor=ColorTranslator.FromHtml("#fff")
    ref.element("txt_funcName").borderStyle=0
 
   .BR ()'--------------------------------------------------
   .OffsetX = 4: :.insertY = 237
    ' el=.addButton ("test4"     , ">"  , "#BEC8C8"):el.autosize=false:el.height=20: el.width=22
    el=.addLabel ("xxx77"     , "   file:"  , ,"#BEC8C8"):el.autosize=false:el.height=20: el.width=35
   .OffsetX = 35 :.insertY = 237
   '.OffsetX = 4: :.insertY = 227
   
    el = .addTextbox ("fileName" ,  -2         , "  Code"+vbNewLine+""    , 0,"#bbb" , , , )
    'ref.element("txt_codeTextarea").WordWrap=false
    'ref.element("txt_codeTextarea").text=code
    ref.element("txt_fileName").Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    ref.element("txt_fileName").backColor=ColorTranslator.FromHtml("#84888C")
    ref.element("txt_fileName").foreColor=ColorTranslator.FromHtml("#ccc")
    ref.element("txt_fileName").borderStyle=0
    .BR (26)'--------------------------------------------------



  end with
  
  iniIconList()
  insertIconsFromIconList(toolBar)


  startTimer()

  globAddHandler()

end sub





'-->  ...

  Sub createOpenedTabList(DockPanel1)
      trace "--------------", "------------------------------ ???"
    Dim doc as object
    'doc = listDocumentPane(DockPanel1)
    
    If doc Is Nothing Then Exit Sub
    'Dim scrollPos = tbOpenedFiles.IGrid1.VScrollBar.Value
    'tbOpenedFiles.IGrid1.Rows.Clear()
    'tbOpenedFiles.IGrid1.Rows.Count = doc.Contents.Count
    dim i 
    For i = 0 To doc.Contents.Count - 1
      trace "xxx",  doc.Contents(i).DockHandler.TabText
      'tbOpenedFiles.IGrid1.Cells(i, 0).Value = doc.Contents(i).DockHandler.TabText
      'tbOpenedFiles.IGrid1.Rows(i).Tag = DirectCast(doc.Contents(i).DockHandler.Form, DockContent).GetPersistString()
      'If doc.Contents(i) Is MAIN.ActiveMdiChild Then tbOpenedFiles.IGrid1.SetCurRow(i)
    Next
    'tbOpenedFiles.IGrid1.VScrollBar.Value = scrollPos
  End Sub





:Function listDocumentPane(Pane, showState) As object 'DockPane
  dim mainWin as object
  mainWin = ZZ.IDEHelper.MainWindow
  dim i as integer
  dim isMdiChild as boolean
  dim tag as object
  'dim DockState as object
  dim DockContent as WeifenLuo.WinFormsUI.Docking.DockContent

  dim panel1 as object = mainWin.DockPanel1
  dim p as object
  dim dockHandler as object
  For Each p In panel1.Panes
    For i = 0 To p.Contents.Count - 1
      isMdiChild=false
      dockHandler = p.Contents(i).DockHandler
      if p.Contents(i) Is mainWin.ActiveMdiChild Then isMdiChild=true
      ''if p.DockState = Pane or p.DockState = 7 or p.DockState = 8 or p.DockState = 9 or p.DockState = 10 then 
     if p.DockState = Pane  then 
        'p.show()
        'p.visible=true
       
        'p.parent.visible= not showState
        p.parent.visible= not p.parent.visible
        exit function
      end if
      
      ''trace p.DockState,  isMdiChild.toString+"  "+dockHandler.isHidden.toString+"  "+ dockHandler.TabText
      ''monitorAdd(p.DockState,  isMdiChild.toString+"  "+dockHandler.isHidden.toString+"  "+ dockHandler.TabText)
     
     
     ''tag=DirectCast(p.Contents(i).DockHandler.Form, DockContent).GetPersistString()
      'tbOpenedFiles.IGrid1.Rows(i).Tag = DirectCast(doc.Contents(i).DockHandler.Form, DockContent).GetPersistString()
      ''trace p.DockState,  isMdiChild.toString+"  "+p.Contents(i).DockHandler.TabText
      'if dockHandler.isFloat = true  then 
      'end if
      
      'tbOpenedFiles.IGrid1.Cells(i, 0).Value = doc.Contents(i).DockHandler.TabText
      'tbOpenedFiles.IGrid1.Rows(i).Tag = DirectCast(doc.Contents(i).DockHandler.Form, DockContent).GetPersistString()
      'If doc.Contents(i) Is MAIN.ActiveMdiChild Then tbOpenedFiles.IGrid1.SetCurRow(i)
    Next
    '' If p.DockState = 7 'DockState.Document Then
    ''    Return p
    ''  End If
  Next
End Function






'-->
'--> E V E N T S



Sub onWebBrowser1(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! 'onWebBrowser1': ...in Arbeit")
  'msgBox("'onWebBrowser1': " + "...Coming soon ")
  '...
  'dim exe as string="C:\yPara\scriptIDE4\compiledScripts\es_webbrowser3\es_webbrowser3.exe"
  dim exe as string="C:\yPara\scriptIDE\compiledScripts\es_webbrowser3\es_webbrowser3.exe"
  dim Para as string="http://www.iconfinder.com/icondetails/17829/24/"
  : ZZ.shellExecute(exe, para) : Err.Clear
  
End Sub






sub txt_speed1_KeyDown(e)
  trace "txt_speed1_KeyDown(e)",e.keyString
end sub



Sub onButtonEvent(e)
  monitorClear
  setMonitorRef()
  errorText("")
  printLine 11, "" , e.Sender.Name
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim tag as string = e.sender.tag.toString
  dim tagDATA()= Split(tag, "<³³>")
  dim action as string =name(1)
  monitorAdd("==============================================")
  monitorAdd("Sender.Name: " + e.Sender.Name+"<--")
  monitorAdd("")
  monitorAdd("      action: " +action+" <==")
  monitorAdd("     tagPara: " +tag+" <<<")
  monitorAdd(" MouseButton: " +e.MouseButton+" <<<")
  monitorAdd("   ClassName: " +e.ClassName+" <<<")
  monitorAdd(" ControlType: " +e.ControlType+" <<<")
  monitorAdd("      MouseX: " +e.MouseX.toString+" <<<")
  monitorAdd("==============================================")

  Select Case action.toLower
    '' case "togglepane"      : onTogglePane         (e, tagDATA)
    '' case "url"             : onNavigateUrl        (e, tagDATA)
    '' case "cmdoptions"      : onCmdOptions         (e, tagDATA)
    '' case "cmdopenexplorer" : onCmdOpenExplorer    (e, tagDATA)
    '' 
    '' Case "toggle"             : onToggleDockWindow   (e, tagDATA)
    '' Case "cmdtoggleweb"    : oncmdToggleWeb       (e, tagDATA)
    Case "saverun"            : onSaveRun             (e, tagDATA)
    '' Case "test1"           : ontest1               (e, tagDATA)
    Case "exec"               : onExecute             (e, tagDATA)
    '' Case "exe"             : onExe                 (e, tagDATA)
    '' case "mnuexe"          : onExe                 (e, tagDATA)
    '' Case "nav"             : onNavigateUrl         (e, tagDATA)
    '' Case "toggleimg"       : onToggleImg           (e, tagDATA)
    '' Case "checkbox"        : onToggleImg           (e, tagDATA)
    '' case "createIconList"  : onCreateIconList      (e, tagDATA)
    '' 'Case "expand"         : onToggleDockWindow    (e, tagDATA)
    '' 'Case "expand"         : toogleToolbarItems    (e, tagDATA)
    '' 'Case "switch"         : onSwitch              (e, tagDATA)
    '' 
    '' Case "cmdtoolbarman" :onCmdToolbarMan          (e, tagDATA)
    '' Case "cmdiconman"    :onCmdIconMan             (e, tagDATA)
    '' Case "item"            '... nichts weiter machen
    '' Case "caption"       '... nichts weiter machen
    Case else         
      callCmdByName(e)
     '' dim errMes as string = "ERR: onButtonEvent(e): '"+name(1)+"' ...Typ nicht erkannt: "
     ''  errorText(errMes)
     ''  monitorAdd("! === ! === ! === ! === ! === ! === ! === ! === ! === ")
     ''  monitorAdd(errMes)
     ''  monitorAdd("! === ! === ! === ! === ! === ! === ! === ! === ! === ")
  End Select
  
  'ref auf toolStripPanel über parent...
  'dim id2 ="btn_expand_01"
  'dim el =  ref.element(id2)
  'dim el =  e.sender
  'onTimerEvent()
  'checkDisplayState(el)
End Sub


Sub onButtonEvent2(e)
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
  MonitorAdd("======================== onLabelEvent")
  MonitorAdd("SenderFullName: " + e.Sender.Name)
  MonitorAdd("___MouseButton: " + e.MouseButton)
  MonitorAdd("________Action: " + action)
  MonitorAdd("======================== onLabelEvent")
  
    callCmdByName(e)

end sub



'--> ... --> AUSLAGERN
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
'--> E V E N T S - 2


Sub toggleTab(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  static state as boolean =false
  state= not state
  dim fileSpec="loc:/C:/yPara/scriptIDE4/scriptClass/es_kommentar.txt"
    dim tab = ZZ.IDEHelper.NavigateFile(fileSpec)
    if state=true then
      tab.hide
    else
      tab.show
    end if
End Sub



Sub cmdCallByName(e)  ' ...
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  monitorAdd(ActiveTabFilespec)
  dim fileSpec as string
  fileSpec="loc:/C:/yPara/scriptIDE4/scriptClass/es_schnellTester.vb"
  'fileSpec="ftp:/teamwiki.net/httpdocs/static/js/scriptLib01.js"
 
  dim tab = ZZ.IDEHelper.NavigateFile(fileSpec)

End Sub


Sub cmdInsertEventHandler(e)  ' ...
  statustext("!!! 'cmdInsertEventHandler': ...in Arbeit")
  'msgBox("'cmdInsertEventHandler': " + "...Coming soon ")
  if globLastEventHandlerFound = false then
    monitorAdd("...insertEventHandler")
  else 
    monitorAdd("...navigate-1: "+globLastEventHandlerEventName)
    monitorAdd("...navigate-2: "+globLastEventHandlerScriptClass)
  end if



  '' globLastEventHandlerEventName = evName
  '' globLastEventHandlerScriptClass = scriptClass
  '' globLastEventHandlerFound = found
  '' dim el =ref.element("lab_cmdInsertEventHandler")
  '' 

End Sub


Sub togglePanel(e)  ' ...
  toggleDockPanel(8)'left
End Sub

Function toggleDockPanel(intPanel as integer, optional forceState as integer=-1)
  static state as boolean = true: 'state = not state
  state = not state
  if forceState=0 then state=false
  if forceState=1 then state=true
  dim mainWin as object = ZZ.IDEHelper.MainWindow
  ' mainWin.xxxxxxxxxx    '''FEHLER-TEST
  dim Panel as object = mainWin.DockPanel1
  dim p as object: dim i as integer
  For Each p In Panel.Panes
     For i = 0 To p.Contents.Count - 1
       if p.DockState = intPanel  then 
         'p.parent.visible = not p.parent.visible
         p.parent.visible = state
         exit Function
       end if
     Next
  Next
End Function


Sub cmdSearch(e)  ' ...
  dim id ="ToolBar|##|tbScriptWin|##|es_internsuche.mainwin"
  dim toolWindow=ZZ.getDockPanelRef(id)
  if toolWindow is nothing then exit Sub
  dim curState as boolean  = toolWindow.visible
  dim force as integer=0
  if curState=false then force=1

  'hideMainTools()
  ToggleDockWindowFromId(id, force)
End Sub



Sub cmdHistory(e)  ' ...
  dim id ="ToolBar|##|tbScriptWin|##|es_codeList.history"
  dim toolWindow=ZZ.getDockPanelRef(id)
  if toolWindow is nothing then exit Sub
  dim curState as boolean  = toolWindow.visible
  dim force as integer=0
  if curState=false then force=1

  'hideMainTools()
  'force=1 ' ... nur in den vordergrund bringen
  ToggleDockWindowFromId(id, force)
End Sub


Sub cmdBookmarks(e)  ' ...
  dim id = "ToolBar|##|tbScriptWin|##|es_codeList.bookmarks"
  dim toolWindow=ZZ.getDockPanelRef(id)
  if toolWindow is nothing then exit Sub
  dim curState as boolean  = toolWindow.visible
  dim force as integer=0
  if curState=false then force=1

  'hideMainTools()
  ToggleDockWindowFromId(id, force)
End Sub


sub hideMainTools()
  dim id as string
  id = "ToolBar|##|tbScriptWin|##|es_internsuche.mainwin": ToggleDockWindowFromId(id, 0)
  id = "ToolBar|##|tbScriptWin|##|es_codeList.bookmarks": ToggleDockWindowFromId(id, 0)
  'id = "ToolBar|##|tbScriptWin|##|es_codeList.history": ToggleDockWindowFromId(id, 0)
end sub


Sub toggle(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! 'toggle': ...in Arbeit")
  'msgBox("'toggle': " + "...Coming soon ")
  '...
  dim tag as string = e.sender.tag.toString
  dim tagDATA()= Split(tag, "<³³>")
  dim tagPara as string = trim(tagDATA(1))
  monitorAdd("tagPara",tagPara)
  if tagPara = "" then exit sub

  dim id = (trim(tagPara))
  ToggleDockWindowFromId(id)

End Sub



sub onToggleDockWindow(e)
  dim tag as string = e.sender.tag.toString
  dim tagDATA()= Split(tag, "<³³>")
  dim tagPara as string = trim(tagDATA(1))
  monitorAdd("tagPara",tagPara)
  if tagPara = "" then exit sub

  dim id = (trim(tagPara))
  ToggleDockWindowFromId(id)
end sub


Sub toggleDebugWin(e) 
  with zz.scriptClass("es_testLabor")
    .toggleTraceWindow()
  end with
End Sub




Sub switch(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! 'toggle': ...in Arbeit")
  'msgBox("'toggle': " + "...Coming soon ")
  '...
  dim tag as string = e.sender.tag.toString
  dim tagDATA()= Split(tag, "<³³>")
  dim tagPara as string = trim(tagDATA(1))
  monitorAdd("tagPara",tagPara)
  if tagPara = "" then exit sub
  dim id = (trim(tagPara))
  dim curState as boolean= isToolWindowVisible(id)
  dim force as integer=0
  if curState=false then force=1
  hideOpenedSidebarPanels()
  ToggleDockWindowFromId(id,force)
End Sub

function isToolWindowVisible(id as string) as boolean
  isToolWindowVisible=false
  dim toolWindow=ZZ.getDockPanelRef(id)
  if toolWindow is nothing then exit function
  return toolWindow.visible
end function


sub hideOpenedSidebarPanels()
  'exit sub
  dim id as string
  id="ToolBar|##|tbScriptWin|##|es_snippetManager.mainWin"  :ToggleDockWindowFromId(id, 0)
  id="ToolBar|##|tbScriptWin|##|es_internSuche.mainWin"     :ToggleDockWindowFromId(id, 0)
  id="ToolBar|##|tbScriptWin|##|es_bookmarks2.mainWin"      :ToggleDockWindowFromId(id, 0)

end sub


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
  dim id3 as string="ToolBar|##|tbScriptWin|##|es_iconBar_B.mainWin"

  
  'newState=true
  if newState = true then
    toolWindow.show()
    toolWindow.visible =true
    toolWindow.parent.visible =true
    if id=id2 or id=id3 then toolWindow.parent.parent.visible =true   
  else
    toolWindow.hide()
    toolWindow.visible =false
    if id=id2 or id=id3 then toolWindow.parent.parent.visible =false
  end if
  
  return newState
end function
'' 
'' function ToggleDockWindowFromId(id as string, optional forceState as integer=-1)as boolean
''   dim toolWindow=ZZ.getDockPanelRef(id)
''   
''   '' exit function
''   if toolWindow is nothing then exit function
''   dim curState as boolean  = toolWindow.visible
''   
''   ''msgBox(curState)
''   dim newState as boolean= not curState
''   if forceState=0 then newState=false
''   if forceState=1 then newState=true
''   dim id2 as string="ToolBar|##|tbScriptWin|##|es_testLabor.mainWin"
''   ''monitorAdd(id)
''   ''monitorAdd(id2)
''   
''   'newState=true
''   if newState = true then
''     toolWindow.show()
''     toolWindow.visible =true
''     toolWindow.parent.visible =true
''     if id=id2 then toolWindow.parent.parent.visible =true   
''   else
''     toolWindow.hide()
''     toolWindow.visible =false
''     if id=id2 then  toolWindow.parent.parent.visible =false
''   end if
''   
''   return newState
'' end function
'' 
'' 

Sub cmdTitleLab(e)  ' ...
  dim id as string = "Addin|##|siaidemain|##|OpenedFiles"
  ToggleDockWindowFromId(id) 
End Sub


Sub cmdScriptAppIcon(e)  ' ...
  dim id as string = "ToolBar|##|tbScriptWin|##|es_testLabor.mainWin"
  ToggleDockWindowFromId(id) 
End Sub


Sub cmdTest(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  
  statustext("!!! 'cmdTest': ...in Arbeit")
  'msgBox("'cmdTest': " + "...Coming soon ")
  '...

  dim timerStart1 = GetTime()
  'dim tagPara as string = trim(tagDATA(1))
  'monitorClear()
  statustext_reset()
  ' monitorAdd("startTest")
 
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  ActiveTabFilespec=mid(ActiveTabFilespec,6)
  monitorAdd(ActiveTabFilespec)
  
  dim scriptClass = getScriptClassFromFileSpec(ActiveTabFilespec)
  monitorAdd("scriptClass.name: "+scriptClass.name)

  dim sc=scriptClass
  if sc is nothing then exit sub

  scriptClass.setMonitorRef()  '... falls autoStart noch nicht aufgerufen wurde
  'sc.test1()
  
  dim funcName="test"+index.toString
  trace "funcName", funcName
  monitorAdd("funcName: "+funcName)

  dim i as integer
  dim deltaTime as integer
  
  'CallByName(sc, funcName, Microsoft.VisualBasic.CallType.Method)
  
  Dim myType As Type = sc.GetType()
  :Dim myMethod As System.Reflection.MethodInfo = myType.GetMethod(funcName)
  monitorAdd("myMethod.toString: "+myMethod.toString+"<--")
  if myMethod.toString = "" then
    dim errMes="ERR: Sub '"+funcName+"' ...nicht gefunden"   '@@-
    statustext(errMes)
    monitorAdd(errMes)
    exit sub
  end if
  
  'monitorAdd("procName: "+myMethod.toString)
  dim args() 

  '' 
  '' 'where "strFnName" is your method name and "Args" is an array of parameter
  '' 'values (objects) for the method's arguments in the same order.
  '' 
  '' 
   
  '' if ERR.number<>0 then
  ''   ERR.number=0: dim errMes="ERR: Sub '"+funcName+"' ...nicht gefunden"  '@@-
  ''   statustext(errMes)
  ''   'printLine 11, "ERR", errMes
  ''   'trace         "ERR", errMes
  ''   'ref.ResumeLayout(False)
  ''   'ref.PerformLayout()
  ''   exit sub
  '' End if
  '' 
  
  
  monitorAdd("___________________________START")
  dim timerStart2 = GetTime()
  for i = 1 to 1' 100  ' 1000 '''1000000
    'msgbox(GetAsyncKeyState(19).toString)
    if GetAsyncKeyState(19) <> 0 then      '...PAUSE = abbrechen
      'msgbox("break")
      errorText("...Abgebrochen")
      monitorAdd("...Abgebrochen")
      exit for
    end if

    ' ACHTUNG: 'myMethod.Invoke' ist rund 7 mal so schnell(fast 10 mal schneller) !!!
    ' ACHTUNG: ... und ich kann heausfinden, ob es die proc überhaupt gibt (siehe weiter oben)
    '          ... damit komme ich auf circa 350 000 aufrufe einer leeren funktion
    '          ... bzw. rund 370 000 ohne checkAsyncKeyState
    '' '' :Dim myMethod As System.Reflection.MethodInfo = myType.GetMethod(funcName)
    '' '' if myMethod.toString = "" then ...

    myMethod.Invoke(sc, Args)

    If i Mod 22177 = 0 Then
      deltaTime=GetTime()-timerStart2
      statustext("sek: "+deltaTime.toString("00,000")+ "  anz: "+i.toString("0000"))
      'zz.doEvents
    end if
  next    
  deltaTime=GetTime()-timerStart2
  monitorAdd("___________________________ENDE")
  monitorAdd("-------------- DONE")
  monitorAdd("anz schleifen je sek: "+(i / deltaTime * 1000).toString("0000"))
  monitorAdd("")
  statustext("sek: "+deltaTime.toString("00,000")+ "  anz: "+i.toString("0000"))
  'ref.ResumeLayout(False)
  'ref.PerformLayout()


  
End Sub



:function getScriptClassFromFileSpec(filespec as string)
  '' printLine 11, "...", "..." 
  '' printLine 1, "splitFilespecData", fileSpec
  dim path, name, ext as Object
  dim scriptClass as object
  :zz.splitFilespecData(filespec, path, name, ext)
  '' :printLine 2, "splitFilespecData", name
  :statustext(name)
  :scriptClass=zz.scriptClass(name)
  '' :zz.doEvents
  '' :zz.doEvents
  '' :zz.doEvents
  :if scriptClass is Nothing then
     :ERR.number=0
     printLine 11, "ERR", "ERR: scriptKlasse '"+name+"' ...nicht gefunden" 
     trace         "ERR", "ERR: scriptKlasse '"+name+"' ...nicht gefunden" 
     statustext("ERR(scriptClass): "+filespec)
     return nothing
  end if
  return scriptClass
End function




sub sendMesTest2
   monitorAdd("schnellTester-3","...ich bin die nachricht")
exit sub

  'msgBox("OK - 3")
   monitorAdd("schnellTester-3","...ich bin die nachricht")
  dim funcName =ref.element("funcName").text
  ''msgbox(funcName)
  dim para1 = "paaaaaaaaaara"
  ''printLine 1, "MouseClick(e)", funcName
    ''msgbox("myTest: "+para1)

  if funcName<>"" then
    callMySub(funcName, para1)
    '' CallByName(Me, funcName, Microsoft.VisualBasic.CallType.Method, para1)
    '' exit sub
  else
    callMySub("test3",nothing)
  end if
  'msgBox("....OK - 3")
end sub



sub sendMesTest3
   monitorAdd("schnellTester-2","...ich bin die nachricht")
exit sub

  'msgBox("OK - 2")
   monitorAdd("schnellTester-2","...ich bin die nachricht")
     printLine 11, "MouseClick(e)", "test2"
    dim activeTab         = ZZ.getActiveTab()
  dim activeTabType     = ZZ.getActiveTabType()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
 
  trace "ActiveTabFilespec", ActiveTabFilespec
  trace "activeTabType", activeTabType
  trace "", 
  trace "", 
 
  dim filespec as string
  dim path, name, ext as Object
  filespec=ActiveTabFilespec
  zz.splitFilespecData(filespec, path, name, ext)
  trace "path",  path
  'printLine 2, "splitFilespecData", name
  dim scriptClass as object
  scriptClass=zz.scriptClass(name)
  if not scriptClass is nothing then
     : scriptClass.test1()
     : if err.number <>0 then
     :   printLine 11, "ERR", "test1 nicht gefunden"
     :   err.number =0
     : end if
  end if
  ''msgBox ("fileAction 333")
  
  '' msgBox(name)

 ' msgBox("....OK - 2")
end sub

  

'--> ACHTUNG: besser gelöst in 'testLabor'
sub callMySub(funcName, para)
  dim fileSpec as string="C:\yPara\scriptIDE\scriptClass\es_schnellTester.vb"
  dim path, name, ext as Object
  dim scriptClass as object
  :zz.splitFilespecData(filespec, path, name, ext)
  :printLine 2, "splitFilespecData", name
  :scriptClass=zz.scriptClass(name)
  :zz.doEvents
  :zz.doEvents
  :zz.doEvents
  :if scriptClass is Nothing then
     :ERR.number=0
     printLine 11, "ERR", "ERR: scriptKlasse '"+name+"' ...nicht gefunden" 
     trace         "ERR", "ERR: scriptKlasse '"+name+"' ...nicht gefunden" 
     exit sub
  end if
  
  'CallByName(scriptClass, funcName, Microsoft.VisualBasic.CallType.Method, para)
  : CallByName(scriptClass, funcName, Microsoft.VisualBasic.CallType.Method)
  : if ERR.number<>0 then
  :   ERR.number=0
      printLine 11, "ERR", "ERR: Sub 'test2()' ...nicht gefunden"   '@@-
      trace         "ERR", "ERR: Sub 'test2()' ...nicht gefunden"   '@@-
      exit sub
  end if
end sub




sub __test1_MouseClick(e)
  statustext_reset()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  : dim scriptClass = callScriptClassTestProc(ActiveTabFilespec)
  if scriptClass is nothing then exit sub
  : scriptClass.test1()
  : if ERR.number<>0 then
  :   ERR.number=0: dim errMes="ERR: Sub 'test2()' ...nicht gefunden" '@@-
      statustext(errMes)
      'printLine 11, "ERR", errMes
      'trace         "ERR", errMes
    End if
exit sub

  '' 'printLine 3, "aaa","bbbbb"
  '' 'monitorClear()
  '' monitorAdd("=== START")
  '' statustext_reset()
  '' zz.doevents
  '' '' ZZ.IDEHelper.Exec("Window.Reflection", "zz","")
  '' 
  '' ZZ.IDEHelper.Exec("File.Save", "","")
  '' 'monitorAdd("------- SAVED")
  '' 'zz.doevents
  '' 
  '' ZZ.IDEHelper.Exec("Debug.Run", "","")
  '' monitorAdd("----------- COMPILED")
  '' zz.doevents
  '' 
  '' dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  '' : dim scriptClass = callScriptClassTestProc(ActiveTabFilespec)
  '' if scriptClass is nothing then exit sub
  ''     
  '' ''': scriptClass.test1()
  ''  'CallByName(scriptClass, "test1", Microsoft.VisualBasic.CallType.Method)
  '' 
  ''  'CallByName(scriptClass, "test999", Microsoft.VisualBasic.CallType.Method)
  '' 
  ''  
  ''  
  ''  
  ''  dim i as integer
  ''  for i = 0 to 0
  ''    CallByName(scriptClass, "test1", Microsoft.VisualBasic.CallType.Method)
  ''  next
  '' 
  '' 
  '' 
  '' 
  ''  'CallByName(scriptClass, "test1", Microsoft.VisualBasic.CallType.Method)
  '' 
  '' : if ERR.number<>0 then
  '' :   ERR.number=0: dim errMes="ERR: Sub 'test1()' ...nicht gefunden" '@@-
  ''     statustext(errMes)
  ''     'printLine 11, "ERR", errMes
  ''     'trace         "ERR", errMes
  ''   else
  ''      statustext("OK: Test1 ausgeführt")
  ''   End if
  '' monitorAdd("-------------- DONE")
  '' monitorAdd("")
  '' 
'' ''zz.IDEHelper.window.reflection("zz")
''   'ZZ.IIDEIndexList.RestorePos(30,20)
'' 
''   'ZZ.IDEHelper.indexList.RestorePos(30,20)
''    'add(ZZ.getActiveRTF.RTF.Lines.Current.Text)
'' exit sub



  '' statustext_reset()
  '' dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  '' : dim scriptClass = callScriptClassTestProc(ActiveTabFilespec)
  '' if scriptClass is nothing then exit sub
  '' : scriptClass.test1()
  '' : if ERR.number<>0 then
  '' :   ERR.number=0: dim errMes="ERR: Sub 'test1()' ...nicht gefunden" '@@-
  ''     statustext(errMes)
  ''     'printLine 11, "ERR", errMes
  ''     'trace         "ERR", errMes
  ''   End if
End Sub

sub __test2_MouseClick(e)
  statustext_reset()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  : dim scriptClass = callScriptClassTestProc(ActiveTabFilespec)
  if scriptClass is nothing then exit sub
  : scriptClass.test2()
  : if ERR.number<>0 then
  :   ERR.number=0: dim errMes="ERR: Sub 'test2()' ...nicht gefunden" '@@-
      statustext(errMes)
      'printLine 11, "ERR", errMes
      'trace         "ERR", errMes
    End if
End Sub

sub __test3_MouseClick(e)
  statustext_reset()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  : dim scriptClass = callScriptClassTestProc(ActiveTabFilespec)
  if scriptClass is nothing then exit sub
  : scriptClass.test3()
  : if ERR.number<>0 then
  :   ERR.number=0: dim errMes="ERR: Sub 'test3()' ...nicht gefunden" '@@-
      statustext(errMes)
      'printLine 11, "ERR", errMes
      'trace         "ERR", errMes
    End if
End Sub



:function callScriptClassTestProc(filespec as string)
  '' printLine 11, "...", "..." 
  '' printLine 1, "splitFilespecData", fileSpec
  dim path, name, ext as Object
  dim scriptClass as object
  :zz.splitFilespecData(filespec, path, name, ext)
  '' :printLine 2, "splitFilespecData", name
  :statustext(name)
  :scriptClass=zz.scriptClass(name)
  :zz.doEvents
  :zz.doEvents
  :zz.doEvents
  :if scriptClass is Nothing then
     :ERR.number=0
     printLine 11, "ERR", "ERR: scriptKlasse '"+name+"' ...nicht gefunden" 
     trace         "ERR", "ERR: scriptKlasse '"+name+"' ...nicht gefunden" 
     statustext("ERR(scriptClass): "+filespec)
     return nothing
  end if
  return scriptClass
End function


'--> ...



sub onExecute(e, tagDATA)
  dim tagPara as string = trim(tagDATA(1))
  : ZZ.IDEHelper.Exec(tagPara, "",""): Err.Clear
end sub
    

sub onExe(e, tagDATA)
  dim tagPara as string = trim(tagDATA(1))
  : ZZ.shellExecute(tagPara) : Err.Clear
end sub

   


sub onSaveRun (e, tagDATA)
  dim tagPara as string = trim(tagDATA(1))
  monitorClear()
     
  monitorAdd("onSaveRun")
  monitorAdd("------- 1 SAVING")
  
  'ZZ.IDEHelper.Exec("File.Save", "","")
  
  : monitorAdd("------- 2 SAVED"): err.number=0
  zz.doevents
  : ZZ.IDEHelper.Exec("Debug.Run", "","") : err.number=0
  : monitorAdd("------- 3 READY"): err.number=0
  zz.doevents
  ''ZZ.shellExecute(tagPara) : Err.Clear
  
  ref.ResumeLayout(False)
  ref.PerformLayout()
end sub




sub aufZu_MouseClick(e)
  'msgBox("aufZu_MouseClick")
  dim parentWin = ref.parent.parent.parent
  static state as boolean =false
  state=not state
  '' monitorAdd(ref.parent.Proportion.toString)
  '' monitorAdd(ref.parent.parent.Proportion.toString)
  '' monitorAdd(ref.parent.parent.parent.Proportion.toString)
  'exit sub
  dim el as object
  monitorAdd(parentWin.controls.count)
  dim i as integer
  '' for i = 0 to parentWin.controls.count-1
  ''   dim element as string=parentWin.controls(i).toString
  ''   if element.contains("SplitterControl") then
  ''     parentWin.controls(i).top=33
  ''     monitorAdd(parentWin.controls(i).controls(0).toString)
  ''     monitorAdd(parentWin.controls(i).controls(1).toString)
  ''     monitorAdd(parentWin.controls(i).controls(2).toString)
  ''   end if
  ''   monitorAdd(parentWin.controls(i).toString)
  '' next
  
  '' for each el in parentWin.controls.count
  ''   monitorAdd(el.name)
  '' next
  
  
  parentWin.SuspendLayout()
  if state =true then
   parentWin.height=70
  else
   parentWin.height=200
  end if
  parentWin.ResumeLayout(False)
  parentWin.PerformLayout()

end sub




'-->

'--> ICON-LIST



: Function iniIconList() as string
on error resume next
dim curEditorLine as integer=zLN
#Data myIconList as String
''============================================================================================================================================
'  saveRun      |Debug.Run                                                |http://es.teamwiki.net/docs/icon24/distributor-logo.png 
toggleDebugWin  |Addin|##|siaCodeCompiler|##|SHDebug                |http://es.teamwiki.net/docs/icon24/debug.png
toggle          |ToolBar|##|tbScriptWin|##|es_testLabor.mainWin              |http://es.teamwiki.net/docs/icons/gnome-monitor3.png
switch          |ToolBar|##|tbScriptWin|##|es_snippetManager.mainWin         |http://es.teamwiki.net/docs/icon24/insert-object.png
switch          |ToolBar|##|tbScriptWin|##|es_internSuche.mainWin            |http://es.teamwiki.net/docs/icon24/SearchFernglas.png
switch          |Addin|##|siaIDEMain|##|OpenedFiles                         |http://es.teamwiki.net/docs/icon24/FavouritesGelb.png
onWebBrowser1   |                       |http://es.teamwiki.net/docs/agt_web.png
'toggle         |ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sg_memo.dll|##|root_sg_memo.sg_memo        |http://es.teamwiki.net/docs/icon24/ModifyEdit.png
toggleTab       |ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sg_memo.dll|##|root_sg_memo.sg_memo        |http://es.teamwiki.net/docs/icon24/ModifyEdit.png
toggle          |ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\mwTimer.dll|##|mwTimer.sw_analogClock      |http://es.teamwiki.net/docs/icon24/Clock.png
togglePanel     |p: 8           |http://es.teamwiki.net/docs/icon24/white01.png
toggle          |p: para        |http://es.teamwiki.net/docs/icon24/stock_test-mode.png
toggle          |p: para        |http://es.teamwiki.net/docs/icon24/ViewSearch.png
toggle          |p: para        |http://es.teamwiki.net/docs/icon24/white01.png
toggle          |p: para        |http://es.teamwiki.net/docs/icon24/gtk-edit.png 
 '' 
''toggle       |Addin|##|siaCodeCompiler|##|SHDebug                                      |http://es.teamwiki.net/docs/icon24/debug.png
''toggle       |Addin|##|siaCodeCompiler|##|SHDebug                                        |http://es.teamwiki.net/docs/icon24/debug.png
 '' 
 '' 010 | http://es.teamwiki.net/docs/icon24/checkered_flag.png 
 '' 011 | http://es.teamwiki.net/docs/icons/navigation-090-white.png 
 '' 012 | http://es.teamwiki.net/docs/icons/navigation-180-white.png
 '' 013 | http://es.teamwiki.net/docs/icons/navigation-270-white.png 
 '' 014 | xxx 
 '' 015 | http://es.teamwiki.net/docs/icons/navigation-090-button.png 
 '' 016 | http://es.teamwiki.net/docs/icons/navigation-180-button.png 
 '' 017 | http://es.teamwiki.net/docs/icons/navigation-270-button.png 
 '' 017 | http://es.teamwiki.net/docs/icon24/Good-mark.png
 '' 018 | xxx 
 '' 019 | xxx 
 
''============================================================================================================================================
#EndData
  '' Trick 17: Zeilennummer dazupacken
  dim RESULT as string=curEditorLine & vbNewLine & myIconList
  redim iconList(111)
  dim DATA() as string=split(RESULT,vbNewLine)
  dim LINE() as string
  dim temp as string
  dim index as integer
  dim counter as integer =0
  'dim icon,type, para as string
  dim max as integer=DATA.length
  dim i as integer
  for i =1 to 20 '   max-1
    temp=trim(DATA(i))
    if temp="" then continue for
    if temp.startsWith("'") then continue for
    temp = replace(temp, "|##|","_##_")+"||"             '...scriptedPanelId maskieren
    temp = replace(temp, "|","<³³>")                     '...scriptedPanelId maskieren
    temp = replace(temp, "_##_", "|##|")           '...scriptedPanelId maskieren
    LINE = split(DATA(i), "<³³>")
    index = val(LINE(0))
    iconList(counter)=temp
    counter=counter+1
  next
  return RESULT
end function



: sub insertIconsFromIconList(PnlRef)
  on error resume next
  dim index, left, top, deltaTop, deltaLeft as integer
  dim text, icon, item  as string
  dim DATA() as string
  dim el as object
  ''dim bgColor="#696E73"
  dim bgColor="#4A7CDB"  'E0DDD7
  deltaTop=0
  deltaLeft=0
  for index =0 to 11
    text="" ' index.toString
    item=iconList(index)
    if trim(item)="" then continue for

    DATA=split(item,"<³³>")
    'if icon="xxx" then icon="http://es.teamwiki.net/docs/icons/edit-number.png"
    icon=trim(DATA(2))
    if icon="xxx" then 
      icon=""
      text="" '"<"+index.toString+">"
    end if
    left=(index mod 10)*30
    top=(index - (index mod 10)) *4
    dim id as string=trim(DATA(0))+"_"+index.toString
    ''monitorAdd(top.toString, left.tostring)
   
     el=PnlRef.addButton(id ,text,bgColor,Left+deltaLeft,Top+deltaTop,32 ,32 ,icon) ' ,flags,handler)
     el.ImageAlign = System.Drawing.ContentAlignment.MiddleCenter
    'el.TextAlign = System.Drawing.ContentAlignment.BottomCenter
    'el.TextImageRelation = System.Windows.Forms.TextImageRelation.Overlay 'ImageAboveText
    el.foreColor=ColorTranslator.FromHtml("#fff")
    'el.UseVisualStyleBackColor = True
    el.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    el.FlatAppearance.BorderSize = 1
    el.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml("#FFD350")  '#4A7CDB
    el.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#5A8CeB")
    el.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0)
    el.margin = New System.Windows.Forms.Padding(0, 0, 0, 0)
    el.tag=item
    el.AccessibleRole = System.Windows.Forms.AccessibleRole.CheckButton
    'el.checked=true
    'el.Checked = True
    'el.CheckState = System.Windows.Forms.CheckState.Checked

  next
end sub


:sub insertIconsFromIndex(ref as object,index as integer, left as integer, top as integer)
  on error resume next
  dim bgColor as string = "#fff"
  dim txt as string=val(index)
  left=(index mod 10)*30
  top=(index - (index mod 10)) *40
  
  
  'ref.addButton(id,txt,bgColor,Left,Top,30 ,40 ,iconList(index)) ' ,flags,handler)

  dim el as object
  with ref
    .insertX = 5:.insertY = 5:
    el=.addButton  ("imgList01"      , "33"    , "#ccc")
    el.tag="...tag": el.height=40:el.width=30
    el.image=ZZ.GetImageCached(iconList(40))
    el.ImageAlign = System.Drawing.ContentAlignment.TopCenter
    el.TextAlign = System.Drawing.ContentAlignment.BottomCenter
    
    .insertX = 40:.insertY = 150:
    el=.addButton  ("imgList02"      , "45"    , "#ccc")
    el.tag="...tag": el.height=40:el.width=30
    el.image=ZZ.GetImageCached(iconList(41))
    el.ImageAlign = System.Drawing.ContentAlignment.TopCenter
    el.TextAlign = System.Drawing.ContentAlignment.BottomCenter

   .insertX = 10:.insertY = 190:
    el=.addButton  ("imgList03"      , "45"    , "#ccc")
    el.tag="...tag": el.height=40:el.width=30
    el.image=ZZ.GetImageCached(iconList(41))
    el.ImageAlign = System.Drawing.ContentAlignment.TopCenter
    el.TextAlign = System.Drawing.ContentAlignment.BottomCenter
  end with

end sub 





'-->
'--> outMONITOR


public globLastEventHandlerEventName as string
public globLastEventHandlerScriptClass as string
public globLastEventHandlerFound as Boolean

sub onSetLastEventHandlerPara(evName as string, scriptClass as string, found as boolean)
  globLastEventHandlerEventName = evName
  globLastEventHandlerScriptClass = scriptClass
  globLastEventHandlerFound = found
  dim el =ref.element("lab_cmdInsertEventHandler")
  el.text=evName
  if found=false then
    el.foreColor=ColorTranslator.FromHtml("#f66"): 
  else
    el.foreColor=ColorTranslator.FromHtml("#999"): 
  end if
end sub


public globMonitorRef as object

sub setLastEventHandlerPara(evName as string, scriptClass as string, found as boolean)
  if evName="cmdInsertEventHandler" then exit sub
  'dim scriptClass = zz.scriptClass("es_schnellTester")
  onSetLastEventHandlerPara(evName, scriptClass, found)
end sub



 sub clearAll()
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




