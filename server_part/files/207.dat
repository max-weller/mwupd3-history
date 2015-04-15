dim appIcon as string ="http://es.teamwiki.net/docs/img/icons/switch_off.png"



#Para DebugMode internal
#Para SilentMode true


#Imports System.Diagnostics

'-->
'--> G L O B A L - C O N F I G


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


dim iproc as sys_interproc
public iconList() as string


public WithEvents FormRef As Form
public WithEvents ref As ScriptedPanel


public shared toolBar As Object
public shared statBar As Object
public shared container1 As Object
public toolBarColor as string

Public Const winId as string ="es_HauptUmschalter.window"
public winCaption as string = "Hauptschalter"
dim WithEvents myTextArea as textbox
dim WithEvents myPicture as pictureBox      
public withEvents myWin2 as object
dim myImg as object
shared skipAllEvents as boolean =false

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
Public Declare Function ReleaseCapture Lib "user32" () As Integer
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer


Public Enum WindowShowStyle As UInteger
    ''' <summary>Hides the window and activates another window.</summary>
    ''' <remarks>See SW_HIDE</remarks>
    Hide = 0
    '''<summary>Activates and displays a window. If the window is minimized
    ''' or maximized, the system restores it to its original size and
    ''' position. An application should specify this flag when displaying
    ''' the window for the first time.</summary>
    ''' <remarks>See SW_SHOWNORMAL</remarks>
    ShowNormal = 1
    ''' <summary>Activates the window and displays it as a minimized window.</summary>
    ''' <remarks>See SW_SHOWMINIMIZED</remarks>
    ShowMinimized = 2
    ''' <summary>Activates the window and displays it as a maximized window.</summary>
    ''' <remarks>See SW_SHOWMAXIMIZED</remarks>
    ShowMaximized = 3
    ''' <summary>Maximizes the specified window.</summary>
    ''' <remarks>See SW_MAXIMIZE</remarks>
    Maximize = 3
    ''' <summary>Displays a window in its most recent size and position.
    ''' This value is similar to "ShowNormal", except the window is not
    ''' actived.</summary>
    ''' <remarks>See SW_SHOWNOACTIVATE</remarks>
    ShowNormalNoActivate = 4
    ''' <summary>Activates the window and displays it in its current size
    ''' and position.</summary>
    ''' <remarks>See SW_SHOW</remarks>
    Show = 5
    ''' <summary>Minimizes the specified window and activates the next
    ''' top-level window in the Z order.</summary>
    ''' <remarks>See SW_MINIMIZE</remarks>
    Minimize = 6
    '''   <summary>Displays the window as a minimized window. This value is
    '''   similar to "ShowMinimized", except the window is not activated.</summary>
    ''' <remarks>See SW_SHOWMINNOACTIVE</remarks>
    ShowMinNoActivate = 7
    ''' <summary>Displays the window in its current size and position. This
    ''' value is similar to "Show", except the window is not activated.</summary>
    ''' <remarks>See SW_SHOWNA</remarks>
    ShowNoActivate = 8
    ''' <summary>Activates and displays the window. If the window is
    ''' minimized or maximized, the system restores it to its original size
    ''' and position. An application should specify this flag when restoring
    ''' a minimized window.</summary>
    ''' <remarks>See SW_RESTORE</remarks>
    Restore = 9
    ''' <summary>Sets the show state based on the SW_ value specified in the
    ''' STARTUPINFO structure passed to the CreateProcess function by the
    ''' program that started the application.</summary>
    ''' <remarks>See SW_SHOWDEFAULT</remarks>
    ShowDefault = 10
    ''' <summary>Windows 2000/XP: Minimizes a window, even if the thread
    ''' that owns the window is hung. This flag should only be used when
    ''' minimizing windows from a different thread.</summary>
    ''' <remarks>See SW_FORCEMINIMIZE</remarks>
    ForceMinimized = 11

    End Enum


Dim m_FormBorder As New clsFormBorder()

'--> 
'--> T E S T 
Sub test1()
  setMonitorRef()
  'msgBox("OK - 1")
  monitorClear()
  enumAllTopLevelWin()
End Sub


Sub test2()
  setMonitorRef()
  dim hwnd as int32
  hwnd=findWindowInstring("W-E-B")
  ShowWindow (hwnd, WindowShowStyle.ShowMinimized)
End Sub

Sub test3()
  setMonitorRef()
  'msgBox("OK - 3")
  dim hwnd as int32
  hwnd=findWindowInstring("W-E-B")
  'ShowWindow (hwnd, WindowShowStyle.ShowMinimized)
  'ShowWindow (hwnd, WindowShowStyle.ShowMaximized)
   'ShowWindow (hwnd, WindowShowStyle.ShowMinimized)
   ShowWindow (hwnd, WindowShowStyle.ShowNormal)

End Sub


Function showWindowMaximize(titlePartial as string)

End Function


Function showWindowMinimized(titlePartial as string)

End Function

Function enumAllTopLevelWin()
  dim enumerator = new  WindowsEnumerator
  dim topLevelWin As List(Of ApiWindow) = enumerator.GetTopLevelWindows()
  dim item as apiWindow
   for each item in topLevelWin
      monitorAdd(item.hwnd.toString+"     " +vbTab+item.className, vbTab+"          " +item.MainWindowTitle )
   next 
  monitorAdd("Anzahl der topLevelWindows: "+topLevelWin.count.toString)
End Function


:Function findWindowInstring(titlePartial as string)as int32
  dim such as string=titlePartial.toLower
  dim enumerator = new  WindowsEnumerator
  dim topLevelWin As List(Of ApiWindow) = enumerator.GetTopLevelWindows()
  monitorAdd("topLevelWindows: "+topLevelWin.count.toString)
  dim item as apiWindow
  for each item in topLevelWin
    if item.MainWindowTitle.toLower.contains(such) then
      monitorAdd(item.hwnd.toString+"     " +vbTab+item.className, vbTab+"          " +item.MainWindowTitle )
      return item.hwnd
      'exit function
    end if
  next 
End Function

'-->
'--> I N I 

: Sub onTerminate(optional intern as string="")
   : dim extern as string =""
   : if intern="" then
   :    extern ="=================EXTERN===: "+name
   :    intern=""
   :  else
   :  intern=intern+": "+name
   : end if
   : trace "onTerminate: "+intern,"ON-TERMINATE "+extern
   : ' formRef.dispose()
   : 'exit sub
   : 'dim dockPanel as WeifenLuo.WinFormsUI.Docking.DockPanel
   : 'dockPanel = cType(zz.ideHelper.mainwindow.DockPanel1, WeifenLuo.WinFormsUI.Docking.DockPanel)
   : 'removeHandler dockPanel.ActivePaneChanged, AddressOf  DockPanel1_ActivePaneChanged
   : 'ZZ.IDEHelper.RemoveToolbar("toolbar.tb1")
   : err.number=0
End Sub


Sub GetFormRef_xxx()
  'If ref IsNot Nothing Then Exit Sub
  ref = ZZ.IDEHelper.CreateDockWindow(winID, 4): err.number=0
  formRef = ref.form
  formRef.text=winCaption
  : CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = false
End Sub

'' Sub GetFormRef()
''   ''DAS MACHT EINE NORMALE FORM
''   formRef = New frmTB_scriptWin 
''   'ref = ZZ.IDEHelper.CreateForm(winID)
''   ref = CType(FormRef, Object).PNL
''   'ref = CType(FormRef, ScriptedPanel).PNL
''   FormRef.Text = winCaption
'' End Sub
''


Sub GetFormRef()
  If ref IsNot Nothing Then Exit Sub
  'trace winID
  'ref = ZZ.CreateWindow("es_hauptumschalter.mainWin", DWndFlags.StdWindow Or DWndFlags.NoAutoShow)
  ref = ZZ.CreateWindow("es_hauptumschalter.mainWin", DWndFlags.StdWindow)
  FormRef = ref.Form
End Sub


'-->
'-->A U T O S T A R T 
 

Sub sys_interproc_Message(source As String, cmdString As String, para As String) ''... Handles sys_interproc.Message
  If cmdString = "Goto" Then
   trace "sys_interproc_Message", para
  End If
End Sub


'' ...on initialize
'' Sub OnMyForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
''   iproc=New sys_interproc("HAUPT_UMSCHALTER")
''   iproc.Commands.Add("CMD  ", "Goto", "x|y", "")
''   AddHandler iproc.Message, AddressOf sys_interproc_Message
'' End Sub
'' 

sub autoStart()
  iproc=New sys_interproc("HAUPT_UMSCHALTER")
  iproc.Commands.Add("CMD  ", "Goto", "x|y", "")
  AddHandler iproc.Message, AddressOf sys_interproc_Message
  ''iproc.SendCommand("appbar", "TEST", "1-2-3-4-5-6-7-8-9") 'para = beliebiger String, wird als MsgBox bzw. Trace ausgegeben 

  zz.traceClear()
  'formRef.close
  ''onTerminate("intern!!!")

  GetFormRef()
  dim el as object
  toolBarColor="#BFBFBF"



With ref
    .resetControls ()
    dim element as control
    .activateEvents = "|IconMouseDown|   |TextboxKeyDown|"
 
 '--> ...topbar ========================================================
     .insertX = 10 :.insertY = 30 :  
     'el=.addButton ("cmdTerminate"  , "x"          , "#f00"): 
     el=.addIcon("cmdTerminate_02", "http://es.teamwiki.net/docs/img/icons/Close_Box_Red.png" )

    .insertX = -1 :.insertY = -23 ' 28 '-5
    .insertX = -115 :.insertY = -10 ' 28 '-5
     'myPicture=.addIcon("appPicture", "http://es.teamwiki.net/docs/icons/xconsole.png" )
     myPicture=.addIcon("appPicture", "http://portal.0815team.de/pics/grip.jpg")

    toolBar     =.addPanel("toolBar"      ,   , intHeight:=33)
    toolbar.top=21
    toolbar.left=40
    toolbar.width=3333
    toolbar.visible=true
    toolBar.resetControls  (10,5,1)
    toolbar.height=38
    ''toolbar.BackColor = ColorTranslator.FromHtml(toolBarColor)
    toolbar.BackColor = ColorTranslator.FromHtml("#5A8CEB")
    
   
    
    
    '--> ...textBox

    .BR(25) '--------------------------------------------------
   .insertX = 50 ::.insertY = 3
     el=.addLabel  ("lab02", "Url" ,  ,"#ddd",,,25,16) : 
     el.backColor=ColorTranslator.FromHtml("#363E57")
     ': makeJumboLabel(el)
   .addTextbox ("urlBar" ,  200  , "xyz" , 0,"#bbb" , , , )
    el = ref.element("urlBar")
    el.text="xyz"
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.foreColor=ColorTranslator.FromHtml("#000") :el.borderStyle=0
    el.backColor=ColorTranslator.FromHtml("#FFF48A")
    el.WordWrap=false
 
        
    .insertX = 300 :.insertY = 3
     el=.addLabel  ("lab01", "Suchen" ,  ,"#ddd",,,44,16) : 
     el.backColor=ColorTranslator.FromHtml("#363E57")
     ': makeJumboLabel(el)
    .addTextbox("searchBar" ,  200         , " nicName"    , 0  ,"#bbb" , , , )
    el = ref.element("searchBar")
    el.text="abc"
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.foreColor=ColorTranslator.FromHtml("#000") :el.borderStyle=0
    el.backColor=ColorTranslator.FromHtml("#ddd")
    el.WordWrap=false

  '--> knöpfe
    '.insertX = 150: .insertY = 0
     'el=.addButton ("test9" ,               "titleBar off"    , "#FB7929"): el.width=66
     '.insertX = 215: el=.addButton ("test1"     , "SelectAll"       , "#B1B13B"): el.width=66

     : .insertY = 3
                     'el=.addButton ("outMonitor_toClipboard"   , "--> Clipboard"   , "#E0E0E0"): el.width=77
     '.insertX = 120: el=.addButton ("cmdTerminate"  , "x"          , "#f00"): el.width=30
     .insertX = 605: el=.addButton ("test1"         , "test 1"          , "#E0E0E0"): el.width=50
     .insertX = 655: el=.addButton ("test2"         , "test 2"          , "#E0E0E0"): el.width=50
     .insertX = 705: el=.addButton ("test3"         , "test 3"          , "#E0E0E0"): el.width=50


      .insertX = 755: el=.addButton ("cmdNavUrl_02"    , "google"     , "#E0E0E0"): el.width=50
                      el.tag="google.de"
      .insertX = 805: el=.addButton ("cmdNavUrl_01"    , "pinboard"     , "#E0E0E0"): el.width=50
                      el.tag="http://devnet.teamwiki.net/xymemo.html?doc=es-pb-mw-script-ide02"
      .insertX = 855: el=.addButton ("cmdNavUrl_01"    , "sideBar"     , "#E0E0E0"): el.width=50
                      el.tag="http://es.teamwiki.net/sb2-seite2.html"

    .activateEvents = "|TextboxKeyDown|  |TextboxKeyDown|  |TextboxTextChanged|"

 '' ??? addHandler textarea.KeyDown, AddressOf  textarea_KeyDown
    element=ref.element("test1"): addHandler element.Click, AddressOf  test1_Click
    element=ref.element("test2"): addHandler element.Click, AddressOf  test2_Click
    element=ref.element("test3"): addHandler element.Click, AddressOf  test3_Click

'--> ...textarea
    .insertX = 200 :.insertY = 200
    .addTextbox ("outMonitor" ,  -2         , "inBox"    , 0  , "#FFFF99", , , "multiline=260")
       myTextArea=ref.element("txt_outMonitor")
       myTextArea.WordWrap=false
       myTextArea.scrollbars=System.Windows.Forms.ScrollBars.Vertical
       myTextArea.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

  End With
  
  iniIconList()
  insertIconsFromIconList(toolBar)
  
  formRef.text="Hauptumschalter"
  formRef.show
  formRef.left=222
  formRef.left=1888 '555
  formRef.top=-6
  formRef.width=888
  formRef.height=66
  formRef.topmost=true
  formRef.backColor=ColorTranslator.FromHtml("#363E57")

  'exit sub
  
  m_FormBorder.client = formRef
  m_FormBorder.Titlebar = False
  
  'm_FormBorder.Moveable=true
  m_FormBorder.AutoDrag=true
  m_FormBorder.showInTaskBar=false
  
  'formRef.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
  formRef.topmost=true
  
End Sub


'-->
'--> E V E N T S 


Sub onButtonEvent(e)
  'msgbox("onButtonEvent")
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

sub onTextboxEvent(e)
  setMonitorRef()
  MonitorAdd("======= onTextEvents: "+e.Sender.Name)
  'MonitorAdd("======= onTextEvents: "+e.eventName)
  'monitorAdd("   ClassName: " +e.ClassName+" <<<")
  'monitorAdd(" ControlType: " +e.ControlType+" <<<")
  'monitorAdd(" keyString: " +e.keyString+" <<<")
  'if e.eventName="KeyDown" then
    'monitorAdd(" =====>>>>: " +e.keyString+" <<<")
  'end if
  if e.keyString="RETURN" then
    if e.sender.name="txt_urlBar" then
      dim url as string=e.sender.text
      navigateUrl(url)
    end if
    if e.sender.name="txt_searchBar" then
      dim url as string=e.sender.text
      dim prefix as string="http://www.google.de/search?num=100&hl=de&q="
      navigateUrl(prefix+url)
    end if
  end if 
  
end sub

'ON txt_urlBar_KeyDown

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
'-->E V E N T S - 2


sub onExe(e)
  dim tag as string = e.sender.tag
  dim tagDATA()= Split(tag, "<³³>")
  dim tagPara as string = trim(tagDATA(1))
  : ZZ.shellExecute(tagPara) : Err.Clear
end sub



sub onExecute(e)
  dim tag as string = e.sender.tag
  dim tagDATA()= Split(tag, "<³³>")
  dim tagPara as string = trim(tagDATA(1))
  : ZZ.IDEHelper.Exec(tagPara, "",""): Err.Clear
end sub


sub cmdCallByName (e, tagDATA)
  dim tagPara as string = trim(tagDATA(1))
  statustext_reset()
  monitorAdd("cmdCallByName: "+tagPara)
  
  'dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  'dim scriptClass = getScriptClassFromFileSpec(ActiveTabFilespec)
  dim scriptClass = me
  if scriptClass is nothing then exit sub
  dim funcName as string=tagPara
  dim i as integer
  for i = 0 to 0
    CallByName(scriptClass, funcName, Microsoft.VisualBasic.CallType.Method)
    statustext(i.tostring)
  : if ERR.number<>0 then
  :   ERR.number=0: dim errMes="ERR: Sub 'test1()' ...nicht gefunden" 
      statustext(errMes)
      'printLine 11, "ERR", errMes
      'trace         "ERR", errMes
      exit sub
    else
      statustext("OK: '"+funcName + "' ... ausgeführt")
    End if
  next    
end sub


Sub onWebBrowser1(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! 'onWebBrowser1': ...in Arbeit")
  'msgBox("'onWebBrowser1': " + "...Coming soon ")
  '...
  dim exe as string="C:\yEXE\scriptEXE\es_webbrowser3\es_webbrowser3.exe"
  dim Para as string
  para ="http://www.iconfinder.com/icondetails/17829/24/"
  para ="http://secure.teamwiki.net/mydocs2?view=mini#"
  : ZZ.shellExecute(exe, para) : Err.Clear
  
End Sub



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


Sub toggleDebugWin(e) 
  with zz.scriptClass("es_testLabor")
    .toggleTraceWindow()
  end with
End Sub



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


Sub cmdTerminate(e)  ' ...
  formRef.close
End Sub


Sub cmdNavUrl(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))

  dim url=e.sender.tag.toString
  navigateUrl(url)
End Sub

sub navigateUrl(url as string)
 'iproc.SendCommand("appbar", "TEST", "1-2-3-4-5-6-7-8-9") 'para = beliebiger String, wird als MsgBox bzw. Trace ausgegeben 
  iproc.SendCommand("web_browser", "NAVIGATEURL", url) 'para = url|||PARA??? 
end sub



sub btn_test9_MouseClick(e)
  msgBox("btn_test9_MouseClick")
end sub


Sub myPicture_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles myPicture.MouseDown
  if e.Button=Windows.Forms.MouseButtons.Right then
    'hide()
    exit Sub
  end if
  dim hwnd as integer
  Call ReleaseCapture()
  Call SendMessage(FormRef.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub



  'Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRunCommand.Click

Sub test1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  'msgbox("ok ???")
  test1()
exit sub


  For Each p As Process In Process.GetProcesses()
       trace p.MainWindowTitle 
   Next



  '' 'trace "PictureBox1_MouseDown","..........."
  '' 'formRef.parent.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
  '' if e.Button=Windows.Forms.MouseButtons.Right then
  ''   hide()
  ''   exit Sub
  '' end if
  '' 
  '' dim hwnd as integer
  '' Call ReleaseCapture()
  '' ''Call SendMessage(myPicture.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  '' 
  '' ''Call SendMessage(FormRef.parent.parent.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  '' Call SendMessage(FormRef.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Sub test2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  'msgbox("ok ???")
  test2()
End Sub

Sub test3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  'msgbox("ok ???")
  test3()
End Sub

sub test1_MouseClick(e)
  '' msgbox("ok ???")
  test1()
end sub

sub test2_MouseClick(e)
  ''msgbox("ok ???")
  test2()
end sub

sub test3_MouseClick(e)
  ''msgbox("ok ???")
  test3()
end sub




'-->

'--> ICON-LIST



: Function iniIconList() as string
on error resume next
dim curEditorLine as integer=zLN
#Data myIconList as String
''============================================================================================================================================
'  saveRun      |Debug.Run                                                |http://es.teamwiki.net/docs/icon24/distributor-logo.png 
toggle          |ToolBar|##|tbScriptWin|##|es_testLabor.mainWin              |http://es.teamwiki.net/docs/icons/gnome-monitor3.png
toggleDebugWin  |Addin|##|siaCodeCompiler|##|SHDebug                         |http://es.teamwiki.net/docs/icon24/debug.png
toggle       |Addin|##|siacodecompiler|##|SHPrintline            |http://es.teamwiki.net/docs/icon24/gnome-dev-printer.png
'cmdCallByName   |clearAll                                                    |http://es.teamwiki.net/docs/icon24/emblem-noread.png
toggleTab       |ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sg_memo.dll|##|root_sg_memo.sg_memo        |http://es.teamwiki.net/docs/icon24/ModifyEdit.png
togglePanel     |p: 8                                                        |http://es.teamwiki.net/docs/icon24/back_orange.png

switch          |ToolBar|##|tbScriptWin|##|es_snippetManager.mainWin         |http://es.teamwiki.net/docs/icon24/insert-object.png
switch          |ToolBar|##|tbScriptWin|##|es_internSuche.mainWin            |http://es.teamwiki.net/docs/icon24/SearchFernglas.png
'switch          |Addin|##|siaIDEMain|##|OpenedFiles                         |http://es.teamwiki.net/docs/icon24/FavouritesGelb.png
onWebBrowser1   |                       |http://es.teamwiki.net/docs/agt_web.png
'toggle         |ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sg_memo.dll|##|root_sg_memo.sg_memo        |http://es.teamwiki.net/docs/icon24/ModifyEdit.png
toggle          |ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\mwTimer.dll|##|mwTimer.sw_analogClock      |http://es.teamwiki.net/docs/icon24/Clock.png
'' togglePanel     |p: 8           |http://es.teamwiki.net/docs/icon24/white01.png
'' toggle          |p: para        |http://es.teamwiki.net/docs/icon24/stock_test-mode.png
'' toggle          |p: para        |http://es.teamwiki.net/docs/icon24/ViewSearch.png
'' toggle          |p: para        |http://es.teamwiki.net/docs/icon24/white01.png
'' toggle          |p: para        |http://es.teamwiki.net/docs/icon24/gtk-edit.png 
''  
toggle         |ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sw_colorPicker.dll|##|sw_colorPicker.sg_colorPicker        |http://icons3.iconfinder.netdna-cdn.com/data/icons/nuvola2/32x32/apps/kwikdisk.png
onExe          |C:\yEXE\traceMonitor.exe                         |http://es.teamwiki.net/docs/icon24/agentsvr_113-32.png
onExe          |C:\yEXE\ide3\IprocViewer.exe                     |http://es.teamwiki.net/docs/icon24/checkered_flag.png
onExe          |C:\yEXE\extractAssociatedIcon.exe                |http://es.teamwiki.net/docs/icon24/stock_draw-selection.png 
onExe          |C:\yEXE\mwSidebar.exe                            |http://es.teamwiki.net/docs/icon24/SysTrayDemo2_1-32.png 

toggle          |p: para        |http://icons3.iconfinder.netdna-cdn.com/data/icons/glaze/32x32/apps/konsole.png
toggle          |p: para        |http://icons3.iconfinder.netdna-cdn.com/data/icons/gnomeicontheme/32x32/actions/redhat-home.png
toggle          |p: para        |http://icons3.iconfinder.netdna-cdn.com/data/icons/glaze/32x32/filesystems/folder_home.png
toggle          |p: para        |http://es.teamwiki.net/docs/icon24/stock_test-mode.png
toggle          |p: para        |http://es.teamwiki.net/docs/icon24/stock_test-mode.png
toggle          |p: para        |http://es.teamwiki.net/docs/icon24/stock_test-mode.png
toggle          |p: para        |http://es.teamwiki.net/docs/icon24/stock_test-mode.png





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
  'for i =1 to 30 '   max-1
  for i =1 to max-1
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
  for index =0 to 19
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
    'left=(index mod 10)*30
    left=index *30
    left=index *35
    'top=(index - (index mod 10)) *4
    top=0
    dim id as string=trim(DATA(0))+"_"+index.toString
    ''monitorAdd(top.toString, left.tostring)
   
     'el=PnlRef.addButton(id ,text,bgColor,Left+deltaLeft,Top+deltaTop,32 ,32 ,icon) ' ,flags,handler)
     el=PnlRef.addButton(id ,text,bgColor,Left+deltaLeft,Top+deltaTop,36 ,38 ,icon) ' ,flags,handler)
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
'--> c l s F O R M B O R D E R

'Namespace MVPS

  ' *********************************************************************
  '  Copyright ©1995-2001 Karl E. Peterson, All Rights Reserved
  '  http://www.mvps.org/vb
  ' *********************************************************************
  '  You are free to use this code within your own applications, but you
  '  are expressly forbidden from selling or otherwise distributing this
  '  source code without prior written consent.
  ' *********************************************************************
  '  GENERAL USAGE NOTE:  Be sure to set the Client form with these
  '  properties, in order to insure the toggles in this class work:
  '   * BorderStyle:  2 - Sizable
  '   * ControlBox:   True
  '  You may freely change these and all other properties at runtime.
  ' *********************************************************************


': Namespace MVPS
  Class clsFormBorder


    ' Point type used to track cursor.
    Private Structure POINTAPI
      Dim x As Integer
      Dim Y As Integer
    End Structure

    ' Win32 APIs used to toggle border styles.
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
    Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Integer) As Integer

    ' Win32 APIs used to automate drag and sysmenu support.
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal revert As Integer) As Integer
    Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
    Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Integer) As Integer
    Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Integer, ByVal un As Integer, ByVal b As Integer, ByVal lpMenuItemInfo As MENUITEMINFO) As Integer
    Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Integer, ByVal un As Integer, ByVal bool As Integer, ByVal lpcMenuItemInfo As MENUITEMINFO) As Integer
    Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Declare Function ReleaseCapture Lib "user32" () As Integer
    Private Declare Function GetCursorPos Lib "user32" (ByVal lpPoint As POINTAPI) As Integer

    ' Used to get menu information.
    Private Structure MENUITEMINFO
      Dim cbSize As Integer
      Dim fMask As Integer
      Dim fType As Integer
      Dim fState As Integer
      Dim wID As Integer
      Dim hSubMenu As Integer
      Dim hbmpChecked As Integer
      Dim hbmpUnchecked As Integer
      Dim dwItemData As Integer
      Dim dwTypeData As String
      Dim cch As Integer
    End Structure


    ' Used to support captionless drag
    Private Const WM_NCLBUTTONDOWN = &HA1
    Private Const HTCAPTION = 2

    ' Undocumented message constant.
    Private Const WM_GETSYSMENU = &H313

    ' Used to select which menu to remove.
    Private Const MF_BYCOMMAND = &H0&
    Private Const MF_BYPOSITION = &H400

    ' Toggles enabled state of menu item.
    Private Const MF_ENABLED = &H0&
    Private Const MF_GRAYED = &H1&
    Private Const MF_DISABLED = &H2&

    ' Menu information constants.
    Private Const MIIM_STATE As Integer = &H1
    Private Const MIIM_ID As Integer = &H2
    Private Const MIIM_SUBMENU As Integer = &H4
    Private Const MIIM_CHECKMARKS As Integer = &H8
    Private Const MIIM_TYPE As Integer = &H10
    Private Const MIIM_DATA As Integer = &H20

    ' Used to get window style bits.
    Private Const GWL_STYLE = (-16)
    Private Const GWL_EXSTYLE = (-20)

    ' Style bits.
    Private Const WS_MAXIMIZEBOX = &H10000
    Private Const WS_MINIMIZEBOX = &H20000
    Private Const WS_THICKFRAME = &H40000
    Private Const WS_SYSMENU = &H80000
    Private Const WS_CAPTION = &HC00000

    ' Extended Style bits.
    Private Const WS_EX_TOPMOST = &H8
    Private Const WS_EX_TOOLWINDOW = &H80
    Private Const WS_EX_CONTEXTHELP = &H400
    Private Const WS_EX_APPWINDOW = &H40000

    ' Force total redraw that shows new styles.
    Private Const SWP_FRAMECHANGED = &H20
    Private Const SWP_NOMOVE = &H2
    Private Const SWP_NOZORDER = &H4
    Private Const SWP_NOSIZE = &H1

    ' Used to toggle into topmost layer.
    Private Const HWND_TOPMOST = -1
    Private Const HWND_NOTOPMOST = -2

    ' System menu command values commonly used by VB.
    Private Const SC_SIZE = &HF000&
    Private Const SC_MOVE = &HF010&
    Private Const SC_MINIMIZE = &HF020&
    Private Const SC_MAXIMIZE = &HF030&
    Private Const SC_CLOSE = &HF060&
    Private Const SC_RESTORE = &HF120&

    ' Enumerated sysmenu items.
    Public Enum SysMenuItems
      smRestore = SC_RESTORE
      smMove = SC_MOVE
      smSize = SC_SIZE
      smMinimize = SC_MINIMIZE
      smMaximize = SC_MAXIMIZE
      smClose = SC_CLOSE
    End Enum

    ' References to client form.
    Private WithEvents m_Client As Form
    'Attribute m_Client.VB_VarHelpID = -1
    'Private WithEvents m_MdiClient As MDIForm
    'Attribute m_MdiClient.VB_VarHelpID = -1
    Private m_hWnd As Integer

    ' Member variables
    Private m_AutoSysMenu As Boolean
    Private m_AutoDrag As Boolean


    ' ************************************************
    '  Sunken Client Events
    ' ************************************************
    Private Sub m_Client_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
      Call ClientMouseDown(Button, Shift, x, Y)
    End Sub

    Private Sub m_Client_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
      Call ClientMouseUp(Button, Shift, x, Y)
    End Sub

    Private Sub m_MdiClient_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
      Call ClientMouseDown(Button, Shift, x, Y)
    End Sub

    Private Sub m_MdiClient_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
      Call ClientMouseUp(Button, Shift, x, Y)
    End Sub

    ''Public Shared Sub moveMeHwnd(ByVal hwnd As Integer, Optional ByVal clickPos As WindowsUtilities.HitTestValues = WindowsUtilities.HitTestValues.HTCAPTION)
    Public Shared Sub moveMeHwnd(ByVal hwnd As Integer)
      Call ReleaseCapture()
      ''Call SendMessage(hwnd, WM_NCLBUTTONDOWN, clickPos, 0&)
      Call SendMessage(hwnd, WM_NCLBUTTONDOWN, 0&, 0&)
    End Sub

    Public Sub MoveMe()
      Call ReleaseCapture()
      Call SendMessage(m_hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End Sub

    Private Sub ClientMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
      ' Automatically allow user to drag using
      ' any portion of form, not just titlebar,
      ' when user depresses left mousebutton.
      ' Useful for captionless forms.
      If Button = MouseButtons.Left Then
        If m_AutoDrag Then
          Call ReleaseCapture()
          Call SendMessage(m_hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End If
      End If
    End Sub

    Private Sub ClientMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
      ' Automatically handle system menu display
      ' when user right-clicks anywhere on form.
      ' Useful for captionless forms.
      Dim pt As POINTAPI
      ' This is relative to the screen, so we can't
      ' use the coordinates passed in the event
      Call GetCursorPos(pt)
      If Button = MouseButtons.Right Then
        If m_AutoSysMenu Then
          ' Thanks to Klaus Probst for this trick!
          ' http://www.vbbox.com/
          Call ShowSysMenu(pt.x, pt.Y)
        End If
      End If
    End Sub


    ' ************************************************
    '  Public Properties: Read/Write
    ' ************************************************
:    Public Property AutoDrag() As Boolean
:      Set(ByVal value As Boolean)

        ' Automatically allow user to drag using
        ' any portion of form, not just titlebar,
        ' when user depresses left mousebutton.
        ' Useful for captionless forms.
:        m_AutoDrag = value
:      End Set
:      Get
:        AutoDrag = m_AutoDrag
:      End Get
:    End Property


:    Public Property AutoSysMenu() As Boolean
:      Set(ByVal value As Boolean)
:
:        ' Automatically allow user to drag using
:        ' any portion of form, not just titlebar,
:        ' when user depresses left mousebutton.
:        ' Useful for captionless forms.
:        m_AutoSysMenu = value
:      End Set
:      Get
:        AutoSysMenu = m_AutoSysMenu
:      End Get
:    End Property
:
:    Public Property client() As Object
:      Get
:        ' Return reference to client.
:        If Not m_Client Is Nothing Then
:          client = m_Client
:        End If
:     End Get
:      Set(ByVal value As Object)
:        ' Clear cached handle.
:        m_hWnd = 0

        ' Store object reference and handle to client.
:        If TypeOf value Is Form Then
:          m_Client = value
:          m_hWnd = m_Client.Handle
:        End If
:      End Set
:    End Property

:    Public Property controlBox() As Boolean
      Get
        ' Return value of WS_SYSMENU bit.
        controlBox = CBool(Style() And WS_SYSMENU)
      End Get
      Set(ByVal value As Boolean)
        ' Set WS_SYSMENU On or Off as requested.
        Call FlipBit(WS_SYSMENU, value)
      End Set
:    End Property

:    Public Property MaxButton() As Boolean
      Get
        ' Return value of WS_MAXIMIZEBOX bit.
        MaxButton = CBool(Style() And WS_MAXIMIZEBOX)
      End Get
      Set(ByVal value As Boolean)
        ' Set WS_MAXIMIZEBOX On or Off as requested.
        Call FlipBit(WS_MAXIMIZEBOX, value)
        Call EnableMenuItem(SysMenuItems.smMaximize, value)
      End Set
:    End Property

:    Public Property MinButton() As Boolean
      Get
        ' Return value of WS_MINIMIZEBOX bit.
        MinButton = CBool(Style() And WS_MINIMIZEBOX)
      End Get
      Set(ByVal value As Boolean)
        ' Set WS_MINIMIZEBOX On or Off as requested.
        Call FlipBit(WS_MINIMIZEBOX, value)
        Call EnableMenuItem(SysMenuItems.smMinimize, value)
      End Set
    End Property

:    Public Property Moveable() As Boolean
      Get
        ' Return whether SC_MOVE menu is enabled.
        Moveable = Not CBool(GetMenuItemState( _
           GetSystemMenu(m_hWnd, False), _
           GetMenuItemPosition(SysMenuItems.smMove)))
      End Get
      Set(ByVal value As Boolean)
        ' Toggle SC_MOVE menu appropriately.
        Call EnableMenuItem(SysMenuItems.smMove, value)
      End Set
    End Property

:    Public Property Sizeable() As Boolean
      Get
        ' Return value of WS_THICKFRAME bit.
        Sizeable = CBool(Style() And WS_THICKFRAME)
      End Get
      Set(ByVal value As Boolean)
        ' Toggle SC_SIZE menu appropriately,
        ' or else it gets removed!
        Call EnableMenuItem(SysMenuItems.smSize, value)
        ' Set WS_THICKFRAME On or Off as requested.
        Call FlipBit(WS_THICKFRAME, value)
      End Set
    End Property
    
:    Public Property ShowInTaskbar() As Boolean
      Get
        ' Return value of WS_EX_APPWINDOW bit.
        ShowInTaskbar = CBool(StyleEx() And WS_EX_APPWINDOW)
      End Get
      Set(ByVal value As Boolean)
        ' Set WS_EX_APPWINDOW On or Off as requested.
        ' Toggling this value requires that we also toggle
        ' visibility, flipping the bit while invisible,
        ' forcing the taskbar to update on reshow.
        ' Using LockWindowUpdate prevents some flicker.
        Call LockWindowUpdate(m_hWnd)
        Call ShowWindow(m_hWnd, vbHide)
        Call FlipBitEx(WS_EX_APPWINDOW, value)
        Call ShowWindow(m_hWnd, vbNormalFocus)
        Call LockWindowUpdate(0&)
      End Set
    End Property



:    Public Property Titlebar() As Boolean
      Get
        ' Return value of WS_CAPTION bit.
        Titlebar = CBool(Style() And WS_CAPTION)
      End Get
      Set(ByVal value As Boolean)
        ' Set WS_CAPTION On or Off as requested.
        Call FlipBit(WS_CAPTION, value)
      End Set
    End Property

:    Public Property ToolWindow() As Boolean
      Get
        ' Return value of WS_EX_TOOLWINDOW bit.
        ToolWindow = CBool(StyleEx() And WS_EX_TOOLWINDOW)
      End Get
      Set(ByVal value As Boolean)
        ' Set WS_EX_TOOLWINDOW On or Off as requested.
        Call FlipBitEx(WS_EX_TOOLWINDOW, value)
      End Set
    End Property

:    Public Property TopMost() As Boolean
      Get
        ' Return value of WS_EX_TOPMOST bit.
        TopMost = CBool(StyleEx() And WS_EX_TOPMOST)
      End Get
      Set(ByVal value As Boolean)
        Const swpFlags = SWP_NOMOVE Or SWP_NOSIZE
        ' Unlike most style bits, WS_EX_TOPMOST must be
        ' set with SetWindowPos rather than SetWindowLong.
        If value Then
          Call SetWindowPos(m_hWnd, HWND_TOPMOST, 0, 0, 0, 0, swpFlags)
        Else
          Call SetWindowPos(m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, swpFlags)
        End If
        ' Additional references on VB use of SetWindowPos...
        ' BUG: SetWindowPos API Does Not Set Topmost Window in VB
        ' -- http://support.microsoft.com/support/kb/articles/Q192/2/54.ASP
        ' FIX: TopMost Window Does Not Stay on Top in Design Environment
        ' -- http://support.microsoft.com/support/kb/articles/Q150/2/33.ASP
      End Set
    End Property

:    Public Property WhatsThisButton() As Boolean
      Get
        ' Return value of WS_EX_CONTEXTHELP bit.
        WhatsThisButton = CBool(StyleEx() And WS_EX_CONTEXTHELP)
      End Get
      Set(ByVal value As Boolean)
        ' Set WS_EX_CONTEXTHELP On or Off as requested.
        Call FlipBitEx(WS_EX_CONTEXTHELP, value)
      End Set
    End Property


    ' ************************************************
    '  Public Properties: Read-only

:    Public ReadOnly Property hwnd() As Boolean
      Get
        hwnd = m_hWnd
      End Get
    End Property

    ' ************************************************
    '  Public Methods
    ' ************************************************
    Public Sub EnableMenuItem(ByVal MenuItem As SysMenuItems, Optional ByVal Enabled As Boolean = True)
      ' This routine is automatically called whenever the
      ' MinButton, MaxButton, or Movable properties are
      ' set.
      Dim hMenu As Integer
      Dim nPosition As Integer
      'Dim uFlags As Integer
      Dim mii As MENUITEMINFO
      Const HighBit As Integer = &H8000&

      ' Retrieve handle to system menu.
      hMenu = GetSystemMenu(m_hWnd, False)

      ' Translate ID to position.
      nPosition = GetMenuItemPosition(MenuItem)
      If nPosition >= 0 Then

        ' Initialize structure.
        mii.cbSize = Len(mii)
        mii.fMask = MIIM_STATE Or MIIM_ID Or MIIM_DATA Or MIIM_TYPE
        mii.dwTypeData = StrDup(80, vbNullChar)
        mii.cch = Len(mii.dwTypeData)
        Call GetMenuItemInfo(hMenu, nPosition, MF_BYPOSITION, mii)

        ' Set appropriate state.
        If Enabled Then
          mii.fState = MF_ENABLED
        Else
          mii.fState = MF_GRAYED
        End If

        ' New ID uses highbit to signify that
        ' the menu item is enabled.
        If Enabled Then
          mii.wID = MenuItem
        Else
          mii.wID = MenuItem And Not HighBit
        End If

        ' Modify the menu!
        mii.fMask = MIIM_STATE Or MIIM_ID
        Call SetMenuItemInfo(hMenu, nPosition, MF_BYPOSITION, mii)
      End If
    End Sub

    Public Sub Redraw()
      ' Redraw window with new style.
      Const swpFlags As Integer = _
         SWP_FRAMECHANGED Or SWP_NOMOVE Or _
         SWP_NOZORDER Or SWP_NOSIZE
      SetWindowPos(m_hWnd, 0, 0, 0, 0, 0, swpFlags)
    End Sub

    Public Sub RemoveMenuItem(ByVal MenuItem As SysMenuItems)
      Dim hMenu As Integer

      ' Retrieve handle to system menu.
      hMenu = GetSystemMenu(m_hWnd, False)

      ' Special case processing...
      Select Case MenuItem
        Case SysMenuItems.smClose
          ' when removing the Close menu, also
          ' remove the separator over it
          RemoveMenu(hMenu, _
             GetMenuItemPosition(MenuItem) - 1, _
             MF_BYPOSITION)

        Case SysMenuItems.smMinimize
          ' Ensure buttons are consistent.
          Me.MinButton = False

        Case SysMenuItems.smMaximize
          ' Ensure buttons are consistent.
          Me.MaxButton = False
      End Select

      ' Remove requested entry.
      Call RemoveMenu(hMenu, MenuItem, MF_BYCOMMAND)
    End Sub

    Public Sub ShowSysMenu(ByVal x As Integer, ByVal Y As Integer)
      ' Must be in screen coordinates.
      Call SendMessage(m_hWnd, WM_GETSYSMENU, 0, MakeLong(Y, x))
    End Sub

    ' ************************************************
    '  Private Methods
    ' ************************************************
    Private Function MakeLong(ByVal WordHi As Integer, ByVal WordLo As Integer) As Integer
      ' High word is coerced to Long to allow it to
      ' overflow limits of multiplication which shifts
      ' it left.
      MakeLong = (CLng(WordHi) * &H10000) Or (WordLo And &HFFFF&)
    End Function

    Private Function Style(Optional ByVal NewBits As Integer = 0) As Integer
      ' Attempt to set new style bits.
      If NewBits Then
        Call SetWindowLong(m_hWnd, GWL_STYLE, NewBits)
      End If
      ' Retrieve current style bits.
      Style = GetWindowLong(m_hWnd, GWL_STYLE)
    End Function

    Private Function StyleEx(Optional ByVal NewBits As Integer = 0) As Integer
      ' Attempt to set new style bits.
      If NewBits Then
        Call SetWindowLong(m_hWnd, GWL_EXSTYLE, NewBits)
      End If
      ' Retrieve current style bits.
      StyleEx = GetWindowLong(m_hWnd, GWL_EXSTYLE)
    End Function

    Private Function FlipBit(ByVal Bit As Integer, ByVal Value As Boolean) As Boolean
      Dim nStyle As Integer

      ' Retrieve current style bits.
      nStyle = GetWindowLong(m_hWnd, GWL_STYLE)

      ' Attempt to set requested bit On or Off,
      ' and redraw
      If Value Then
        nStyle = nStyle Or Bit
      Else
        nStyle = nStyle And Not Bit
      End If
      Call SetWindowLong(m_hWnd, GWL_STYLE, nStyle)
      Call Redraw()

      ' Return success code.
      FlipBit = (nStyle = GetWindowLong(m_hWnd, GWL_STYLE))
    End Function

    Private Function FlipBitEx(ByVal Bit As Integer, ByVal Value As Boolean) As Boolean
      Dim nStyleEx As Integer

      ' Retrieve current style bits.
      nStyleEx = GetWindowLong(m_hWnd, GWL_EXSTYLE)

      ' Attempt to set requested bit On or Off,
      ' and redraw.
      If Value Then
        nStyleEx = nStyleEx Or Bit
      Else
        nStyleEx = nStyleEx And Not Bit
      End If
      Call SetWindowLong(m_hWnd, GWL_EXSTYLE, nStyleEx)
      Call Redraw()

      ' Return success code.
      FlipBitEx = (nStyleEx = GetWindowLong(m_hWnd, GWL_EXSTYLE))
    End Function

    Private Function GetMenuItemPosition(ByVal MenuItem As SysMenuItems) As Integer
      Dim hMenu As Integer
      Dim ID As Integer
      Dim i As Integer
      Const HighBit As Integer = &H8000&

      ' Default to returning -1 in case of
      ' failure, since menu is 0-based.
      GetMenuItemPosition = -1

      ' Retrieve handle to system menu.
      hMenu = GetSystemMenu(m_hWnd, False)

      ' Loop through system menu, scanning
      ' for requested standard menu item.
      For i = 0 To GetMenuItemCount(hMenu) - 1
        ID = GetMenuItemID(hMenu, i)
        If ID = MenuItem Then
          ' Return position of normal
          ' enabled menu item.
          GetMenuItemPosition = i
          Exit For
        ElseIf ID = (MenuItem And Not HighBit) Then
          ' This item is disabled.
          ' Return position and alter
          ' MenuItem with new ID.
          MenuItem = ID
          GetMenuItemPosition = i
          Exit For
        End If
      Next i
    End Function

    Private Function GetMenuItemText(ByVal hMenu As Integer, ByVal nPosition As Integer) As String
      Dim mii As MENUITEMINFO

      ' Initialize structure.
      mii.cbSize = Len(mii)
      mii.fMask = MIIM_TYPE
      mii.dwTypeData = StrDup(80, vbNullChar)
      mii.cch = Len(mii.dwTypeData)
      Call GetMenuItemInfo(hMenu, nPosition, MF_BYPOSITION, mii)

      ' Return current menu text
      If mii.cch > 0 Then
        GetMenuItemText = Left$(mii.dwTypeData, mii.cch)
      End If
    End Function

    Private Function GetMenuItemState(ByVal hMenu As Integer, ByVal nPosition As Integer) As Integer
      Dim mii As MENUITEMINFO

      ' Initialize structure.
      mii.cbSize = Len(mii)
      mii.fMask = MIIM_STATE
      Call GetMenuItemInfo(hMenu, nPosition, MF_BYPOSITION, mii)

      ' Return current state.
      GetMenuItemState = mii.fState
    End Function




  End Class
  
  
 
 
'-->
'--> F I N D - W I N D O W - W I L D  ???
  
''http://kellychronicles.spaces.live.com/blog/cns!A0D71E1614E8DBF8!217.entry
''vb.net enumerate all top level windows netframework

#Imports System.Collections.Generic
#Imports System.Runtime.InteropServices
#Imports System.Text
 
Public Class ApiWindow
  Public MainWindowTitle As String = ""
  Public ClassName As String = ""
  Public hWnd As Int32
End Class
 
''' <summary> 
''' Enumerate top-level and child windows 
''' </summary> 
''' <example> 
''' Dim enumerator As New WindowsEnumerator()
''' For Each top As ApiWindow in enumerator.GetTopLevelWindows()
'''    Console.WriteLine(top.MainWindowTitle)
'''    For Each child As ApiWindow child in enumerator.GetChildWindows(top.hWnd) 
'''        Console.WriteLine(" " + child.MainWindowTitle)
'''    Next child
''' Next top
''' </example> 
Public Class WindowsEnumerator
 
  Private Delegate Function EnumCallBackDelegate(ByVal hwnd As Integer, ByVal lParam As Integer) As Integer
 
  ' Top-level windows.
  Private Declare Function EnumWindows Lib "user32" _
   (ByVal lpEnumFunc As EnumCallBackDelegate, ByVal lParam As Integer) As Integer
 
  ' Child windows.
  Private Declare Function EnumChildWindows Lib "user32" _
   (ByVal hWndParent As Integer, ByVal lpEnumFunc As EnumCallBackDelegate, ByVal lParam As Integer) As Integer
 
  ' Get the window class.
  Private Declare Function GetClassName _
   Lib "user32" Alias "GetClassNameA" _
   (ByVal hwnd As Integer, ByVal lpClassName As StringBuilder, ByVal nMaxCount As Integer) As Integer
 
  ' Test if the window is visible--only get visible ones.
  Private Declare Function IsWindowVisible Lib "user32" _
   (ByVal hwnd As Integer) As Integer
 
  ' Test if the window's parent--only get the one's without parents.
  Private Declare Function GetParent Lib "user32" _
   (ByVal hwnd As Integer) As Integer
 
  ' Get window text length signature.
  Private Declare Function SendMessage _
   Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Int32, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As Int32) As Int32
 
  ' Get window text signature.
  Private Declare Function SendMessage _
   Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Int32, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As StringBuilder) As Int32
 
  Private _listChildren As New List(Of ApiWindow)
  Private _listTopLevel As New List(Of ApiWindow)
 
  Private _topLevelClass As String = ""
  Private _childClass As String = ""
 
  
  
  '-->
  '--> H E L P E R
  
  ''' <summary>
  ''' Get all top-level window information
  ''' </summary>
  ''' <returns>List of window information objects</returns>
  Public Overloads Function GetTopLevelWindows() As List(Of ApiWindow)
 
    EnumWindows(AddressOf EnumWindowProc, &H0)
 
    Return _listTopLevel
 
  End Function
 
  Public Overloads Function GetTopLevelWindows(ByVal className As String) As List(Of ApiWindow)
 
    _topLevelClass = className
 
    Return Me.GetTopLevelWindows()
 
  End Function
 
  ''' <summary>
  ''' Get all child windows for the specific windows handle (hwnd).
  ''' </summary>
  ''' <returns>List of child windows for parent window</returns>
  Public Overloads Function GetChildWindows(ByVal hwnd As Int32) As List(Of ApiWindow)
 
    ' Clear the window list.
    _listChildren = New List(Of ApiWindow)
 
    ' Start the enumeration process.
    EnumChildWindows(hwnd, AddressOf EnumChildWindowProc, &H0)
 
    ' Return the children list when the process is completed.
    Return _listChildren
 
  End Function
 
  Public Overloads Function GetChildWindows(ByVal hwnd As Int32, ByVal childClass As String) As List(Of ApiWindow)
 
    ' Set the search
    _childClass = childClass
 
    Return Me.GetChildWindows(hwnd)
 
  End Function
 
  ''' <summary>
  ''' Callback function that does the work of enumerating top-level windows.
  ''' </summary>
  ''' <param name="hwnd">Discovered Window handle</param>
  ''' <returns>1=keep going, 0=stop</returns>
:  Private Function EnumWindowProc(ByVal hwnd As Int32, ByVal lParam As Int32) As Int32
 
    ' Eliminate windows that are not top-level.
    If GetParent(hwnd) = 0 AndAlso CBool(IsWindowVisible(hwnd)) Then
 
      ' Get the window title / class name.
      Dim window As ApiWindow = GetWindowIdentification(hwnd)
 
      ' Match the class name if searching for a specific window class.
      If _topLevelClass.Length = 0 OrElse window.ClassName.ToLower() = _topLevelClass.ToLower() Then
        _listTopLevel.Add(window)
      End If
 
    End If
 
    ' To continue enumeration, return True (1), and to stop enumeration 
    ' return False (0).
    ' When 1 is returned, enumeration continues until there are no 
    ' more windows left.
 
    Return 1
 
  End Function
 
  ''' <summary>
  ''' Callback function that does the work of enumerating child windows.
  ''' </summary>
  ''' <param name="hwnd">Discovered Window handle</param>
  ''' <returns>1=keep going, 0=stop</returns>
  Private Function EnumChildWindowProc(ByVal hwnd As Int32, ByVal lParam As Int32) As Int32
 
    Dim window As ApiWindow = GetWindowIdentification(hwnd)
 
    ' Attempt to match the child class, if one was specified, otherwise
    ' enumerate all the child windows.
    If _childClass.Length = 0 OrElse window.ClassName.ToLower() = _childClass.ToLower() Then
      _listChildren.Add(window)
    End If
 
    Return 1
 
  End Function
 
  ''' <summary>
  ''' Build the ApiWindow object to hold information about the Window object.
  ''' </summary>
:  Private Function GetWindowIdentification(ByVal hwnd As Integer) As ApiWindow
 
    Const WM_GETTEXT As Int32 = &HD
    Const WM_GETTEXTLENGTH As Int32 = &HE
 
    Dim window As New ApiWindow()
 
    Dim title As New StringBuilder()
 
    ' Get the size of the string required to hold the window title.
    Dim size As Int32 = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
 
    ' If the return is 0, there is no title.
    If size > 0 Then
      title = New StringBuilder(size + 1)
 
      SendMessage(hwnd, WM_GETTEXT, title.Capacity, title)
    End If
 
    ' Get the class name for the window.
    Dim classBuilder As New StringBuilder(64)
    GetClassName(hwnd, classBuilder, 64)
 
    ' Set the properties for the ApiWindow object.
    window.ClassName = classBuilder.ToString()
    window.MainWindowTitle = title.ToString()
    window.hWnd = hwnd
 
    Return window
 
  End Function
 
End Class
'End Namespace




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









