
'-->
'--> G L O B A L 



#Para SilentMode true
#Para DebugMode intern
#runtime EXE

'--> ....

'NEU...
#Reference System.Windows.Forms.dll
#Reference System.Data.dll
#Reference System.Drawing.dll
#Reference WeifenLuo.WinFormsUI.Docking.dll
#Reference TenTec.Windows.iGridLib.iGrid.dll

#Reference WindowsHookLib.dll

#Imports System.Windows.Forms
#Imports System.Data
#Imports System.Drawing
#Imports ScriptIDE.Core
#Imports ScriptIDE.ScriptHost
#Imports ScriptIDE.ScriptWindowHelper
#Imports Tentec.Windows.iGridLib


#Reference Microsoft.VisualBasic.dll
#Imports Microsoft.VisualBasic

#Imports System.Windows.Forms.Form

''# Reference C:\yEXE\wb_test_es1.dll
#Reference C:\yEXE\Skybound.Gecko.dll
''# Imports wb_test_es1.geckoWB

'#Imports Skybound.Gecko
#Imports System.Collections.Generic








'''#####################################################




public iconList() as string
public globMonitorRef as object

dim iproc as sys_interproc

public Shared meScriptClass as object
public WithEvents FormRef As Form
Public shared  globSkipNewWindow as boolean =false
public panelRef as object
public toolBarColor="#334565"

Dim glob As cls_globPara
public lastHeight as integer=0
public lastWidth as integer=0
 

dim WithEvents FlowLayoutPanel1 As System.Windows.Forms.FlowLayoutPanel





Dim WithEvents Timer1 As System.Windows.Forms.Timer

Public shared  WithEvents WB2 as Skybound.Gecko.GeckoWebBrowser
Public shared  WithEvents WB3 as Skybound.Gecko.GeckoWebBrowser
'Public shared  WithEvents WB2 as wbEx
'Public shared  WithEvents WB3 as wbEx

Structure wbDataStruc
   public id as string
   public nicName as string
   public url as string
   public wbEx as object
   public deltaLeft as integer
   public deltaRight as integer
   public visible as boolean
End Structure

Public shared wbData As New Dictionary(Of String, wbDataStruc)
Public shared curBrowserKey as string
public shared WithEvents check1 as System.Windows.Forms.CheckBox
public shared WithEvents check2 as System.Windows.Forms.CheckBox
public shared WithEvents check3 as System.Windows.Forms.CheckBox

public shared toolBar  As ScriptedPanel
public shared statBar  As ScriptedPanel
public shared sideBar  As ScriptedPanel
public shared container1  As ScriptedPanel
public shared popUp  As ScriptedPanel

public shared trace1  As object
public shared trace2 As object
public shared trace3  As object
public shared trace4  As object
public shared trace5  As object
public shared trace6  As object
public shared trace7  As object
public shared trace8  As object
public shared myTextArea as textBox


dim WithEvents myPicture as pictureBox 
dim WithEvents myPicture2 as pictureBox 

Private WithEvents kbHook As New KeyboardHook


public shared splitA1  As Object
Dim splitA1_Left, splitA1_Right As ScriptedPanel


public shared splitB2  As Object
Dim splitB2_Left, splitB2_Right As ScriptedPanel

public shared splitB3  As Object
Dim splitB3_Left, splitB3_Right As ScriptedPanel


public shared splitC4  As Object
Dim splitC4_Left, splitC4_Right As ScriptedPanel

public shared splitC5  As Object
Dim splitC5_Left, splitC5_Right As ScriptedPanel

public shared splitC6  As Object
Dim splitC6_Left, splitC6_Right As ScriptedPanel



Public Declare Function GetTime Lib "winmm.dll" Alias "timeGetTime" () As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
Public Declare Function ReleaseCapture Lib "user32" () As Integer
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2



Private Declare Function getActiveWindow Lib "user32.dll" Alias "GetActiveWindow" () As Integer

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
'--> F U N C . . . 

Function createFlowPanel(byRef container,byRef el)
  el = New System.Windows.Forms.FlowLayoutPanel

  'el.Location = New System.Drawing.Point(150, -6)
  'el.Size = New System.Drawing.Size(2222, 38)

  el.Location = New System.Drawing.Point(0, 100)
  el.Size = New System.Drawing.Size(75, 150)

  el.Name = "FlowLayoutPanel1"
  el.TabIndex = 1
  el.BackColor = ColorTranslator.FromHtml("#2C3C58")
  el.BackColor = ColorTranslator.FromHtml("#62645F")
  
  el.visible = true
  container.Controls.Add(el)
  return el
end function


Function createLinkLabel(byRef container, nicName, url)
    dim el as New System.Windows.Forms.LinkLabel
    el.AutoSize = True
    el.BackColor = System.Drawing.Color.Transparent
    el.BackColor = ColorTranslator.FromHtml("#536585")
    el.BackColor = ColorTranslator.FromHtml("#62645F")
    el.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline
    el.LinkColor = ColorTranslator.FromHtml("#ccc")
    el.Location = New System.Drawing.Point(3, 0)
    'el.Padding = New System.Windows.Forms.Padding(2, 10, 2, 4)
    'el.Margin = New System.Windows.Forms.Padding(0, 0, 3, 0)

    ''el.Size = New System.Drawing.Size(63, 33)
    el.Size = New System.Drawing.Size(63, 22)

    'el.TabIndex = 0
    el.TabStop = True
    el.Text = nicName
    el.tag=url
    container.Controls.Add(el)
    addHandler el.mouseClick, AddressOf onLinkLabel_mouseClick
    return el
end Function


function createWbEx(container as object, key as string,url as string)
  '--> WebBrowser erzeugen
  trace "createWbEx", "START"
  dim id as string=key
  dim newBrowser=New wbEx
  
  curBrowserKey=key
  WB2 = newBrowser
  WB2.name="main"
  curBrowserKey=key
  
  dim DATA as wbDataStruc
  if not wbData.ContainsKey(key) then
     DATA=new wbDataStruc
     wbData.Add(key, DATA)
  else
     DATA=wbData.item(key)
  end if
  DATA.id=key
  DATA.url=url
  DATA.wbEx=WB2
  trace "curBrowserKey", DATA.id
  trace "curBrowserURL", DATA.url
  
  wbData.item(key)=DATA
  
  container.Controls.Add(WB2)
  wb2.Dock = DockStyle.Fill

  ''''''WB2.Navigate("http://www.google.de/ie")
  WB2.visible=true
  WB2.Navigate(url)
  trace "autostart", "5555555555555"
  WB2.bringtoFront
end Function


function createWbEx2(container as object, key as string,url as string)
  '--> WebBrowser erzeugen
  trace "createWbEx", "START"
  dim id as string=key
  dim newBrowser=New wbEx
  
  curBrowserKey=key
  WB3 = newBrowser
  WB3.name="sidebar"
  'msgBox(WB3.name)
  curBrowserKey=key
  
  dim DATA as wbDataStruc
  if not wbData.ContainsKey(key) then
     DATA=new wbDataStruc
     wbData.Add(key, DATA)
  else
     DATA=wbData.item(key)
  end if
  DATA.id=key
  DATA.url=url
  DATA.wbEx=WB2
  trace "curBrowserKey", DATA.id
  trace "curBrowserURL", DATA.url
  
  wbData.item(key)=DATA
  
  container.Controls.Add(WB3)
  wb3.Dock = DockStyle.Fill

  ''''''WB2.Navigate("http://www.google.de/ie")
  WB3.visible=true
  WB3.Navigate(url)
  trace "autostart", "5555555555555"
  WB3.bringtoFront
end Function





function wbExNavigate(key as string)
  if not wbData.ContainsKey(key) then
    msgbox("Id "+key+"...nicht gefunden")
    exit function
  end if
  curBrowserKey=key
  dim DATA=wbData.item(key)
  dim curBrowser = DATA.wbEx
  WB2 = curBrowser
  curBrowser.bringtofront

end Function





'-->
'--> W I N - F O R M - H A N D L I N G



Sub globAddHandler()
  'AddHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  'AddHandler Timer1.Tick, AddressOf Timer1_Tick
  AddHandler formRef.PreviewKeyDown,       AddressOf Form_PreviewKeyDown
  AddHandler formRef.Resize,               AddressOf Form_Resize
  AddHandler formRef.GotFocus,             AddressOf Form_GotFocus
  AddHandler formRef.Activated,            AddressOf Form_Activated
  AddHandler Timer1.Tick, AddressOf onTimerEvent
  
end sub


Sub globRemoveHandler()
  trace "globRemoveHandler","try..."
  if formRef is nothing then exit sub
  'RemoveHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  'RemoveHandler Timer1.Tick, AddressOf Timer1_Tick
  RemoveHandler formRef.PreviewKeyDown,    AddressOf Form_PreviewKeyDown
  RemoveHandler formRef.Resize,            AddressOf Form_Resize
  RemoveHandler formRef.GotFocus,          AddressOf Form_GotFocus
  RemoveHandler formRef.Activated,         AddressOf Form_Activated
  RemoveHandler Timer1.Tick, AddressOf onTimerEvent
  trace "globRemoveHandler","DONE"
end sub


:Sub onResizeControls()
    dim wbContainer=splitA1_Left
    'container1.left=0

    'container1.top=0
    container1.top=toolbar.top+toolbar.height+2
 
    dim sidebarWidth=sidebar.width
    container1.left=sidebarWidth
    'container1.width=panelRef.width-sidebarWidth
    
    
    'container1.height= panelRef.Height-toolbar.height-statbar.Height
    'container1.height= panelRef.Height-statbar.Height-container1.top
    'container1.width= panelRef.width-sidebarWidth-container1.left
    container1.height= panelRef.Height-statbar.Height-container1.top
    container1.width= panelRef.width-container1.left
 
    'WB2.Height = panelRef.Height - 115 '75
    ''WB2.Top = 80:  WB2.Height = container1.Height - 115 '75
    
    ''WB2.Top = 0:  ' WB2.Height = 444 ' container1.Height - 28 '75
    ''WB2.left = 0 '/WB.Top
    ''WB2.Width = wbContainer.Width - 0
    ''WB2.Height = wbContainer.Height - 0 '75
    '' WB2.Top = 80:  WB2.Height = panelRef.Height - 115 '75
    '' WB2.Top = 28:  WB2.Height = panelRef.Height - 28 '75
   
    'panelRef.Controls("txt_status").Top = panelRef.Height - 25
    'panelRef.Controls("txt_status").Width = panelRef.Width - 25
    container1.bringtofront
    sideBar.bringToFront()
    toolbar.bringToFront()
    if popup.visible=true then popup.bringToFront()
  End Sub



Sub Frm_FormClosing(Sender As Object,e As System.Windows.Forms.FormClosingEventArgs) Handles FormRef.FormClosing
  ''glob.saveFormPos(FormRef)
  ''glob.saveParaFile()
  saveGlobPara()
  globRemoveHandler()   
End Sub


Sub onTerminate()
  trace "es_webbrowser", "onTerminate"
  globRemoveHandler() 
  
  'saveGlobPara
  
  'globRemoveHandler()
'If VLC1.playlist.isPlaying Then VLC1.playlist.stop()
  'stop timer  : fehlen noch
  'stop resizer 
End Sub

sub saveGlobPara()
  trace "saveGlobPara","try..."
  if glob is nothing then exit sub
  if formRef is nothing then exit sub
  glob.saveFormPos(FormRef)
  '' glob.saveTuttiFrutti(FormRef)
  glob.para("lastHeight")=lastHeight
  glob.para("lastWidth")=lastWidth
  glob.saveParaFile()
  '' trace "saveGlobPara","DONE"
 
  
  
end sub


Sub sys_interproc_Message(source As String, cmdString As String, para As String) ''... Handles sys_interproc.Message
  If cmdString = "navigateUrl" Then
   trace "sys_interproc_Message(web_browser)", para
   navigateWebbrowser(para)
   FormRef.show
   formRef.BringToFront()
   formref.TopMost=true
   formref.TopMost=false
  End If
End Sub





sub startTimer
  timer1.Enabled = True
end sub



sub stopTimer
  timer1.Enabled = false
end sub



 : Sub onTimerEvent(ByVal sender As Object, ByVal e As System.EventArgs)
 on error resume next
   ''msgBox ("onTimerEvent")

   checkForShowPopupWindow(e)
   
   dim deltaTime=GetTime()-globMouseHookTimerStart
   if deltaTime>2222then 
     stopTimer()
   end if

End Sub










'-->
'--> A U T O S T A R T

Sub AutoStart()
  on error resume next
  dim el 
  iproc=New sys_interproc("web_browser")
  iproc.Commands.Add("CMD  ", "navigateUrl", "url|||PARA???", "Para gibts noch nicht")
  AddHandler iproc.Message, AddressOf sys_interproc_Message

  Timer1 = New System.Windows.Forms.Timer()
  timer1.Interval = 55
  timer1.Enabled = false

  ''Skybound.Gecko.Xpcom.Initialize("C:\yEXE\xulrunner\")

  ''Skybound.Gecko.Xpcom.Initialize("C:\Programme\Mozilla Firefox\")

  Skybound.Gecko.Xpcom.Initialize("C:\Programme\Mozilla Developer Preview 3.7 Alpha 4\")
   
  meScriptClass = me
  printLine 1, ".", ""
  printLine 2, "..", ""
  printLine 3, "...", ""
  printLine 4, "....", ""
  printLine 1, "autostart", now.tostring
  printLine 5, ".....", ""
  printLine 6, "......", ""
  printLine 7, ".......", ""
  printLine 8, "........", ""

'--> Formreferenz holen  
  dim winId as string="es_webbrowser3.tab"
  'panelRef= ZZ.IDEHelper.CreateDockWindow("{ScriptClass}.tab")
  'PanelRef = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)'  : err.number=0
  
  'PanelRef = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)'  : err.number=0
  PanelRef = ZZ.CreateWindow(DWndFlags.StdWindow Or DWndFlags.NoAutoShow, 5)

 ''FormRef = ZZ.getDockPanelRef("Toolbar|##|tbScriptWin|##|webbrowser1.tab")'AsSystem.Windows.Forms,System.Windows.Forms.Form
  FormRef=panelRef.form
  
  'formRef.visible=false
  FormRef.left=-5555
  FormRef.Text = "==> W-E-B"
  
  
  dim labelText as string
  with panelRef
  .resetControls  (10,5,1)
   ''toolBar     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=75)
   popUp       =.addPanel("popUp"        )
   toolBar     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=77)
   'statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
   statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=1)
   sidebar     =.addPanel("sideBar"      , DockStyle.left, intWidth:=60)
   container1  =.addPanel("container1"   , )
   popUp.resetControls  (10,5,1)
   toolBar.resetControls  (10,5,1)
   statBar.resetControls  (10,5,1)
   sideBar.resetControls  (10,5,1)
   container1.resetControls  (10,5,1)
   
   toolbar.BackColor = ColorTranslator.FromHtml(toolBarColor)
   sideBar.BackColor = ColorTranslator.FromHtml("#62645F")
   sideBar.BackgroundImage=ZZ.GetImageCached("http://snap.teamwiki.net/100518-144233-es-bg-grau-schwarz.png")
   container1.BackColor = ColorTranslator.FromHtml("#fff")
   statBar.BackColor = ColorTranslator.FromHtml("#fff")
  end with
  
  
  '-->splitContainer_A1
  with  container1
    el=.addSplitcontainer("splitA1", "left", splitA1_Left, "right", splitA1_Right, DockStyle.Fill)
    el.Orientation = System.Windows.Forms.Orientation.Horizontal
    el.backColor=ColorTranslator.FromHtml("#88f")
    'el.IsSplitterFixed = True
    el.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
    
    el.splitterDistance=container1.height-66 ''   ... kommt später nochmal
    el.width=container1.width ''   ... kommt später nochmal
    
    el.Panel2Collapsed = True
    el.BackColor = ColorTranslator.FromHtml("#90918E")
    el.Panel1.BackColor =ColorTranslator.FromHtml("#FFFFFF")
    el.Panel2.BackColor = ColorTranslator.FromHtml("#62645F")

    el.SplitterWidth = 6
    splitA1=el
  end with

  
  '-->splitContainer_B2
  with  splitA1_Left
    el=.addSplitcontainer("splitB2", "left", splitB2_Left, "right", splitB2_Right, DockStyle.Fill)
    'el.Orientation = System.Windows.Forms.Orientation.Horizontal
    el.backColor=ColorTranslator.FromHtml("#88f")
    'el.IsSplitterFixed = True
    el.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
    ''el.splitterDistance=450  kommt später
    el.Panel1Collapsed = True
    el.BackColor = ColorTranslator.FromHtml("#306BF8")
    el.SplitterWidth = 6
    el.Panel1.BackColor =ColorTranslator.FromHtml("#FFFFFF")
    el.Panel2.BackColor = ColorTranslator.FromHtml("#FFFFFF")
    splitB2=el
  end with


  '--> popup
  
  with  popUp
    '.BackColor = ColorTranslator.FromHtml("#D4D0C8")
    '.BackColor = ColorTranslator.FromHtml("#BCBCBC")
    .BackgroundImage =ZZ.GetImageCached("http://www.allcrunchy.com/Web_Stuff/Gradient_Generator/images/000000444444120sinusoidalFull.png")
    .foreColor = ColorTranslator.FromHtml("#000")
    .BackColor = ColorTranslator.FromHtml("#333")
    .foreColor = ColorTranslator.FromHtml("#fff")
    .top=300
    .left=400
    .width=540
    .height=90
    .BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    .visible=false
    ' .visible=true
    '.bringToFront()
    
    .insertX = 25 :.insertY = 12 '28
    .addIcon("cmdGoHome"       , "http://cdn.iconfinder.net/data/icons/diagram/64x64/diagram-06.png")
    .addIcon("expand_right02"  , "http://cdn.iconfinder.net/data/icons/snowish/64x64/actions/go-first.png")
    .addIcon("cmdGoBack"       , "http://cdn.iconfinder.net/data/icons/snowish/64x64/actions/back.png"  )
    .addIcon("expand_right04"  , "http://cdn.iconfinder.net/data/icons/snowish/64x64/actions/gtk-about.png")
    ''.addIcon("expand_right03"  , "http://cdn.iconfinder.net/data/icons/snowish/72x72/actions/gtk-media-play-ltr.png")
    ''.addIcon("expand_right04"  , "http://cdn.iconfinder.net/data/icons/snowish/72x72/actions/gtk-media-forward-ltr.png")
    .addIcon("cmdGoForward"  , "http://cdn.iconfinder.net/data/icons/snowish/64x64/actions/forward.png")
    .addIcon("expand_right08"  , "http://cdn.iconfinder.net/data/icons/snowish/64x64/actions/finish.png")
    '.addIcon("expand_right05"  , "http://cdn.iconfinder.net/data/icons/snowish/64x64/apps/boot.png")
    .addIcon("expand_right07"  , "http://cdn.iconfinder.net/data/icons/snowish/64x64/devices/3floppy_unmount.png")

  end with

   '--> debugArea - trace

  with splitA1_Right
   .insertX = 100 :.insertY = 5 '28
    labelText ="aaaaaaa"
    el=.addLabel  ("cmdDebugLab_01", labelText ,  ,"",,,150,22) :
    el.backColor=ColorTranslator.FromHtml("#90918E")
    el.foreColor=ColorTranslator.FromHtml("#fff"): 
   'el.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    trace1=el
    
    labelText ="bbbbbbbbbbb"
    el=.addLabel  ("cmdDebugLab_02", labelText ,  ,"",,,150,22) :
    el.backColor=ColorTranslator.FromHtml("#90918E")
    el.foreColor=ColorTranslator.FromHtml("#fff"): 
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    trace2=el
  
    labelText ="ccccccccc"
    el=.addLabel  ("cmdDebugLab_03", labelText ,  ,"",,,150,22) :
    el.backColor=ColorTranslator.FromHtml("#90918E")
    el.foreColor=ColorTranslator.FromHtml("#fff"): 
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    trace3=el
  
    labelText ="dddddddddd"
    el=.addLabel  ("cmdDebugLab_04", labelText ,  ,"",,,150,22) :
    el.backColor=ColorTranslator.FromHtml("#90918E")
    el.foreColor=ColorTranslator.FromHtml("#fff"): 
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    trace4=el
  
   .insertX = 100 :.insertY = 30 '28
    labelText ="eee"
    el=.addLabel  ("cmdDebugLab_04", labelText ,  ,"",,,-2,22) :
    el.backColor=ColorTranslator.FromHtml("#90918E")
    el.foreColor=ColorTranslator.FromHtml("#fff"): 
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    trace5=el
  
  

'--> ...textarea
    .insertX = 100 :.insertY = 7
    .addTextbox ("outMonitor" ,  -2         , "inBox"    , 0  , "#FFFF99", , , "multiline=260")
       myTextArea=splitA1_Right.element("txt_outMonitor")
       myTextArea.WordWrap=false
       myTextArea.scrollbars=System.Windows.Forms.ScrollBars.Vertical
       myTextArea.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

  
  
  end with

  '--> sideBar
   
  With sideBar
      .activateEvents = "|IconMouseDown|   |TextboxKeyDown|"
    '' .insertX = 5 :.insertY = 3 :  
    '' myPicture2 =.addIcon("cmdTerminate_02", "http://es.teamwiki.net/docs/img/icons/Close_Box_Red.png" )

    .insertX = -1 :.insertY = -3 ' 28 '-5
    .insertX = 0 :.insertY = -0 ' 28 '-5
     myPicture= .addIcon("grip"  , "http://es.teamwiki.net/docs/img/backgrounds/grip.jpg" ,0 ,0   )
    .addIcon("expand_right"  , "http://es.teamwiki.net/docs/icons/arrow_right.png" ,7 ,37 ,16,16  )
    .addIcon("expand_left"   , "http://es.teamwiki.net/docs/icons/arrow_left.png"  ,7 ,37 ,16,16  )
 
    .addIcon("expand_down"   , "http://es.teamwiki.net/docs/icons/arrow_down.png" ,35 ,37 ,16,16  )
    .addIcon("expand_up"     , "http://es.teamwiki.net/docs/icons/arrow_up.png"   ,35 ,37 ,16,16  )
    .bringToFront()
  end with
  
   
    check1 = New System.Windows.Forms.CheckBox
    check1.Location = New System.Drawing.Point(2, 55)
    check1.Size = New System.Drawing.Size(55, 19)
    check1.Text = "lockMode"
    'check1.UseVisualStyleBackColor = True
    check1.BackColor = ColorTranslator.FromHtml("#62645F")
    check1.visible=true
    sideBar.Controls.Add(check1)
  
    
    '' createFlowPanel(toolbar,FlowLayoutPanel1)
    '' createFlowPanel(toolbar,FlowLayoutPanel1)
    createFlowPanel(sideBar,FlowLayoutPanel1)
    createLinkLabel(FlowLayoutPanel1,  "Pinboard",      "http://devnet.teamwiki.net/xymemo.html?doc=es-pb-tw-editor01")
    createLinkLabel(FlowLayoutPanel1,  "Editor",        "http://devnet.teamwiki.net/webeditor3.html")
    createLinkLabel(FlowLayoutPanel1,  "google",        "google.de")
    createLinkLabel(FlowLayoutPanel1,  "htmlGrabsch",   "javascript:var%20mes='nachricht...';%20%20%20var%20url1='http://static.teamwiki.net/js/scriptLib01.js';%20var%20url2='http://static.teamwiki.net/js/htmlGrabsch.js';%20%20%20loadScriptOnDemand=function(scriptFileSpec){;%20%20%20var%20script%20=%20document.createElement('SCRIPT');%20%20%20script.src%20=%20scriptFileSpec;%20%20%20document.body.appendChild(script);};%20%20loadScriptOnDemand(url1);%20%20loadScriptOnDemand(url2);")
    createLinkLabel(FlowLayoutPanel1,  "trace2   ",     "javascript:var%20mes='nachricht...';%20%20%20var%20url1='http://teamwiki.de/static/js/scriptLib01.js';%20var%20url2='http://teamwiki.de/static/js/dev_tools.js';%20%20%20loadScriptOnDemand=function(scriptFileSpec){;%20%20%20var%20script%20=%20document.createElement('SCRIPT');%20%20%20script.src%20=%20scriptFileSpec;%20%20%20document.body.appendChild(script);};%20%20loadScriptOnDemand(url1);%20%20loadScriptOnDemand(url2);")
    createLinkLabel(FlowLayoutPanel1,  "Err     ",      "sidebar:chrome://global/content/console.xul")
    createLinkLabel(FlowLayoutPanel1,  "sideBar",       "sidebar:http://es.teamwiki.net/sb2-seite2.html")
    createLinkLabel(FlowLayoutPanel1,  "sideBar(m)",       "http://es.teamwiki.net/sb2-seite2.html")
    


   '' createLinkLabel(FlowLayoutPanel1,  "trace2",        "javascript:var%20mes="nachricht...";%20%20%20var%20url1="http://teamwiki.de/static/js/scriptLib01.js";%20var%20url2="http://teamwiki.de/static/js/dev_tools.js";%20%20%20loadScriptOnDemand=function(scriptFileSpec){;%20%20%20var%20script%20=%20document.createElement("SCRIPT");%20%20%20script.src%20=%20scriptFileSpec;%20%20%20document.body.appendChild(script);};%20%20loadScriptOnDemand(url1);%20%20loadScriptOnDemand(url2);")

    FlowLayoutPanel1.PerformLayout()



  '--> Toolbar erzeugen
  ''With DirectCast(FormRef, Object)

  With toolbar
    .height=0
    .backColor=ColorTranslator.FromHtml("#5D5F5A")
     .activateEvents = "|ButtonMouseClick|   |TextboxKeyDown|"
 
    '.BR  '--------------------------------------------------
    ''.insertY=5: .insertX=5

    .addIcon("nav_back"      , "http://es.teamwiki.net/docs/icons/buttons_06.png"   ,70 ,0 ,32,32  )
    .addIcon("nav_forward"   , "http://es.teamwiki.net/docs/icons/buttons_22.png"   ,100 ,0 ,32,32  )
    
    
 
 
    .br(30) '--------------------------------------------------------
    'id         txtXX   lab             labXX labBG  X     Y   noBR
     .addTextbox ("url",      400   , "URL:"         , 80  , toolBarColor)
     .Controls("txt_url").foreColor=ColorTranslator.FromHtml("#f00")
     
     
     
    .addButton  ("nav",      ">>>" , "#8ca")
    .br(22) '--------------------------------------------------------
    .addTextbox ("phpsuch",  400   , "PHP-Referenz:", 80  , toolBarColor)
    .addButton  ("phpsuch",  "php" , "#8ca")
     el = .addButton  ("vbsuch_1",  "vbAbout"        , "#8ca")
     el.tag="http://visualbasic.about.com/sitesearch.htm?terms=||SUCH||&SUName=visualbasic&TopNode=99"

     el = .addButton  ("vbsuch_2",  "stackOverflow"  , "#8ca")
     el.tag="http://www.google.com/#hl=en&num=100&newwindow=0&q=site%3Ahttp%3A%2F%2Fstackoverflow.com+||SUCH||"
 
      
   
 
    'Hintergrund.toolBar
    .insertY=0: .insertX=0
    '' el=.addLabel  ("lab01", "" ,  ,"#ffffff",,,2222,77) 
    '' el.backColor=ColorTranslator.FromHtml("#334565")
    '' el.backColor=ColorTranslator.FromHtml(toolBarColor)
  End With

'--> ...iconBar

  iniIconList()
  ''insertIconsFromIconList(toolBar)
  insertIconsFromIconList(sidebar)


  with statbar
     .addTextbox ("status",   350   ,                ,     ,       , 0   , 0 )
  End With


  glob = ZZ.newGlobPara()
  glob.readFormPos(FormRef)
  '' 
  'glob.readTuttiFrutti(FormRef)

  
  
  
  globAddHandler()

  onResizeControls()
  
    
  m_FormBorder.client = formRef
  m_FormBorder.Titlebar = False
  
  'm_FormBorder.Moveable=true
  m_FormBorder.AutoDrag=true
  m_FormBorder.showInTaskBar=false
  
  'formRef.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
  'formRef.topmost=true
  
  lastHeight=cInt(glob.para("lastHeight"))
  lastWidth=cInt(glob.para("lastWidth"))
  FormRef.width=lastWidth
  FormRef.height=lastHeight
  
  splitA1.splitterDistance=container1.height -66
  splitB2.splitterDistance=450
   
   
  FormRef.Show()
  'FormRef.visible=true
  
  
  
  err.clear
  err.description="xxx"
    'msgBox(err.description)
  
  ''MouseHook1.InstallHook()': msgBox(err.description)
  'err.clear: msgBox(MouseHook1.State.ToString) :msgBox(err.description)
  
  sideBar.bringToFront()
  
  'WB.Top = 30
  
    '--> navigate
  dim url as string
  url="http://devnet.teamwiki.net/xymemo.html?doc=es-pb-nerv"
  url="google.de"
  dim args() as string=Environment.GetCommandLineArgs()
  'msgBox(args(0))
  'msgbox(args.length)
  if args.length>1 then url=(args(1))
 
  ''msgbox(url)


  '--> create Browser 1, 2
  ''createWbEx(container1, "myKey", url)
  createWbEx2(splitB2_Left, "myKey", "about:blank")
  createWbEx(splitB2_Right, "myKey", "about:blank")
  
  zz.doevents
  zz.doevents
  WB2.Navigate(url)
  
  zz.doevents
  zz.doevents
  'url = "http://es.teamwiki.net/sb2-seite2.html"
  'createWbEx2(splitB2_Left, "myKey", url)
  
  
  popup.bringToFront()

  
End Sub

sub navigateWebbrowser(url)
  WB2.Navigate(url)
  FormRef.show()
  FormRef.visible=true
  formRef.bringToFront()
  '' msgBox("???")
end sub






'-->

'--> ICON-LIST


: Function iniIconList() as string
dim curEditorLine as integer=zLN
#Data myIconList as String
''============================================================================================================================================
cmdAction01   |clearAll    |http://es.teamwiki.net/docs/icon24/emblem-noread.png
cmdAction02   |Debug.Run   |http://es.teamwiki.net/docs/icon24/Good-mark.png 
cmdAction03   |test1          |http://es.teamwiki.net/docs/icon24/white01.png 
cmdAction04   |test2          |http://es.teamwiki.net/docs/icon24/white02.png 
cmdAction05   |test3          |http://es.teamwiki.net/docs/icon24/white03.png 
cmdAction06   |ToolBar|##|tbScriptWin|##|es_direktfenster.mainwin |http://es.teamwiki.net/docs/icon24/24Go.png
              
toggle       |ToolBar|##|tbScriptWin|##|es_testLabor.mainWin     |http://es.teamwiki.net/docs/icons/gnome-monitor3.png
toggle       |Addin|##|siaCodeCompiler|##|SHDebug                |http://es.teamwiki.net/docs/icon24/debug.png
toggle       |ToolBar|##|tbScriptWin|##|es_internSuche.mainWin   |http://es.teamwiki.net/docs/icon24/ViewSearch.png 
toggle       |ToolBar|##|tbScriptWin|##|es_snippetManager.mainWin      |http://es.teamwiki.net/docs/icon24/insert-object.png
toggle       |ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sw_colorPicker.dll|##|sw_colorPicker.sg_colorPicker        |http://icons3.iconfinder.netdna-cdn.com/data/icons/nuvola2/32x32/apps/kwikdisk.png


'' toggle       |ToolBar|##|tbScriptWin|##|es_bookmarks2.mainWin    |http://es.teamwiki.net/docs/icon24/FavouritesGelb.png
toggle       |ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sg_memo.dll|##|root_sg_memo.sg_memo        |http://es.teamwiki.net/docs/icon24/ModifyEdit.png
toggle       |ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\mwTimer.dll|##|mwTimer.sw_analogClock      |http://es.teamwiki.net/docs/icon24/Clock.png

exe          |C:\yEXE\traceMonitor.exe                         |http://es.teamwiki.net/docs/icon24/agentsvr_113-32.png
exe          |C:\yEXE\ide3\IprocViewer.exe                     |http://es.teamwiki.net/docs/icon24/checkered_flag.png
exe          |C:\yEXE\extractAssociatedIcon.exe                |http://es.teamwiki.net/docs/icon24/stock_draw-selection.png 
exe          |C:\yEXE\mwSidebar.exe                            |http://es.teamwiki.net/docs/icon24/SysTrayDemo2_1-32.png 

startTest    |AutoStart        |http://es.teamwiki.net/docs/icon24/pifmgr_14-32.png 

startTest    |test4            |http://es.teamwiki.net/docs/icon24/darkblue01.png
startTest    |test4            |http://es.teamwiki.net/docs/icon24/gtk-edit.png 
startTest    |test4            |http://es.teamwiki.net/docs/icon24/stock_effects-preview.png 
startTest    |test4            |http://es.teamwiki.net/docs/icon24/redhat-home.png
startTest    |test4            |http://es.teamwiki.net/docs/icon24/Textpreview.png 
startTest    |test4            |http://es.teamwiki.net/docs/icon24/stock_effects-preview.png 
startTest    |test4            |http://es.teamwiki.net/docs/icon24/ViewSearch.png 
startTest    |test4            |http://es.teamwiki.net/docs/icon24/stock_calc-cancel.png
startTest    |test4            |http://es.teamwiki.net/docs/icon24/ViewSearch.png 
startTest    |test4            |http://es.teamwiki.net/docs/icon24/ViewSearch.png 
 '' 
 '' 
 '' 010 | http://es.teamwiki.net/docs/icon24/checkered_flag.png 
 '' 011 | http://es.teamwiki.net/docs/icon24/SearchFernglas.png
 '' 012 | http://es.teamwiki.net/docs/icons/navigation-180-white.png
 '' 013 | http://es.teamwiki.net/docs/icons/navigation-270-white.png 
 '' 014 | xxx 
 '' 015 | http://es.teamwiki.net/docs/icons/navigation-090-button.png 
 '' 016 | http://es.teamwiki.net/docs/icons/navigation-180-button.png 
 '' 017 | http://es.teamwiki.net/docs/icons/navigation-270-button.png 
 '' 017 | http://es.teamwiki.net/docs/icon24/Good-mark.png
 '' 018 | xxx http://es.teamwiki.net/docs/icon24/stock_test-mode.png
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
  for i =1 to 25'   max-1
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
  PnlRef.SuspendLayout()
  dim index, left, top, deltaTop, deltaLeft as integer
  dim text, icon, item  as string
  dim DATA() as string
  dim el as object
  ''dim bgColor="#696E73"
  dim bgColor="#62645F" '"#3B5951" '"#2F394E"  '"#4A7CDB"  'E0DDD7
  deltaTop=300
  deltaLeft=4
  for index =0 to 25
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
   left=2
    top=index*29'34
    dim id as string=trim(DATA(0))+"_"+index.toString
    ''monitorAdd(top.toString, left.tostring)
   
     el=PnlRef.addButton(id ,text,bgColor,Left+deltaLeft,Top+deltaTop,45 ,29 ,icon) ' ,flags,handler)
     'el=PnlRef.addButton(id ,text,bgColor,Left+deltaLeft,Top+deltaTop,32 ,32) ' ,flags,handler)
    
     'el.BackgroundImage = ZZ.GetImageCached(icon) ''CType(resources.GetObject("Button1.BackgroundImage"), System.Drawing.Image)
     'el.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom

     el.ImageAlign = System.Drawing.ContentAlignment.MiddleCenter
    'el.TextAlign = System.Drawing.ContentAlignment.BottomCenter
    'el.TextImageRelation = System.Windows.Forms.TextImageRelation.Overlay 'ImageAboveText
    el.foreColor=ColorTranslator.FromHtml("#fff")
    'el.UseVisualStyleBackColor = True
    el.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    el.FlatAppearance.BorderSize = 0
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

  PnlRef.ResumeLayout(False)
  PnlRef.PerformLayout()

end sub


: sub insertIconsFromIndex(ref as object,index as integer, left as integer, top as integer)
  
  dim bgColor as string ="#2C3C58"
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
'--> EVENTS



sub onLinkLabel_mouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
  dim url=sender.tag.toString
  dim such as string
  dim url2 as string
  such ="sidebar:"
  if url.startsWith (such) then
    ''msgBox(such.length)
    url2=mid(url,such.length+1)
    msgBox("-->"+url2+"<--")
    WB3.Navigate(url2)
    toggleSidebar(1)
    exit sub
  end if

  trace "url", url
  WB2.Navigate(url)
End Sub


'' Sub onButtonEvent(e As Object)
''     If e.Sender.Name = "btn_nav" Then
''       WB2.Navigate(toolbar.Controls("txt_url").Text)
''     End If
''     If e.Sender.Name = "btn_phpsuch" Then
''       WB2.Navigate("http://www.php.net/"+toolbar.Controls("txt_phpsuch").Text)
''     End If
''   End Sub
''   
Sub nav(e)
  WB2.Navigate(toolbar.Controls("txt_url").Text)
end Sub

Sub phpsuch(e)
  WB2.Navigate("http://www.php.net/"+toolbar.Controls("txt_phpsuch").Text)
end Sub




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



'-->...

Private Sub kbHook_KeyDown(ByVal Key As System.Windows.Forms.Keys, byRef cancel as boolean) Handles kbHook.KeyDown
   '' msgBox("kbHook_KeyDown: "+key.toString)
   if wb2.url.toString.contains("xymemo") = true then
      'msgBox("xymemo")
      exit sub
   end if
   
   if formref.handle =getActiveWindow() then
      if Key = 27 then 
       cancel=true
        WB2.goBack(): 
     end if
     if Key =112 then 
       cancel=true
       WB2.goForward(): 
     end if
    '' msgbox(Key)
     'cancel=true
   end if
End Sub

Private Sub kbHook_KeyUp(ByVal Key As System.Windows.Forms.Keys) Handles kbHook.KeyUp
    'Debug.WriteLine(Key)
End Sub

Private Sub Form_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs)
  '!!! deer kommt nur beim start --- dann ist tote hose
  'msgBox (e.KeyCode)
End Sub



Sub Form_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
   ''msgBox("Form_GotFocus")
End Sub



Sub Form_Activated(ByVal sender As Object, ByVal e As System.EventArgs)
   '' msgBox("Form_Activated")
   dim textbox=PanelRef.element("txt_url")
   dim  caption=textbox.text
   textbox.focus()
   '' textbox.text=caption+" xxx"
  
   ''msgBox(caption)
   '' zz.doevents()
   WB2.focus()
End Sub



Sub onIconEvent(e)
  FormRef.close
end Sub


'-->
'-->E V E N T S - 2 


Sub cmdGoBack(e)  ' ...
  WB2.goBack()
End Sub


Sub cmdGoForward(e)  ' ...
  WB2.goForward()
End Sub

Sub cmdGoHome(e)  ' ...
  WB2.Navigate("google.de")
End Sub


sub pic_expand_down_MouseClick(e)
  toggleUrlBar(true)
  '' e.sender.visible=false
  '' sidebar.Controls("pic_expand_up").visible=true
  '' toolbar.height=88
  '' onResizeControls()
end sub

sub pic_expand_up_MouseClick(e)
  toggleUrlBar(false)
  '' e.sender.visible=false
  '' sidebar.Controls("pic_expand_down").visible=true
  '' toolbar.height=0
  '' onResizeControls()
end sub



Sub cmdAction06(e)  ' ...
  dim el = splitB2 : el.Panel1Collapsed = not  el.Panel1Collapsed
End Sub


Sub cmdAction05(e)  ' ...
    popup.visible=false
End Sub


Sub cmdAction04(e)  ' ...
    dim el = splitA1: el.Panel2Collapsed = not  el.Panel2Collapsed
End Sub


Sub cmdAction03(e)  ' ...
    dim el = splitA1: el.Panel2Collapsed = not  el.Panel2Collapsed
End Sub



Sub cmdAction02(e)  ' ...
  toggleUrlBar()
End Sub

sub toggleUrlBar(optional force as integer=-1)
  trace1.text=force
  dim state as boolean =false
  dim height as integer=  toolbar.height
  if height<10 then state=true
  if force=0 then state=false
  if force=1 then state =true
  if state=true then
    sidebar.Controls("pic_expand_down").visible=false
    sidebar.Controls("pic_expand_up").visible=true
    toolbar.height=88
    onResizeControls()
  else
    sidebar.Controls("pic_expand_down").visible=true
    sidebar.Controls("pic_expand_up").visible=false
    toolbar.height=0
    onResizeControls()
 end if

end sub


Sub cmdAction01(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  FormRef.close
End Sub


Sub myPicture_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles myPicture.MouseDown
  if e.Button=Windows.Forms.MouseButtons.Right then
    'Toggle winSize
    if formRef.height >111 then
      lastHeight=formRef.height
      lastWidth=formRef.width
       formRef.height=33
       formRef.width=222
       formref.topmost=true
    else 
      formRef.height=lastHeight
      formRef.width=lastWidth
      formref.topmost=false
    end if
    exit Sub
  end if
  dim hwnd as integer
  Call ReleaseCapture()
  Call SendMessage(FormRef.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub




Sub myPicture2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles myPicture2.MouseDown
  FormRef.close
  exit sub
  
  msgbox("onIconEvent")
  
  if e.Button=Windows.Forms.MouseButtons.Right then
    'hide()
    exit Sub
  end if
  dim hwnd as integer
  Call ReleaseCapture()
  Call SendMessage(FormRef.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub



sub vbsuch_1_mouseClick(e)
  trace "vbsuch_1_", "url"
  dim url=e.sender.tag
  trace "url", url
  url=replace(url,"||SUCH||", toolbar.Controls("txt_phpsuch").Text)
  ''msgbox (url)
  WB2.Navigate(url)
End Sub



sub vbsuch_2_mouseClick(e)
  trace "vbsuch_1_", "url"
  dim url=e.sender.tag
  trace "url", url
  url=replace(url,"||SUCH||", toolbar.Controls("txt_phpsuch").Text)
  ''msgbox (url)
  WB2.Navigate(url)
End Sub


:Sub onTextboxEvent(e As Object)
    'TRACE e.Keystring, e.Sender.Name
    :If e.Keystring = "RETURN" And e.Sender.Name = "txt_url" Then
    :  WB2.Navigate(container1.Controls("txt_url").Text)
    :End If
    :If e.Keystring = "RETURN" And e.Sender.Name = "txt_phpsuch" Then
      :WB2.Navigate("http://www.php.net/"+panelRef.Controls("txt_phpsuch").Text)
    :End If
  End Sub
  
  ''Private Sub WB_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WB2.DocumentCompleted
  ''  panelRef.Controls("txt_url").Text = WB2.URL.ToString()
  ''End Sub

  '' public Sub WB_DocumentTitleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WB2.DocumentTitleChanged
  ''   ' panelRef.Text="miniBrowser" 'WB.DocumentTitle
  '' End Sub
  '' 
  '' public Sub WB_StatusTextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WB2.StatusTextChanged
  ''   'panelRef.Controls("txt_status").Text= WB2.StatusText
  '' End Sub
  '' 
  Sub Form_Resize(sender as Object, e as EventArgs) Handles formRef.Resize
   onResizeControls()
  End Sub

 




'-->
'--> wbEx  ...


Class wbEx
Inherits Skybound.Gecko.GeckoWebBrowser
Public Event beforeNavigate(ByVal para1 As String)

private g_myId as string = ""
public Property myId () As String
  Get 
    Return myId
  End Get
  Set (ByVal value As String)
    myId=value
  End Set
End Property

Sub wbEx_CreateWindow(ByVal sender As Object, ByVal e As Skybound.Gecko.GeckoCreateWindowEventArgs) Handles myBase.CreateWindow
  'Dim wb = frm_mdiHost.openNewTab("about:blank")
  'msgbox("wbEx_CreateWindow")
  'msgbox(e.WebBrowser.url.ToString())
  
  'globSkipNewWindow=true
  e.WebBrowser = WB2
End Sub
  
  
Sub WB_Navigating(ByVal sender As Object, ByVal e As Skybound.Gecko.GeckoNavigatingEventArgs) Handles myBase.Navigating
    'MsgBox(e.Uri)
    
    trace5.text= "WB_Navigating: "+e.Uri.ToString 
    myTextArea.text=myTextArea.text + e.Uri.ToString+vbNewLine
    myTextArea.Select(myTextArea.Text.Length - 1, 0)
    myTextArea.ScrollToCaret()
    ''e.Cancel = True
    'trace "check1.selected" ,check1.checked
    'trace "GetType" , e.GetType().toString
    
    ''msgBox(check1.checked)
    ''msgBox(name)
    if check1.checked=true and name="sidebar" then
      dim url as string=e.Uri.ToString
      if url.contains("?view=")= true then exit sub
      ''msgBox(e.Uri.ToString)
      WB2.Navigate(e.Uri.ToString)
      e.Cancel = True
    end if
      
    if globSkipNewWindow=true then
      globSkipNewWindow=false
      e.Cancel = True
      '' msgBox(e.Uri.ToString)
    end if
      
    'if check1.checked=true then e.Cancel = True
    
End Sub

 Private Sub wbEx_DomMouseMove(ByVal sender As Object, ByVal e As Skybound.Gecko.GeckoDomMouseEventArgs) Handles myBase.DomMouseMove
    exit sub
    
    'Debug.Print("DomMouseMove")
    'Debug.Print(e.ClientY.ToString)
    'Debug.Print(e.Button.ToString)
    'trace "MouseMove: ", e.ClientX.ToString 
    e.Handled = True
    'e.Target.SetAttribute("style", "background-color: #f00; ")
    ''e.Target.SetAttribute("background-color", "#f00")
    ''dim color as string=e.Target.GetAttribute("background-color")
    'trace "color" , e.Target.GetAttribute("style")
  End Sub


Sub wbEx_DomKeyDown(ByVal sender As Object, ByVal e As Skybound.Gecko.GeckoDomKeyEventArgs) Handles myBase.DomKeyDown
     'Debug.Print(e.Target.InnerHtml)
     'msgBox(e.KeyCode)
     'e.Handled=true
    ''e.Handled=true
End Sub

''Sub wbEx_DomMouseOver(ByVal sender As Object, ByVal e As Skybound.Gecko.GeckoDomMouseEventArgs) Handles GeckoWebBrowser1.DomMouseOver
Sub wbEx_DomMouseOver(ByVal sender As Object, ByVal e As Skybound.Gecko.GeckoDomMouseEventArgs) Handles myBase.DomMouseOver
    exit sub
    
    dim tag as string=e.Target.TagName
    trace "tag", tag
    'Exit Sub
    trace "style" , e.Target.GetAttribute("style")
    'e.Target.SetAttribute("style", "background-color: #f00; ")
    ''e.Target.SetAttribute("background-color", "#f00")
    ''dim color as string=e.Target.GetAttribute("background-color")
    trace "style" , e.Target.GetAttribute("style")

End Sub


Sub wbEx_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles myBase.KeyDown
  '...der scheint nicht zu kommen
  printLine 2, "wbEx_KeyDown", e.keyCode
  '' msgBox(e.keyCode.toString)
End Sub






Sub wbEx_DocumentTitleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.DocumentTitleChanged
  meScriptClass.panelRef.Text="miniBrowser" 'WB.DocumentTitle
End Sub

Sub wbEx_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles myBase.DocumentCompleted
  meScriptClass.toolbar.Controls("txt_url").Text = me.URL.ToString()
End Sub




Sub WB_StatusTextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.StatusTextChanged
  ' !!! zu viele event !!! meScriptClass.statbar.Controls("txt_status").Text= me.StatusText
  'trace "txt_status", me.StatusText
End Sub
  
  
  




End Class











'-->
'-->M O U S E - H O O K 

Friend WithEvents KeyboardHook1 As WindowsHook.KeyboardHook
Friend WithEvents MouseHook1 As new WindowsHook.MouseHook
Friend WithEvents ClipboardHook1 As WindowsHook.ClipboardHook
public globMouseHookStartPosX as integer
public globMouseHookContainerPosX as integer
public globMouseHookStartPosY as integer
public globMouseHookContainerPosY as integer
public globMouseHookDragMode as boolean=false
public globMouseHookTimerStart as object
public globCheckForShowPopup as boolean=false



sub toggleSidebar(deltaX as integer)
  'msgBox(deltaX.toString)
  if deltax < 0 then splitB2.Panel1Collapsed =true
  if deltax > 0 then splitB2.Panel1Collapsed =false
end sub


sub mouseHook_toggleUrlBar(deltax as integer)
  'msgBox(deltaX.toString)
  if globMouseHookDragMode=true then
    if deltax < 0 then 
       toggleUrlBar(1)
       globMouseHookDragMode=false
    end if
    if deltax > 0 then 
      toggleUrlBar(0)
      globMouseHookDragMode=false
    end if
  end if
end sub



sub checkForShowPopupWindow(e)
  trace5.text=GetTime()
  if globCheckForShowPopup=false then
    stopTimer()
    exit sub
  end if
  if e.Button=Windows.Forms.MouseButtons.Right then
    if globMouseHookDragMode=true then 
       dim deltaTime=GetTime()-globMouseHookTimerStart
       'trace4.text=deltaTime.toString("00.000")
       if deltaTime>1333 then
          popup.visible=true
          popup.bringToFront
          stopTimer()
       end if
''  
''       dim curPosX as integer= e.Location.x
''       dim newPosX=curPosX-globMouseHookStartPosX
''       container1.left=globMouseHookContainerPosX+newPosX
''       dim curPosY as integer= e.Location.Y
''       dim newPosY=curPosY-globMouseHookStartPosY
''       container1.top=globMouseHookContainerPosY+newPosY
''       
''       'printLine 13, "HOOK:", e.Location.ToString
      'printLine 14, "MouseDown:", e.Location.x.ToString
    end if
  else
    'globMouseHookDragMode=false
  end if
 

end sub

Private Sub CleareDisplay()
  '' Me.DisplayTextBox.Clear()
  '' 'Print WindowsHooksLib information
  '' Me.DisplayTextBox.AppendText("Company: " & WindowsHook.MouseHook.Info.CompanyName)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Assembly Version: " & WindowsHook.MouseHook.Info.Version.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Assembly File Version: " & Diagnostics.FileVersionInfo.GetVersionInfo( _
  ''                              WindowsHook.MouseHook.Info.DirectoryPath & "\" _
  ''                              & WindowsHook.MouseHook.Info.AssemblyName & ".dll").FileVersion)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Assembly Name: " & WindowsHook.MouseHook.Info.AssemblyName)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
End Sub


'**************** Low Level Keyboard Hook ****************

'' Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
''   Try
''       'Install the keyboard hook
''       Me.KeyboardHook1.InstallHook()
''   Catch ex As Exception
''       MessageBox.Show(ex.Message)
''       WindowsHook.ErrorLog.ExceptionToFile(ex, TraceEventType.Error)
''   End Try
'' End Sub
'' 
'' Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
''   Try
''       'Remove the mouse hook
''       Me.KeyboardHook1.RemoveHook()
''   Catch ex As Exception
''       MessageBox.Show(ex.Message)
''       WindowsHook.ErrorLog.ExceptionToFile(ex, TraceEventType.Error)
''   End Try
'' End Sub
'' 


Private Sub KeyboardHook1_KeyDown(ByVal sender As Object, ByVal e As WindowsHook.KeyboardEventArgs) Handles KeyboardHook1.KeyDown
  'Set the key down Handled property
  
  ''e.Handled = Me.HandleKeyboardCheckBox.Checked
  
  'Print the key down data
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== KeyDown")
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Handled: " & e.Handled)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("KeyCode: " & e.KeyCode.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("KeyValue: " & e.KeyValue)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("KeyData: " & e.KeyData.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Alt: " & e.Alt)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Control: " & e.Control)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Shift: " & e.Shift)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Modifiers: " & e.Modifiers.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== End KeyDown")
  '' 
End Sub

Private Sub KeyboardHook1_KeyUp(ByVal sender As Object, ByVal e As WindowsHook.KeyboardEventArgs) Handles KeyboardHook1.KeyUp

  'Set the key up Handled property
  
  ''e.Handled = Me.HandleKeyboardCheckBox.Checked
  
  'Print the key up data
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== KeyUp")
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Handled: " & e.Handled)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("KeyCode: " & e.KeyCode.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("KeyValue: " & e.KeyValue)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("KeyData: " & e.KeyData.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Alt: " & e.Alt)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Control: " & e.Control)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Shift: " & e.Shift)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Modifiers: " & e.Modifiers.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== End KeyUp")
  '' 
End Sub

Private Sub KeyboardHook1_StateChanged(ByVal sender As Object, ByVal e As WindowsHook.StateChangedEventArgs) Handles KeyboardHook1.StateChanged
  'Print the keyboard hook state in the KeyboardGroupBox name 
  'Me.KeyboardGroupBox.Text = "Keyboard Hook " & e.State.ToString & "!"
  If e.State = WindowsHook.HookState.Uninstalled Then
      'Clear the Handle Keyboard checkbox
      'Me.HandleKeyboardCheckBox.Checked = False
  End If
End Sub

'*******************************************************
'*********** End Low Level Keyboard Hook ***************
'*******************************************************


'**************** Low Level Mouse Hook *****************

Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)'' Handles Button3.Click
  '' Try
  ''     'Install the mouse hook
  ''     Me.MouseHook1.InstallHook()
  '' Catch ex As Exception
  ''     MessageBox.Show(ex.Message)
  ''     WindowsHook.ErrorLog.ExceptionToFile(ex, TraceEventType.Error)
  '' End Try
End Sub

Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) ''Handles Button4.Click
  '' Try
  ''     'Remove the mouse hook
  ''     Me.MouseHook1.RemoveHook()
  '' Catch ex As Exception
  ''     MessageBox.Show(ex.Message)
  ''     WindowsHook.ErrorLog.ExceptionToFile(ex, TraceEventType.Error)
  '' End Try
End Sub

Private Sub MouseHook1_MouseDown(ByVal sender As Object, ByVal e As WindowsHook.MouseEventArgs) Handles MouseHook1.MouseDown
  ''msgBox("MouseDown")
  if formref.handle <> getActiveWindow() then exit sub
  'e.Handled=true

  printLine 16, Windows.Forms.MouseButtons.Right.toString, e.Button.ToString
     printLine 14, "MouseDownLEFT:", e.Location.x.ToString

     dim x as integer=formRef.left+popup.left
     trace1.text=x.toString()
     dim y as integer=formRef.top+popup.top
     trace3.text=y.toString()
     trace2.text=e.Location.x.toString()
     trace4.text=e.Location.Y.toString()

     if e.Location.x<x then popup.visible=false
     if e.Location.y<y then popup.visible=false
     if e.Location.x>x +popup.width then popup.visible=false
     if e.Location.y>y +popup.height then popup.visible=false


  if e.Button=Windows.Forms.MouseButtons.Right then
     printLine 14, "MouseDown:", e.Location.x.ToString
     globMouseHookTimerStart = GetTime()
     globCheckForShowPopup=true
     startTimer()
     
     globMouseHookStartPosX=e.Location.x
     globMouseHookStartPosY=e.Location.y
     globMouseHookContainerPosX =container1.left
     globMouseHookContainerPosY =container1.top

     globMouseHookDragMode=true
     '' popup.visible=true
     '' popup.bringToFront

     'e.Handled=true
     ' Call ReleaseCapture()
     ' Call SendMessage(FormRef.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
     
     'container1.left=container1.left+10
  else
     'container1.left=container1.left-10
  end if 
  ''end if
  'Set the mouse down Handled property
  
  '' e.Handled = Me.HandleMouseCheckBox.Checked And Not Me.mList.Contains(CType(sender, IntPtr))
  
  
  'Print the mouse down data
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== MouseDown")
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Handled: " & e.Handled)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Button: " & e.Button.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Clicks: " & e.Clicks)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Delta: " & e.Delta)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Location: " & e.Location.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Control Handle: " & sender.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== End MouseDown")
  '' 
End Sub

Private Sub MouseHook1_MouseUp(ByVal sender As Object, ByVal e As WindowsHook.MouseEventArgs) Handles MouseHook1.MouseUp
  if formref.handle <> getActiveWindow() then exit sub
  'Set the mouse up Handled property
  if e.Button=Windows.Forms.MouseButtons.Right then
    if globMouseHookDragMode=true then 
       dim deltaTime=GetTime()-globMouseHookTimerStart
       'trace4.text=deltaTime.toString("00.000")
       if deltaTime<222 then
          dim curPosX as integer=  e.Location.X -formRef.Left
          dim curPosY as integer=  e.Location.Y -formRef.top
          ' msgbox(curPosY)
          if curPosY>50
            'e.Handled=true
            popup.left=curPosX-200
            popup.top=curPosY-40
            popup.visible=true
            popup.bringToFront
            stopTimer()
          end if
       end if
     end if
   end if
  stopTimer()
  '' if e.Button=Windows.Forms.MouseButtons.Right then
  ''      printLine 15, "MouseUp:", e.Location.ToString
  ''      dim deltaTime=GetTime()-globMouseHookTimerStart
  ''      trace4.text=deltaTime.toString("00,000")
  ''      if deltaTime>777 then
  ''         popup.visible=true
  ''         popup.bringToFront
  ''      end if
  ''    'e.Handled=true
  '' end if
  '' 
  '' onResizeControls()
  '' 
  
  ''e.Handled = Me.HandleMouseCheckBox.Checked And Not Me.mList.Contains(CType(sender, IntPtr))
  
  'Print the mouse up data
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== MouseUp")
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Handled: " & e.Handled)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Button: " & e.Button.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Clicks: " & e.Clicks)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Delta: " & e.Delta)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Location: " & e.Location.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Control Handle: " & sender.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== End MouseUp")
  '' 
End Sub

Private Sub MouseHook1_MouseClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MouseHook1.MouseClick

  'Print the mouse click data
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== MouseClick")
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Button: " & e.Button.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Clicks: " & e.Clicks)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Delta: " & e.Delta)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Location: " & e.Location.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Control Handle: " & sender.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== End MouseClick")
  '' 
End Sub

Private Sub MouseHook1_MouseDoubleClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MouseHook1.MouseDoubleClick

  'Print the mouse double click data
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== MouseDoubleClick")
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Button: " & e.Button.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Clicks: " & e.Clicks)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Delta: " & e.Delta)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Location: " & e.Location.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Control Handle: " & sender.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== End MouseDoubleClick")
  '' 
End Sub

Private Sub MouseHook1_MouseWheel(ByVal sender As Object, ByVal e As WindowsHook.MouseEventArgs) Handles MouseHook1.MouseWheel

  'Set the mouse wheel Handled property
  
  ''e.Handled = Me.HandleMouseCheckBox.Checked
  
  '' 'Print the mouse wheel data
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== MouseWheel")
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Handled: " & e.Handled)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Button: " & e.Button.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Clicks: " & e.Clicks)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Delta: " & e.Delta)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("Location: " & e.Location.ToString)
  '' Me.DisplayTextBox.AppendText(System.Environment.NewLine)
  '' Me.DisplayTextBox.AppendText("===================== End MouseWheel")
  '' 
End Sub

:Private Sub MouseHook1_MouseMove(ByVal sender As Object, ByVal e As WindowsHook.MouseEventArgs) Handles MouseHook1.MouseMove
   
  if formref.handle <> getActiveWindow() then exit sub
     'trace2.text=e.Location.x.toString()
     'trace4.text=e.Location.Y.toString()

  if e.Button=Windows.Forms.MouseButtons.Right then
    if globMouseHookDragMode=true then 
    
      dim deltaTime=GetTime()-globMouseHookTimerStart
      'trace4.text=deltaTime.toString("00.000")
      '' if deltaTime>777 000 then
      ''    popup.visible=true
      ''    popup.bringToFront
      '' end if
        
      dim curPosX as integer= e.Location.x
      dim newPosX=curPosX-globMouseHookStartPosX
      ' container1.left=globMouseHookContainerPosX+newPosX
      dim curPosY as integer= e.Location.Y
      dim newPosY=curPosY-globMouseHookStartPosY
      
      'container1.top=globMouseHookContainerPosY+newPosY
      dim deltaX =globMouseHookContainerPosX+newPosX
      dim deltaY =globMouseHookContainerPosY+newPosY
      
      if deltaX>55 or deltaX< -55 then globCheckForShowPopup=false
      if deltaX>150 or deltaX<-150 then toggleSidebar(deltax)
      
      if deltaY>55 or deltaY< -55 then globCheckForShowPopup=false
      if deltaY>150 or deltaY<-150 then mouseHook_toggleUrlBar(deltaY *-1)
      
      'trace4.text=deltaX.tostring
      
      'printLine 13, "HOOK:", e.Location.ToString
      'printLine 14, "MouseDown:", e.Location.x.ToString
    end if
  else
    globMouseHookDragMode=false
  end if
  
    ''Print the mouse move data
  'Me.DisplayTextBox.AppendText(Environment.NewLine)
  'Me.DisplayTextBox.AppendText("===================== MouseMove")
  'Me.DisplayTextBox.AppendText(Environment.NewLine)
  'Me.DisplayTextBox.AppendText("Button: " & e.Button.ToString)
  'Me.DisplayTextBox.AppendText(Environment.NewLine)
  'Me.DisplayTextBox.AppendText("Clicks: " & e.Clicks)
  'Me.DisplayTextBox.AppendText(Environment.NewLine)
  'Me.DisplayTextBox.AppendText("Delta: " & e.Delta)
  'Me.DisplayTextBox.AppendText(Environment.NewLine)
  'Me.DisplayTextBox.AppendText("Location: " & e.Location.ToString)
  'Me.DisplayTextBox.AppendText(Environment.NewLine)
  'Me.DisplayTextBox.AppendText("Control Handle: " & sender.ToString)
  'Me.DisplayTextBox.AppendText(Environment.NewLine)
  'Me.DisplayTextBox.AppendText("===================== End MouseMove")

End Sub

Private Sub MouseHook1_StateChanged(ByVal sender As Object, ByVal e As WindowsHook.StateChangedEventArgs) Handles MouseHook1.StateChanged
  'Print the mouse hook state in the MouseGroupBox name 
''  Me.MouseGroupBox.Text = "Mouse Hook " & e.State.ToString & "!"
  If e.State = WindowsHook.HookState.Uninstalled Then
      '' 'Clear the mouse location and the Handle Mouse checkbox
      '' Me.MouseLocationTextBox.Text = String.Empty
      '' Me.HandleMouseCheckBox.Checked = False
  End If
End Sub

'*******************************************************
'*************** End Low Level Mouse Hook **************
'*******************************************************


'**************** simulate Mouse Event *****************
'**************** Synthesize Mouse Event *****************

'' Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
''   '' If (e.Modifiers And Keys.Shift) = Keys.Shift And e.KeyCode = Keys.C Then
''   ''     Dim p As Point = Me.GroupBox1.PointToScreen(Me.ClickCheckBox.Location)
''   ''     p.X += 6
''   ''     p.Y += 10
''   ''     WindowsHook.MouseHook.SynthesizeMouseMove(p, WindowsHook.MapOn.PrimaryMonitor, IntPtr.Zero)
''   ''     Me.ClickCheckBox.Focus()
''   ''     WindowsHook.MouseHook.SynthesizeMouseDown(Windows.Forms.MouseButtons.Left, IntPtr.Zero)
''   ''     WindowsHook.MouseHook.SynthesizeMouseUp(Windows.Forms.MouseButtons.Left, IntPtr.Zero)
''   '' End If
'' End Sub
'' 
'*******************************************************
'************** End Synthesize Mouse Event *************
'*******************************************************


'**************** Window Clipboard Hook ****************

'' Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
''   Try
''       'Install the clipboard hook
''       Me.ClipboardHook1.InstallHook(Me)
''   Catch ex As Exception
''       MessageBox.Show(ex.Message)
''       WindowsHook.ErrorLog.ExceptionToFile(ex, TraceEventType.Error)
''   End Try
'' End Sub
'' 
'' Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
''   Try
''       'Remove the clipboard hook
''       Me.ClipboardHook1.RemoveHook()
''   Catch ex As Exception
''       MessageBox.Show(ex.Message)
''       WindowsHook.ErrorLog.ExceptionToFile(ex, TraceEventType.Error)
''   End Try
'' End Sub
'' 
'' Private Sub ClipboardHook1_ClipboardChanged(ByVal sender As Object, ByVal e As WindowsHook.ClipboardEventArgs) Handles ClipboardHook1.ClipboardChanged
''   'Print the mouse wheel data
''   Me.DisplayTextBox.AppendText(System.Environment.NewLine)
''   Me.DisplayTextBox.AppendText("===================== ClipboardChanged")
''   Me.DisplayTextBox.AppendText(System.Environment.NewLine)
''   Me.DisplayTextBox.AppendText("Hooked Window: " & e.HookedWindow.ToString)
''   Me.DisplayTextBox.AppendText(System.Environment.NewLine)
''   Me.DisplayTextBox.AppendText("Source Window : " & e.SourceWindow.ToString)
''   Me.DisplayTextBox.AppendText(System.Environment.NewLine)
''   Me.DisplayTextBox.AppendText("Contains Audio: " & My.Computer.Clipboard.ContainsAudio)
''   Me.DisplayTextBox.AppendText(System.Environment.NewLine)
''   Me.DisplayTextBox.AppendText("Contains Image: " & My.Computer.Clipboard.ContainsImage)
''   Me.DisplayTextBox.AppendText(System.Environment.NewLine)
''   Me.DisplayTextBox.AppendText("Contains Text: " & My.Computer.Clipboard.ContainsText)
''   Me.DisplayTextBox.AppendText(System.Environment.NewLine)
''   Me.DisplayTextBox.AppendText("Contains FileDropList: " & My.Computer.Clipboard.ContainsFileDropList)
''   Me.DisplayTextBox.AppendText(System.Environment.NewLine)
''   Me.DisplayTextBox.AppendText("===================== End ClipboardChanged")
'' End Sub
'' 
'' 
'' Private Sub ClipboardHook1_StateChanged(ByVal sender As Object, ByVal e As WindowsHook.StateChangedEventArgs) Handles ClipboardHook1.StateChanged
''   'Print the mouse hook state in the MouseGroupBox name 
''   Me.Clipboard_GroupBox.Text = "Clipboard Hook " & e.State.ToString & "!"
'' End Sub
'' 
'' '*******************************************************
'' '************** End Window Clipboard Hook *************
'' '*******************************************************
'' 

'******************* Form's Events *********************

'' Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
''   Me.CleareDisplay()
''   Me.MouseGroupBox.Text = "Mouse Hook " & Me.MouseHook1.State.ToString & "!"
''   Me.KeyboardGroupBox.Text = "Keyboard Hook " & Me.KeyboardHook1.State.ToString & "!"
''   Me.Clipboard_GroupBox.Text = "Clipboard Hook " & Me.ClipboardHook1.State.ToString & "!"
''   'Set the alowed controls list
''   For Each ctr As Control In Me.MouseGroupBox.Controls
''       Me.mList.Add(ctr.Handle)
''   Next
''   Me.Label2.Text = "My handle is: " & Me.ClickCheckBox.Handle.ToString
'' End Sub
'' 
'' Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
''   Me.KeyboardHook1.Dispose()
''   Me.MouseHook1.Dispose()
''   Me.ClipboardHook1.Dispose()
'' End Sub
'' 











'-->
'--> K e y b o a r d H o o k 





Public Class KeyboardHook
    ''Constants
Private Const HC_ACTION As Integer = 0
Private Const WH_KEYBOARD_LL As Integer = 13
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105

''Keypress Structure
Private Structure KBDLLHOOKSTRUCT
    Public vkCode As Integer
    Public scancode As Integer
    Public flags As Integer
    Public time As Integer
    Public dwExtraInfo As Integer
End Structure



''--> API Functions


Private Declare Function SetWindowsHookEx Lib "user32" _
Alias "SetWindowsHookExA" _
(ByVal idHook As Integer, _
ByVal lpfn As KeyboardProcDelegate, _
ByVal hmod As Integer, _
ByVal dwThreadId As Integer) As Integer

Private Declare Function CallNextHookEx Lib "user32" _
(ByVal hHook As Integer, _
ByVal nCode As Integer, _
ByVal wParam As Integer, _
ByVal lParam As KBDLLHOOKSTRUCT) As Integer

Private Declare Function UnhookWindowsHookEx Lib "user32" _
(ByVal hHook As Integer) As Integer

''Our Keyboard Delegate
Private Delegate Function KeyboardProcDelegate _
(ByVal nCode As Integer, _
ByVal wParam As Integer, _
ByRef lParam As KBDLLHOOKSTRUCT) As Integer

''The KeyPress events
Public Shared Event KeyDown(ByVal Key As Keys,byRef cancel as boolean)
Public Shared Event KeyUp(ByVal Key As Keys)
''The identifyer for our KeyHook
Private Shared KeyHook As Integer
''KeyHookDelegate
Private Shared KeyHookDelegate As KeyboardProcDelegate

Public Sub New()
    ''Installs a Low Level Keyboard Hook
    KeyHookDelegate = New KeyboardProcDelegate(AddressOf KeyboardProc)
    KeyHook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyHookDelegate, System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32, 0)
End Sub

Private Shared Function KeyboardProc(ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As KBDLLHOOKSTRUCT) As Integer
    ''If it is a keypress
    dim cancel as boolean =false
    If (nCode = HC_ACTION) Then
        Select Case wParam
            ''If it is a Keydown Event
            Case WM_KEYDOWN, WM_SYSKEYDOWN
                ''Activates the KeyDown event in Form 1
                RaiseEvent KeyDown(CType(lParam.vkCode, Keys),cancel)
            Case WM_KEYUP, WM_SYSKEYUP
                ''Activates the KeyUp event in Form 1
                RaiseEvent KeyUp(CType(lParam.vkCode, Keys))
        End Select
    End If
    ''Next
    if cancel=false then 
      Return CallNextHookEx(KeyHook, nCode, wParam, lParam)
    else
      return 1
    end if
End Function

Protected Overrides Sub Finalize()
    ''On close it UnHooks the Hook
    UnhookWindowsHookEx(KeyHook)
    MyBase.Finalize()
End Sub
    
    
End Class



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
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
    Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Integer) As Integer

    ' Win32 APIs used to automate drag and sysmenu support.
    Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal revert As Integer) As Integer
    Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
    Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Integer) As Integer
    Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Integer, ByVal un As Integer, ByVal b As Integer, ByVal lpMenuItemInfo As MENUITEMINFO) As Integer
    Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Integer, ByVal un As Integer, ByVal bool As Integer, ByVal lpcMenuItemInfo As MENUITEMINFO) As Integer
    Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Declare Function ReleaseCapture Lib "user32" () As Integer
    Private Declare Function GetCursorPos Lib "user32" (ByVal lpPoint As POINTAPI) As Integer

    ' Used to get menu information.
    public Structure MENUITEMINFO
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
      If Button = System.Windows.Forms.MouseButtons.Left Then
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
      If Button = System.Windows.Forms.MouseButtons.Right Then
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

    '' Private Function GetMenuItemText(ByVal hMenu As Integer, ByVal nPosition As Integer) As String
    ''   Dim mii As MENUITEMINFO
    '' 
    ''   ' Initialize structure.
    ''   mii.cbSize = Len(mii)
    ''   mii.fMask = MIIM_TYPE
    ''   mii.dwTypeData = StrDup(80, vbNullChar)
    ''   mii.cch = Len(mii.dwTypeData)
    ''   Call GetMenuItemInfo(hMenu, nPosition, MF_BYPOSITION, mii)
    '' 
    ''   ' Return current menu text
    ''   If mii.cch > 0 Then
    ''     'GetMenuItemText = Left$(mii.dwTypeData, mii.cch)
    ''     GetMenuItemText = Left(mii.dwTypeData, mii.cch)
    ''   End If
    '' End Function
    '' 
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
'--> outMONITOR

 sub clearAll()
   on error resume next
   monitorClear()
   zz.traceClear()
   zz.printLineReset()
   err.number=0
end sub



sub setMonitorRef()
  'on error resume next
  globMonitorRef = zz.scriptClass("es_testLabor")
  trace "globMonitorRef", globMonitorRef.name
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
   globMonitorRef.statustext(para1+"<--")
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



















